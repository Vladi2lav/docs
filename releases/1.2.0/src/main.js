// Точка входа: связывает загрузку документов → расчёт → предпросмотр → Word.
// Вся предметная логика лежит в parsers/, domain/, export/; здесь — состояние и обработчики интерфейса.

import { readSheets } from './util/sheet.js';
import { parseGraph, guessStartYear, groupKeys, normGroup } from './parsers/graph.js';
import { HOURS_PER_PAIR, PRACTICE_PREFIX } from './config.js';
import { parseSchedule } from './parsers/schedule.js';
import { parsePractice } from './parsers/practice.js';
import { parsePrevSemester } from './parsers/prev-semester.js';
import { parseLoad, normIndex } from './parsers/load.js';
import { applyLoad, loadCandidates, loadLabel, matchLoad, sameGroup, sameTitle } from './domain/load.js';
import { bellTable, buildGrid, mergeSchedules, parseBells, weeklySummary } from './domain/schedules.js';
import { describe as describePractice, expandBlock, totalHours } from './domain/practice-days.js';
import { buildPracticeDocx, shortSigner } from './export/practice-docx.js';
import { makeCombo } from './ui/combo.js';
import { renderGrid } from './ui/schedule-grid.js';
import { openPracticeDialog } from './ui/practice-dialog.js';
import { initAssistant } from './ui/assistant.js';
import { applyPrevSemester } from './domain/carryover.js';
import { annualProblems, buildAnnualEntries, genitiveCollege, groupByLoad } from './domain/annual.js';
import { buildAnnualDocx } from './export/annual-docx.js';
import { LOAD_TYPES } from './config.js';
import { buildRows, emptyPracticeRow, graphPeriods, matchSupervisor, teacherNames } from './domain/rows.js';
import { createCalendar } from './domain/calendar.js';
import { buildTimesheet } from './domain/timesheet.js';
import { buildDocx } from './export/docx.js';
import { esc, renderSheet } from './ui/preview.js';
import { renderGraph, highlightSet } from './ui/graph-view.js';
import { renderCalendar } from './ui/calendar-view.js';
import { ruDate, weekday, ymIndex, ymOf } from './util/dates.js';
import { MONTH_NAMES, FUNDING } from './config.js';
import { initReleases, initVersion } from './ui/version.js';
import { deleteWorkspace, getWorkspace, listWorkspaces, putWorkspace, workspaceFromJson, workspaceToJson } from './util/workspaces.js';

const $ = id => document.getElementById(id);
const S = {
  graphSheet: null, graph: null, holidays: { fixed: [], dates: [] },   // заменяется данными из data/holidays-kz.json
  schedules: [],                 // { id, file, sheet, buffer, parsed } — по одному на курс (лист книги)
  schedule: null,                // все расписания вместе: { entries, groups, semester, hasSaturday, warnings }
  load: null,                    // нагрузка: { rows }
  signers: { approver: '', upr: '', head: '' },     // подписанты документа практики
  plan: { groups: null, onlyMine: false, showEmpty: false, bells: '' },   // вкладка «Расписание и практики»
  practices: [],                 // { id, file, data, teacher }
  teacher: '', rows: [], warnings: [],
  transfers: {},                 // педагог → [{ id, rowId, from, to, hours }]
  cal: { overrides: {}, transfers: [], workSaturday: null, swapWeeks: false },   // ручные правки календаря (до обновления страницы)
  result: null, activeMonth: 0,
  files: { graph: null, load: null, prev: null },
  prev: null,                    // ведомость 1 семестра: ФИО, специальность, строки (план, выполнено, часы по месяцам)
  extraLoads: {},                // тип нагрузки для строк, которые были только в 1 семестре
  annual: { fullGen: '', collegeGen: '', notes: {} },
  entries: [],   // исходные файлы (имя + содержимое) — нужны для сохранения рабочей области
  teacherStates: {},             // педагог → { rows, warnings, fields, semester }: правки, сделанные для каждого педагога
  calStart: null,                // первый месяц, показанный в календаре ({ y, m })
  funding: 'grant',              // учёт ведётся отдельно: грантовые или коммерческие группы
  fundingOverrides: {},          // группа → 'grant' | 'comm', если цвет в графике неверен
  wsId: null, wsName: '', dirty: false, loading: false, savedAt: null,
};
const FIELD_IDS = ['fullName', 'specialty', 'dept', 'cmk', 'position'];
let specialtyCombo = null;
let groupsCombo = null;

const SOURCES_HTML = `<div class="files">
  <label class="file"><b>График учебного процесса</b> <small>.xlsx</small>
    <input type="file" data-src="graph" accept=".xlsx,.xls"><em data-info="graph">не загружен</em></label>
  <div class="file" data-box="schedule"><b>Расписания занятий</b> <small>.xlsx — по одному на курс, можно несколько</small>
    <label class="filebtn">+ Добавить расписание<input type="file" data-src="schedule" accept=".xlsx,.xls" multiple hidden></label>
    <ul class="list slim" data-list="schedule"></ul><em data-info="schedule">не загружено</em></div>
  <label class="file"><b>Нагрузка</b> <small>.xlsx, необязательно</small>
    <input type="file" data-src="load" accept=".xlsx,.xls"><em data-info="load">не загружена</em></label>
  <label class="file"><b>Расписание практики</b> <small>.docx, необязательно</small>
    <input type="file" data-src="practice" accept=".docx" multiple><em>&nbsp;</em></label></div>`;

init();

async function init() {
  mountSources();
  resetForm();
  mountCombos();

  document.querySelectorAll('#nav button').forEach(b => { b.onclick = () => showTab(b.dataset.tab); });
  $('year').onchange = reparseGraph;
  $('teacher').onchange = e => selectTeacher(e.target.value);
  $('planTeacher').onchange = e => selectTeacher(e.target.value);
  initPlanTab();
  $('semester').onchange = () => { applyLoadToRows(); renderRows(); recompute(); };
  $('fullName').oninput = () => { $('fullName').classList.remove('bad'); refreshPreview(); };
  for (const id of ['specialty', 'college', 'deputy', 'dept', 'cmk', 'position']) $(id).oninput = refreshPreview;
  $('specialty').addEventListener('input', refreshPreview);
  $('fillAll').onclick = fillAll;
  $('prAdd').onclick = addPractice;
  $('carry').onchange = recompute;
  $('prevFile').onchange = e => loadPrevFile(e.target.files[0]);
  $('fullGen').oninput = () => { S.annual.fullGen = $('fullGen').value; renderAnnual(); };
  $('collegeGen').oninput = () => { S.annual.collegeGen = $('collegeGen').value; renderAnnual(); };
  $('annualDownload').onclick = downloadAnnual;
  $('swapWeeks').onclick = () => { S.cal.swapWeeks = !S.cal.swapWeeks;  recompute(); };
  $('download').onclick = download;
  $('trAdd').onclick = addTransfer;
  $('graphQuery').oninput = $('graphMine').onchange = renderGraphTab;
  $('workSat').onchange = () => { S.cal.workSaturday = $('workSat').checked;  afterCalendarChanged(); };
  $('dtAdd').onclick = addDayTransfer;
  renderDayTransfers();
  S.holidays = await (await fetch('data/holidays-kz.json')).json();   // обработчики уже подключены; календарь нужен только для расчёта
  initVersion({ label: $('appVersion'), banner: $('updateBanner'), button: $('updateBtn') });
  initReleases({ select: $('releaseSelect'), badge: $('channel') });
  $('fundSwitch').querySelectorAll('button').forEach(bt => { bt.onclick = () => setFunding(bt.dataset.fund); });
  initWorkspaces();
  initAssistant({ button: $('assistantBtn'), panel: $('assistantPanel'), getContext: assistantContext });
}

// После обновления страницы всё начинается с чистого листа (браузер иначе подставляет прежние значения полей)
function resetForm() {
  document.querySelectorAll('input, select, textarea').forEach(el => {
    if (el.type === 'file') el.value = '';
    else if (el.type === 'checkbox') el.checked = el.defaultChecked;
    else if (el.tagName === 'TEXTAREA') el.value = '';
    else if (el.tagName === 'SELECT') el.selectedIndex = 0;
    else el.value = el.defaultValue;
  });
}

function showTab(name) {
  document.querySelectorAll('.tab').forEach(t => { t.hidden = t.id !== `tab-${name}`; });
  document.querySelectorAll('#nav button').forEach(b => b.classList.toggle('on', b.dataset.tab === name));
  if (name === 'plan') { renderPlan(); renderPracticeBlocks(); }
  if (name === 'graph') renderGraphTab();
  if (name === 'calendar') renderCalendarTab();
}

// ───────── загрузка документов ─────────

async function loadGraphFile(file) {
  if (file) await loadGraphBuffer(file.name, await file.arrayBuffer());
}

async function loadGraphBuffer(name, buffer) {
  try {
    S.graphSheet = pickGraphSheet(buffer);
    S.files.graph = { name, buffer };
    const guessed = guessStartYear(name);
    if (guessed && !S.loading) $('year').value = guessed;
    reparseGraph();
  } catch (e) { fail('graph', e); }
}

// В книге из нескольких листов график — тот лист, где нашлась его таблица
function pickGraphSheet(buffer) {
  const sheets = readSheets(XLSX, buffer);
  if (sheets.length === 1) return sheets[0].sheet;
  const hit = sheets.find(({ sheet }) => { try { parseGraph(sheet, { startYear: 2000 }); return true; } catch { return false; } });
  return (hit || sheets[0]).sheet;
}

function reparseGraph() {
  if (!S.graphSheet) return;
  try {
    S.graph = parseGraph(S.graphSheet, { startYear: Number($('year').value) });
    const g = S.graph;
    setInfo('graph', `${S.files.graph.name}: ${g.weeks.length} нед., ${Object.keys(g.groups).length} групп`, true);
    if (g.college) $('college').value = g.college;
    if (g.deputy) $('deputy').value = g.deputy;
    updateSpecialtyOptions();
    rebuildSchedule();
  } catch (e) { S.graph = null; fail('graph', e); }
}

// Расписания: несколько файлов (по курсам), а в одной книге — несколько листов
async function addScheduleFiles(files) {
  for (const file of files) {
    try { addScheduleBuffer(file.name, await file.arrayBuffer()); } catch (e) { showMessage(e.message, true); }
  }
  rebuildSchedule();
}

const sheetsCache = new WeakMap();
const sheetsOf = buffer => {
  if (!sheetsCache.has(buffer)) sheetsCache.set(buffer, readSheets(XLSX, buffer));
  return sheetsCache.get(buffer);
};

// onlySheet — имя листа (при восстановлении рабочей области); first — только первый лист (рабочие области версии 1.0)
function addScheduleBuffer(file, buffer, { onlySheet = null, first = false } = {}) {
  S.schedules = S.schedules.filter(s => s.file !== file || (onlySheet && s.sheet !== onlySheet));    // тот же файл загружен заново — заменяется
  const sheets = sheetsOf(buffer);
  let added = 0;
  let firstError = '';
  for (const { name: sheetName, sheet } of first ? sheets.slice(0, 1) : sheets) {
    if (onlySheet && sheetName !== onlySheet) continue;
    try {
      const parsed = parseSchedule(sheet);
      parsed.semester ||= Number(/([12])\s*сем/i.exec(file)?.[1] || 0);
      S.schedules.push({ id: `${Date.now()}${S.schedules.length}`, file, sheet: sheetName, buffer, parsed });
      added++;
    } catch (e) { firstError ||= e.message; }
  }
  if (!added) throw new Error(`${file}: ${firstError || 'расписание не найдено'}`);
}

function rebuildSchedule() {
  S.schedule = S.schedules.length
    ? mergeSchedules(S.schedules.map(s => ({ name: scheduleLabel(s), parsed: s.parsed })), { graphYear: S.graph?.startYear || 0 })
    : null;
  renderScheduleList();
  const names = S.schedule ? teacherNames(S.schedule.entries) : [];
  fillTeacherSelects(names);
  if (S.teacher && !names.includes(S.teacher)) S.teacher = '';
  syncBells();
  afterInputsChanged();
}

const scheduleLabel = s => (S.schedules.filter(x => x.file === s.file).length > 1 ? `${s.file} · ${s.sheet}` : s.file);

function renderScheduleList() {
  const info = S.schedule
    ? `${S.schedules.length} расписан., ${S.schedule.groups.length} групп, ${teacherNames(S.schedule.entries).length} педагогов`
    : 'не загружено';
  setInfo('schedule', info, !!S.schedule);
  document.querySelectorAll('[data-list="schedule"]').forEach(ul => {
    ul.innerHTML = S.schedules.map(s => `<li data-id="${s.id}"><span class="grow"><b>${esc(scheduleLabel(s))}</b>
      <small>${s.parsed.course ? `${s.parsed.course} курс` : 'курс не указан'}${s.parsed.semesterNo ? `, ${s.parsed.semesterNo} семестр` : ''}${s.parsed.startYear ? `, ${s.parsed.startYear}–${s.parsed.startYear + 1}` : ''} · групп ${s.parsed.groups.length}</small></span>
      <button type="button" title="Убрать расписание">✕</button></li>`).join('');
    ul.querySelectorAll('li').forEach(li => {
      li.querySelector('button').onclick = () => { S.schedules = S.schedules.filter(x => x.id !== li.dataset.id); rebuildSchedule(); };
    });
  });
}

function fillTeacherSelects(names) {
  const options = '<option value="">— выберите педагога —</option>' + names.map(n => `<option>${esc(n)}</option>`).join('');
  for (const id of ['teacher', 'planTeacher']) { $(id).innerHTML = options; $(id).value = names.includes(S.teacher) ? S.teacher : ''; }
}

async function loadPracticeFiles(files) {
  for (const file of files) await addPracticeBuffer(file.name, await file.arrayBuffer(), '');
  afterInputsChanged();
}

async function addPracticeBuffer(file, buffer, teacher) {
  try {
    const data = await parsePractice(JSZip, buffer);
    S.practices.push({ id: `${Date.now()}${S.practices.length}`, file, buffer, data, teacher });
  } catch (e) { showMessage(`${file}: ${e.message}`, true); }
}

function afterInputsChanged() {
  if (S.schedule) {
    const names = teacherNames(S.schedule.entries);
    for (const p of S.practices) if (!p.teacher) p.teacher = matchSupervisor(names, p.data.supervisor) || '';
  }
  renderPractices();
  fillSigners();
  updateSpecialtyOptions();
  if (S.teacher) selectTeacher(S.teacher, { fresh: true, keepFields: true });   // документы изменились — строки собираются заново
  else { recompute(); }
}

function renderPractices() {
  const names = S.schedule ? teacherNames(S.schedule.entries) : [];
  $('practiceList').innerHTML = S.practices.map(p => `<li data-id="${p.id}">
    <span class="grow"><b>${esc(p.file)}</b> — группа ${esc(p.data.group)}, ${esc(p.data.moduleIndex)}, ${ruDate(p.data.from)} – ${ruDate(p.data.to)}, ${p.data.total} ч</span>
    <select>${['', ...names].map(n => `<option value="${esc(n)}"${n === p.teacher ? ' selected' : ''}>${n ? esc(n) : '— педагог не назначен —'}</option>`).join('')}</select>
    <button type="button" title="Убрать документ">✕</button></li>`).join('');
  for (const li of $('practiceList').children) {
    const p = S.practices.find(x => x.id === li.dataset.id);
    li.querySelector('select').onchange = e => { p.teacher = e.target.value; afterInputsChanged(); };
    li.querySelector('button').onclick = () => { S.practices = S.practices.filter(x => x !== p); afterInputsChanged(); };
  }
}

// ───────── педагог и строки ведомости ─────────

function stashTeacher() {
  if (!S.teacher) return;
  S.teacherStates[S.teacher] = {
    rows: S.rows, warnings: S.warnings, semester: $('semester').value,
    fields: Object.fromEntries(FIELD_IDS.map(id => [id, $(id).value])),
  };
}

// fresh — собрать строки заново из документов; keepFields — не очищать введённые поля
function selectTeacher(name, { fresh = false, keepFields = false } = {}) {
  stashTeacher();
  if (fresh) delete S.teacherStates[name];
  if (S.teacher !== name) S.plan.groups = null;          // таблица недели по умолчанию показывает группы педагога
  S.teacher = name;
  $('teacher').value = name;
  $('planTeacher').value = name;
  S.rows = []; S.warnings = [];
  if (name && S.graph && S.schedule) {
    const saved = S.teacherStates[name];
    if (saved) {
      ({ rows: S.rows, warnings: S.warnings } = saved);
      FIELD_IDS.forEach(id => { $(id).value = saved.fields[id] ?? ''; });
      $('semester').value = saved.semester;
    } else {
      const practiceDocs = S.practices.filter(p => p.teacher === name).map(p => p.data);
      ({ rows: S.rows, warnings: S.warnings } = buildRows({ teacher: name, entries: S.schedule.entries, graph: S.graph, practiceDocs }));
      if (!keepFields) {
        FIELD_IDS.forEach(id => { $(id).value = ''; });
        $('semester').value = '0';
      }
      if (S.rows.length && !S.rows.some(isVisible)) S.funding = S.funding === 'grant' ? 'comm' : 'grant';
    }
  }
  applyLoadToRows();
  renderRows();
  renderPracticeBlocks();
  recompute();
}

function renderRows() {
  const body = $('rowsTable').tBodies[0];
  body.innerHTML = S.rows.map((r, i) => (isVisible(r) ? [r, i] : null)).filter(Boolean).map(([r, i]) => {
    const legend = S.graph && S.graph.legendFor(r.group);
    const mods = legend ? Object.values(legend) : [];
    const known = mods.some(m => m.index === r.moduleIndex);
    const indexCell = legend
      ? `<select data-f="moduleIndex"><option value="">— без модуля —</option>${r.moduleIndex && !known ? `<option value="${esc(r.moduleIndex)}" selected>${esc(r.moduleIndex)}</option>` : ''}${mods.map(m =>
          `<option value="${esc(m.index)}"${m.index === r.moduleIndex ? ' selected' : ''}>${esc(m.index)} ${esc(m.title)}</option>`).join('')}</select>`
      : `<input data-f="moduleIndex" value="${esc(r.moduleIndex)}" placeholder="нет в графике">`;
    const semIdx = Number($('semester').value);
    const loadCell = S.load
      ? `<td class="ld"><select data-f="load"><option value="">—</option>${loadCandidates(S.load, r.group, semIdx).map(c =>
          `<option value="${c.i}"${c.i === r.loadIndex ? ' selected' : ''}>${esc(loadLabel(c, semIdx))}</option>`).join('')}</select></td>`
      : '<td class="ld"></td>';
    return `<tr data-i="${i}">
      <td>${r.kind === 'practice' ? 'практика' : 'теория'}</td>
      <td>${indexCell}</td>
      <td><input data-f="name" value="${esc(r.name)}"></td>
      <td><input data-f="group" value="${esc(r.group)}"></td>
      <td><input data-f="plan" type="number" min="0" value="${r.plan ?? ''}"></td>
      <td class="s2"><input data-f="base" type="number" min="0" value="${r.base ?? ''}"></td>
      ${loadCell}
      <td><button type="button" title="Убрать из ведомости">✕</button></td></tr>`;
  }).join('');

  body.querySelectorAll('tr[data-i]').forEach(tr => {
    const row = S.rows[Number(tr.dataset.i)];
    const nameInput = tr.querySelector('[data-f="name"]');
    tr.querySelectorAll('input[data-f]').forEach(input => {
      input.oninput = () => {
        row[input.dataset.f] = ['plan', 'base'].includes(input.dataset.f) ? (input.value === '' ? null : Number(input.value)) : input.value;
        if (input.dataset.f === 'plan') row.planFromLoad = false;           // введённое вручную больше не заменяется нагрузкой
        recompute();
      };
    });
    // выбор индекса подставляет полное название модуля; «без модуля» возвращает название из документа
    const indexField = tr.querySelector('[data-f="moduleIndex"]');
    const onIndex = () => {
      row.moduleIndex = indexField.value.trim();
      const mod = row.moduleIndex && S.graph && S.graph.moduleOf(row.group, row.moduleIndex);
      row.name = mod ? (row.kind === 'practice' ? PRACTICE_PREFIX : '') + mod.title : row.defaults.name;
      nameInput.value = row.name;
      recompute();
    };
    if (indexField.tagName === 'SELECT') indexField.onchange = onIndex;
    else indexField.onchange = () => { row.moduleIndex = indexField.value.trim(); recompute(); };
    indexField.oninput = indexField.tagName === 'SELECT' ? null : () => { row.moduleIndex = indexField.value; recompute(); };
    const loadSelect = tr.querySelector('select[data-f="load"]');
    if (loadSelect) {
      loadSelect.onchange = () => {
        const sem = Number($('semester').value);
        row.loadIndex = loadSelect.value === '' ? null : Number(loadSelect.value);
        const src = row.loadIndex != null && S.load.rows[row.loadIndex].sems[sem];
        if (src) { row.plan = src.total; row.planFromLoad = true; tr.querySelector('[data-f="plan"]').value = src.total; }
        renderPracticeBlocks(); recompute();
      };
    }
    tr.querySelector('button').onclick = () => { S.rows.splice(Number(tr.dataset.i), 1); renderRows(); renderPracticeBlocks(); recompute(); };
  });
  $('rowsTable').classList.toggle('no-load', !S.load);
  $('trRow').innerHTML = visibleRows().filter(r => r.kind === 'theory').map(r => `<option value="${r.id}">${esc(r.name)} (${esc(r.group)})</option>`).join('');
  $('prAdd').disabled = !(S.teacher && S.graph);
}

// ───────── практика, «Заполнить всё по документам», подсветка обязательных полей ─────────

function addPractice() {
  const group = S.rows[0]?.group || Object.values(S.graph.groups)[0]?.name || '';
  S.rows.push(emptyPracticeRow({ id: `p${Date.now()}`, group, start: S.graph.weeks[0].start }));
  renderRows(); renderPracticeBlocks(); recompute();
}

function fillAll() {
  if (!S.graph || !S.schedule) { $('fillReport').textContent = 'Сначала загрузите график и расписание (шаг 1).'; return; }
  if (!S.teacher) { $('teacher').classList.add('need'); $('teacher').focus(); $('fillReport').textContent = 'Выберите педагога.'; return; }
  const practiceSpecialty = S.practices.filter(p => p.teacher === S.teacher).map(p => p.data.specialty).find(Boolean);
  if (S.graph.college) $('college').value = S.graph.college;
  if (S.graph.deputy) $('deputy').value = S.graph.deputy;
  if (practiceSpecialty) $('specialty').value = practiceSpecialty;
  fillSigners(true);
  selectTeacher(S.teacher, { fresh: true, keepFields: true });      // строки собираются заново: возвращаются удалённые, сбрасываются правки
  const empty = document.querySelectorAll('.need').length;
  const time = new Date().toLocaleTimeString('ru-RU');
  $('fillReport').textContent = `${time} — пересобрано по документам: дисциплин ${S.rows.length}. ` +
    (empty ? `Красным выделено то, чего в документах нет (${empty}) — введите вручную.` : 'Всё заполнено.');
}

// Пустые обязательные поля подсвечиваются красным (заголовки таких полей помечены звёздочкой)
function markRequired() {
  const ready = !!S.teacher;
  $('teacher').classList.toggle('need', !S.teacher && !!S.schedule);
  const sem2 = $('semester').value === '1';
  $('prevBox').classList.toggle('need', sem2 && !S.prev);
  document.querySelectorAll('[data-box="schedule"]').forEach(el => el.classList.toggle('need', sem2 && !!S.schedule?.semester && S.schedule.semester !== 2));
  for (const id of ['fullName', 'college', 'deputy']) $(id).classList.toggle('need', ready && !$(id).value.trim());
  specialtyCombo.setNeed(ready && !$('specialty').value.trim());
  $('rowsTable').querySelectorAll('tr[data-i]').forEach(tr => {
    const row = S.rows[Number(tr.dataset.i)];
    tr.querySelectorAll('input, select').forEach(el => {
      const optional = el.dataset.f === 'load' || (el.dataset.f === 'moduleIndex' && row.kind === 'theory');   // у теории (например, физкультура) модуля может не быть
      const skip = el.dataset.f === 'base' && !sem2;
      const unknownGroup = el.dataset.f === 'group' && !!S.graph && el.value.trim() !== '' && !S.graph.hasGroup(el.value);
      el.classList.toggle('need', (ready && !optional && !skip && el.value.trim() === '') || unknownGroup);
    });
  });
  document.querySelectorAll('.block input[data-f="hoursPerDay"]').forEach(input => {
    input.classList.toggle('need', !(Number(input.value) > 0));
  });
}

// ───────── нагрузка, подписанты, поля с выбором ─────────

function mountSources() {
  for (const host of [$('sourcesMain'), $('sourcesPlan')]) {
    host.innerHTML = SOURCES_HTML;
    host.querySelector('[data-src="graph"]').onchange = e => loadGraphFile(e.target.files[0]);
    host.querySelector('[data-src="schedule"]').onchange = e => { addScheduleFiles([...e.target.files]).then(() => { e.target.value = ''; }); };
    host.querySelector('[data-src="load"]').onchange = e => loadLoadFile(e.target.files[0]);
    host.querySelector('[data-src="practice"]').onchange = e => loadPracticeFiles([...e.target.files]).then(() => { e.target.value = ''; });
  }
}

// Подпись у файла одного вида — на обеих вкладках сразу
const setInfo = (kind, text, ok) => document.querySelectorAll(`[data-info="${kind}"]`).forEach(el => { el.textContent = text; el.className = ok ? 'ok' : ''; });

function mountCombos() {
  specialtyCombo = makeCombo($('specialtyCombo'), {
    placeholder: 'Выберите специальность или впишите свою',
    onChange: () => $('specialty').dispatchEvent(new Event('input', { bubbles: true })),
  });
  // скрытое поле остаётся «источником правды» для остального кода: чтение отдаёт значения через запятую
  Object.defineProperty($('specialty'), 'value', { get: () => specialtyCombo.text, set: v => specialtyCombo.setText(v), configurable: true });
  groupsCombo = makeCombo($('planGroups'), {
    placeholder: 'Выберите группы', allowCustom: false,
    onChange: items => { S.plan.groups = items; markDirty(); renderPlan(); },
  });
}

const specialtyOptions = () => [...new Set([
  ...(S.graph?.specialties || []),
  ...S.practices.flatMap(p => [p.data.specialtyOnly, p.data.qualification]).filter(Boolean),
])];
const updateSpecialtyOptions = () => specialtyCombo.setOptions(specialtyOptions());

// Подписанты документа практики: из загруженных документов практики (force — перезаписать введённое)
function fillSigners(force = false) {
  const docs = S.practices.map(p => p.data.signers).filter(Boolean);
  for (const k of ['approver', 'upr', 'head']) {
    const found = docs.map(d => d[k]).find(Boolean);
    if (found && (force || !S.signers[k])) S.signers[k] = found;
  }
  for (const [id, k] of [['sgApprover', 'approver'], ['sgUpr', 'upr'], ['sgHead', 'head']]) {
    if ($(id) !== document.activeElement) $(id).value = S.signers[k] || '';
  }
}

// ───────── нагрузка ─────────

async function loadLoadFile(file) {
  if (file) await loadLoadBuffer(file.name, await file.arrayBuffer());
}

async function loadLoadBuffer(name, buffer) {
  try {
    const sheets = readSheets(XLSX, buffer);
    let parsed = null;
    let sheetName = '';
    let lastError = null;
    for (const { name: n, sheet } of sheets) {
      try { parsed = parseLoad(sheet); sheetName = n; break; } catch (e) { lastError = e; }
    }
    if (!parsed) throw lastError || new Error('В файле не найдена нагрузка');
    S.load = parsed;
    S.files.load = { name, buffer };
    setInfo('load', `${name}${sheets.length > 1 ? ` · ${sheetName}` : ''}: строк нагрузки ${parsed.rows.length}`, true);
    applyLoadToRows();
    renderRows(); renderPracticeBlocks(); recompute();
  } catch (e) { S.load = null; S.files.load = null; fail('load', e); renderRows(); }
}

// План из нагрузки подставляется только в пустые поля (введённое вручную не трогается);
// подставленное ранее пересчитывается при смене семестра. В 2 семестре с загруженной ведомостью 1 семестра план берётся из неё.
function applyLoadToRows() {
  if (!S.load || !S.rows.length) return;
  const idx = Number($('semester').value);
  if (S.prev && idx === 1) return;
  for (const r of S.rows) if (r.planFromLoad) r.plan = null;
  applyLoad(S.rows, S.load, idx);
}

// ───────── вкладка «Расписание и практики» ─────────

function initPlanTab() {
  for (const [id, k] of [['sgApprover', 'approver'], ['sgUpr', 'upr'], ['sgHead', 'head']]) $(id).oninput = () => { S.signers[k] = $(id).value; };
  $('planOnlyMine').onchange = () => { S.plan.onlyMine = $('planOnlyMine').checked; renderPlan(); };
  $('planShowEmpty').onchange = () => { S.plan.showEmpty = $('planShowEmpty').checked; renderPlan(); };
  $('planMine').onclick = () => { S.plan.groups = null; renderPlan(); };
  $('planSwap').onclick = () => { S.cal.swapWeeks = !S.cal.swapWeeks; recompute(); };
  $('planPrAdd').onclick = addPractice;
  $('bellsText').oninput = () => { S.plan.bells = $('bellsText').value; S.plan.bellsManual = true; renderPlan(); };
  $('bellsAuto').onclick = () => { S.plan.bellsManual = false; syncBells(); renderPlan(); };
}

// Звонки: из расписаний, пока пользователь не правил список вручную
function syncBells() {
  if (!S.plan.bellsManual) S.plan.bells = bellTable(S.schedule?.entries || []).join('\n');
  $('bellsText').value = S.plan.bells;
}

function currentBells() {
  const typed = S.plan.bellsManual ? parseBells(S.plan.bells) : [];
  return typed.length ? typed : bellTable(S.schedule?.entries || []);
}

const teacherGroups = () => (S.schedule && S.teacher ? [...new Set(S.schedule.entries.filter(e => e.teacher === S.teacher).map(e => e.group))] : []);

function renderPlan() {
  const entries = S.schedule?.entries || [];
  $('planOnlyMine').checked = !!S.plan.onlyMine;
  $('planShowEmpty').checked = !!S.plan.showEmpty;
  $('planSwap').textContent = $('swapWeeks').textContent;
  $('bellsInfo').textContent = S.plan.bellsManual ? 'список задан вручную' : `собрано из расписаний: ${currentBells().length} звонков`;
  groupsCombo.setOptions(S.schedule?.groups || []);
  const groups = S.plan.groups ?? teacherGroups();
  groupsCombo.items = groups;
  renderPracticeBlocks();

  if (!entries.length) {
    $('planGrid').innerHTML = '<p class="hint">Загрузите расписание (можно несколько — по одному на курс) и выберите педагога.</p>';
    $('planSummary').innerHTML = '';
    return;
  }
  const grid = buildGrid({ entries, bells: currentBells(), teacher: S.teacher, groups: groups.length ? groups : null, onlyMine: S.plan.onlyMine && !!S.teacher });
  renderGrid($('planGrid'), grid, { showEmpty: S.plan.showEmpty, swap: !!S.cal.swapWeeks });

  const rows = S.teacher ? weeklySummary(entries, S.teacher) : [];
  const h = n => `${n} пар · ${n * HOURS_PER_PAIR} ч`;
  const left = rows.reduce((a, r) => a + r.left, 0);
  const right = rows.reduce((a, r) => a + r.right, 0);
  $('planSummary').innerHTML = rows.length
    ? `<h3>Занятия педагога в неделю</h3><table class="rows"><thead><tr><th>Группа</th><th>Дисциплина</th><th>Левая неделя</th><th>Правая неделя</th></tr></thead><tbody>${
      rows.map(r => `<tr><td>${esc(r.group)}</td><td>${esc(r.subject)}</td><td>${h(r.left)}</td><td>${h(r.right)}</td></tr>`).join('')
    }<tr><th colspan="2">Всего</th><th>${h(left)}</th><th>${h(right)}</th></tr></tbody></table>`
    : '';
}

// ───────── практика: карточки, расширенный вид, документ ─────────

// Календарь практики: суббота — рабочий день (практика идёт и по субботам); праздники и правки календаря учитываются
const practiceCalendar = () => createCalendar({ holidays: S.holidays, state: S.cal, workSaturday: true, years: calendarYears() });

const fundingLabel = row => ({ grant: 'грант', comm: 'коммерция' }[fundingOfRow(row)] || 'группы нет в графике');

function renderPracticeBlocks() {
  fillPracticeCards($('practiceBlocks'), visibleRows().filter(r => r.kind === 'practice'), false);
  fillPracticeCards($('planPractices'), S.teacher ? S.rows.filter(r => r.kind === 'practice') : [], true);
  refreshPracticeSummaries();
  renderSuggestions();
}

// Блок практики: период × часов в день (простой режим) либо точные дни по парам (расширенный вид, документ практики)
function fillPracticeCards(container, rows, onPlanTab) {
  container.innerHTML = rows.length ? rows.map(r => `<div class="block" data-id="${r.id}">
    <b>${esc(r.name)}</b> <small>(${esc(r.group)}${onPlanTab ? ` · ${fundingLabel(r)}` : ''})</small>
    ${r.blocks.map((b, i) => b.days
      ? `<div class="row" data-b="${i}">по дням: ${ruDate(b.from)} – ${ruDate(b.to)}, ${Object.values(b.days).reduce((a, h) => a + h, 0)} ч <button type="button" data-act="del">✕</button></div>`
      : `<div class="row" data-b="${i}">с <input type="date" data-f="from" value="${b.from}"> по <input type="date" data-f="to" value="${b.to}">
          <input type="number" data-f="hoursPerDay" min="0" value="${b.hoursPerDay}" style="width:70px"> ч/день <b class="star">*</b>
          <label class="inline"><input type="checkbox" data-f="saturday"${b.saturday ? ' checked' : ''}> и суббота</label>
          <button type="button" data-act="del">✕</button></div>`).join('')}
    <div class="row"><button type="button" data-act="add">+ период</button>
      <button type="button" data-act="graph">Периоды из графика</button>
      <button type="button" data-act="open" class="primary">Расширенный вид</button>
      <button type="button" data-act="doc">Документ практики (.docx)</button></div>
    <p class="pdesc" data-desc></p></div>`).join('')
    : (onPlanTab ? '<p class="hint">У выбранного педагога пока нет практики: добавьте из списка ниже или кнопкой «+ Практика».</p>' : '');

  container.querySelectorAll('.block').forEach(box => {
    const row = S.rows.find(r => r.id === box.dataset.id);
    box.querySelectorAll('.row[data-b]').forEach(line => {
      const b = row.blocks[Number(line.dataset.b)];
      line.querySelectorAll('input').forEach(input => {
        input.oninput = () => { b[input.dataset.f] = input.type === 'checkbox' ? input.checked : input.value; recompute(); };
      });
      line.querySelector('[data-act="del"]').onclick = () => { row.blocks.splice(Number(line.dataset.b), 1); renderPracticeBlocks(); recompute(); };
    });
    box.querySelector('[data-act="add"]').onclick = () => {
      const last = row.blocks[row.blocks.length - 1];
      row.blocks.push({ from: last?.to || S.graph.weeks[0].start, to: last?.to || S.graph.weeks[0].start, hoursPerDay: last?.hoursPerDay ?? '', saturday: false });
      renderPracticeBlocks(); recompute();
    };
    // периоды берутся только при точном совпадении индекса модуля с «Обозначениями» графика для группы
    box.querySelector('[data-act="graph"]').onclick = () => {
      const periods = graphPeriods(S.graph, row.group, row.moduleIndex.trim());
      if (!periods.length) { showMessage(`В графике не нашлось недель практики для группы ${row.group} и индекса «${row.moduleIndex}». Укажите индекс из подсказки в таблице или добавьте период вручную.`, true); return; }
      row.blocks = periods;
      renderPracticeBlocks(); recompute();
    };
    box.querySelector('[data-act="open"]').onclick = () => openPractice(row);
    box.querySelector('[data-act="doc"]').onclick = () => downloadPracticeDoc(row, row.blocks.map(b => expandBlock(b, practiceCalendar())), practiceDocMeta(row));
  });
}

// Под карточкой: «пн 4-6 пара, вт 4-6 пара … — всего N ч · по нагрузке M ч: ✓ / ⚠»
function refreshPracticeSummaries() {
  const calendar = practiceCalendar();
  document.querySelectorAll('.block[data-id]').forEach(box => {
    const row = S.rows.find(r => r.id === box.dataset.id);
    const el = box.querySelector('[data-desc]');
    if (!row || !el) return;
    const blocks = row.blocks.map(b => expandBlock(b, calendar));
    const total = totalHours(blocks);
    const desc = describePractice(blocks);
    const sem = S.load && row.loadIndex != null ? S.load.rows[row.loadIndex]?.sems[Number($('semester').value)] : null;
    const cmp = sem ? ` · по нагрузке ${sem.total} ч${sem.weeks && sem.perWeek ? ` (${sem.weeks} нед × ${sem.perWeek})` : ''}: ${sem.total === total ? '✓ совпадает' : `⚠ разница ${total - sem.total} ч`}` : '';
    el.textContent = `${desc ? `${desc} — ` : ''}всего ${total} ч${cmp}`;
    el.classList.toggle('warn', !!sem && sem.total !== total);
  });
}

const courseOfGroup = group => S.schedules.find(s => s.parsed.groups.some(g => sameGroup(g, group)))?.parsed.course || '';

// Реквизиты документа практики: введённые ранее → из загруженного документа практики → из специальностей педагога и графика
function practiceDocMeta(row) {
  const doc = row.doc || {};
  const imported = S.practices.find(p => p.teacher === S.teacher && sameGroup(p.data.group, row.group) && p.data.moduleIndex === row.moduleIndex)?.data;
  const items = specialtyCombo.items;
  const mod = S.graph && S.graph.moduleOf(row.group, row.moduleIndex);
  return {
    specialty: doc.specialty ?? imported?.specialtyOnly ?? items.find(s => /^\d{8}\b/.test(s)) ?? '',
    qualification: doc.qualification ?? imported?.qualification ?? items.find(s => /^[34][SW]\d{8}/.test(s)) ?? '',
    course: doc.course ?? imported?.course ?? courseOfGroup(row.group),
    moduleLine: doc.moduleLine ?? (mod ? `${mod.index} ${mod.title}` : [row.moduleIndex, row.name.replace(/^УП\s*-\s*/i, '')].filter(Boolean).join(' ')),
  };
}

function openPractice(row) {
  const sem = S.load && row.loadIndex != null ? S.load.rows[row.loadIndex]?.sems[Number($('semester').value)] : null;
  openPracticeDialog($('prDialog'), {
    row, graph: S.graph, calendar: practiceCalendar(), bells: currentBells(), doc: practiceDocMeta(row), specialties: specialtyOptions(),
    loadInfo: sem ? { weeks: sem.weeks, perWeek: sem.perWeek, total: sem.total } : null,
    onSave: (blocks, doc) => {
      row.blocks = blocks;
      row.doc = { ...doc };
      renderPracticeBlocks(); recompute();
    },
    onDownload: (blocks, doc) => downloadPracticeDoc(row, blocks, doc),
  });
}

async function downloadPracticeDoc(row, blocks, doc) {
  const y = S.graph.startYear;
  const meta = {
    college: $('college').value.trim(), yearLabel: `${y}-${y + 1}`, startYear: y,
    group: row.group, course: doc.course, specialty: doc.specialty, qualification: doc.qualification, moduleLine: doc.moduleLine,
    supervisor: shortSigner($('fullName').value.trim() || S.teacher), approver: S.signers.approver || $('deputy').value.trim(),
    upr: S.signers.upr, head: S.signers.head,
  };
  const blob = await buildPracticeDocx(JSZip, meta, blocks.filter(b => Object.keys(b.days || {}).length), practiceCalendar());
  download_(blob, `Расписание_практики_${row.group.replace(/[\\/]/g, '-')}_${row.moduleIndex || 'УП'}.docx`);
}

// Практики, которые есть в графике для групп педагога или в нагрузке, но ещё не добавлены в ведомость
function practiceSuggestions() {
  if (!S.graph || !S.teacher) return [];
  const semIdx = Number($('semester').value);
  const groups = [...new Set(S.rows.filter(r => r.kind === 'theory').map(r => r.group))];
  const key = (g, index) => `${groupKeys(g).join('/')}|${normIndex(index)}`;
  const have = new Set(S.rows.filter(r => r.kind === 'practice').map(r => key(r.group, r.moduleIndex)));
  const out = [];
  if (S.load) {
    // нагрузка точно говорит, какие модули у групп педагога; период — из графика (тот же индекс, иначе то же название)
    S.load.rows.forEach((lr, i) => {
      const g = lr.moduleIndex && lr.sems[semIdx] && groups.find(x => sameGroup(x, lr.group));
      if (!g) return;
      const legend = Object.values(S.graph.legendFor(g) || {});
      const mod = legend.find(m => normIndex(m.index) === lr.moduleIndex) || legend.find(m => sameTitle(m.title, lr.title));
      out.push({ group: g, moduleIndex: mod ? mod.index : lr.moduleIndex, title: mod ? mod.title : lr.title, loadIndex: i, loadModule: lr.moduleIndex, from: 'нагрузка' });
    });
  } else {
    for (const g of groups) {
      for (const [code, m] of Object.entries(S.graph.legendFor(g) || {})) {
        if (m.index && S.graph.runsOf(g, code).length) out.push({ group: g, moduleIndex: m.index, title: m.title, from: 'график' });
      }
    }
  }
  const list = out.filter(s => !have.has(key(s.group, s.moduleIndex)));
  // период практики — только те, что приходятся на выбранный семестр (график охватывает весь год)
  const sem = S.graph.semesters(groups)[semIdx];
  const inSemester = r => !sem || (ymIndex(ymOf(r.from)) >= ymIndex(sem.from) && ymIndex(ymOf(r.from)) <= ymIndex(sem.to));
  for (const s of list) s.runs = graphPeriods(S.graph, s.group, s.moduleIndex).filter(inSemester);
  return list;
}

function renderSuggestions() {
  const list = practiceSuggestions();
  const periods = s => (s.runs.length ? ` (${s.runs.map(r => `${ruDate(r.from).slice(0, 5)}–${ruDate(r.to).slice(0, 5)}`).join(', ')})` : ' (период в графике не найден)');
  $('planSuggest').innerHTML = list.length
    ? `<h3>Практики по документам для групп педагога</h3><p class="hint">${S.load ? 'Из нагрузки; период — из графика.' : 'Из графика (загрузите нагрузку — останутся только модули педагога).'} Кто ведёт практику, документы не говорят — добавьте нужные.</p><div class="row">${
      list.map((s, i) => `<button type="button" data-i="${i}">+ ${esc(s.group)} · ${esc(s.moduleIndex)} ${esc(s.title.slice(0, 40))}${s.loadModule && s.loadModule !== normIndex(s.moduleIndex) ? ` <small>(в нагрузке ${esc(s.loadModule)})</small>` : ''}${periods(s)}</button>`).join('')}</div>`
    : '';
  $('planSuggest').querySelectorAll('button').forEach(b => { b.onclick = () => addSuggestedPractice(list[Number(b.dataset.i)]); });
}

function addSuggestedPractice(s) {
  const row = emptyPracticeRow({ id: `p${Date.now()}`, group: s.group, start: S.graph.weeks[0].start });
  row.moduleIndex = s.moduleIndex;
  row.name = PRACTICE_PREFIX + s.title;
  if (s.runs.length) row.blocks = s.runs;
  if (s.loadIndex != null) row.loadIndex = s.loadIndex;
  row.defaults = { ...row.defaults, moduleIndex: row.moduleIndex, name: row.name, blocks: row.blocks.map(b => ({ ...b })) };
  S.rows.push(row);
  applyLoadToRows();
  renderRows(); renderPracticeBlocks(); recompute();
}

// ───────── переносы и отработки занятий ─────────

const transfersOf = () => (S.transfers[S.teacher] ||= []);

function addTransfer() {
  const rowId = $('trRow').value;
  const from = $('trFrom').value;
  const to = $('trTo').value;
  const hours = Number($('trHours').value);
  if (!rowId || !from || !to || !(hours > 0)) { showMessage('Для переноса укажите дисциплину, дату занятия, дату проведения и часы.', true); return; }
  transfersOf().push({ id: `${Date.now()}`, rowId, from, to, hours });
  $('trFrom').value = $('trTo').value = '';
  showMessage('');
  recompute();
}

function renderTransfers() {
  const list = transfersOf();
  $('trList').innerHTML = list.map(t => {
    const row = S.rows.find(r => r.id === t.rowId);
    return `<li data-id="${t.id}"><span class="grow">${t.hours} ч — занятие за <b>${ruDate(t.from)}</b> проведено <b>${ruDate(t.to)}</b> · ${esc(row ? row.name : '—')}</span><button type="button">✕</button></li>`;
  }).join('');
  $('trList').querySelectorAll('li').forEach(li => {
    li.querySelector('button').onclick = () => { S.transfers[S.teacher] = list.filter(t => t.id !== li.dataset.id); recompute(); };
  });

  const missed = S.result?.missed || [];
  $('missed').innerHTML = missed.length
    ? `<h3>Занятия, выпавшие на нерабочие дни (${missed.length})</h3><ul class="list">${missed.map((m, i) => `<li><span class="grow">${ruDate(m.date)} — ${esc(m.reason)}: ${esc(m.name)}, ${m.hours} ч</span><button type="button" data-i="${i}">Назначить отработку</button></li>`).join('')}</ul>`
    : '';
  $('missed').querySelectorAll('button').forEach(b => {
    b.onclick = () => {
      const m = missed[Number(b.dataset.i)];
      $('trRow').value = m.rowId; $('trFrom').value = m.date; $('trHours').value = m.hours; $('trTo').focus();
    };
  });
}

// ───────── расчёт и предпросмотр ─────────

const workSaturday = () => {
  if (S.cal.workSaturday !== null && S.cal.workSaturday !== undefined) return S.cal.workSaturday;
  const practiceSat = S.rows.some(r => r.kind === 'practice' && r.blocks.some(b => b.saturday || (b.days && Object.entries(b.days).some(([d, h]) => h > 0 && weekday(d) === 6))));
  return !!(S.schedule && S.schedule.hasSaturday) || practiceSat;
};

// Годы, для которых строятся праздники: всё, что встречается в данных, с запасом
function calendarYears() {
  const ys = [new Date().getFullYear(), S.calStart?.y, S.graph?.startYear, ...Object.keys(S.cal.overrides).map(d => Number(d.slice(0, 4)))].filter(Boolean);
  return [Math.min(...ys) - 1, Math.max(...ys) + 2];
}

function makeCalendar() {
  return createCalendar({
    holidays: S.holidays, state: S.cal, workSaturday: workSaturday(),
    years: calendarYears(),
  });
}

function currentPeriod() {
  if (!S.graph) return null;
  const sems = S.graph.semesters([...new Set(S.rows.map(r => r.group))]);
  const option = $('semester').options[1];
  option.disabled = sems.length < 2;
  if (sems.length < 2) $('semester').value = '0';
  const idx = Number($('semester').value);
  const p = sems[idx];
  $('period').textContent = `${MONTH_NAMES[p.from.m]} ${p.from.y} – ${MONTH_NAMES[p.to.m]} ${p.to.y} (по каникулам из графика)`;
  return { ...p, idx };
}

function recompute() {
  markDirty();
  S.result = null;
  const problems = [...S.warnings, ...(S.schedule?.warnings || [])];
  const period = currentPeriod();
  $('workSat').checked = workSaturday();

  const sem2 = !!period && period.idx === 1;
  $('prevBox').hidden = !sem2;
  $('rowsTable').classList.toggle('is-sem1', !sem2);
  $('rowsTable').classList.toggle('is-sem2', sem2);
  $('annualWrap').hidden = !sem2;
  if (sem2 && S.prev) {
    const found = applyPrevSemester(S.rows, S.prev);          // подставляет только пустые «с 1 сем.» и план
    $('prevInfo').textContent = `${S.files.prev.name}: сопоставлено строк ${found} из ${S.rows.length}` +
      (found < S.rows.length ? ' — у остальных выберите тот же индекс модуля, что в ведомости 1 семестра, или введите «С 1 сем.» вручную' : '');
    // подставленные значения видны в полях таблицы
    $('rowsTable').querySelectorAll('tr[data-i]').forEach(tr => {
      const row = S.rows[Number(tr.dataset.i)];
      for (const f of ['plan', 'base']) {
        const input = tr.querySelector(`[data-f="${f}"]`);
        if (input && row[f] != null && input.value === '') input.value = row[f];
      }
    });
  }

  updateFundSwitch();
  if (S.graph && visibleRows().length && period) {
    S.result = buildTimesheet({
      rows: visibleRows(), graph: S.graph, calendar: makeCalendar(), from: period.from, to: period.to,
      carryOver: $('carry').checked && period.idx === 0, transfers: transfersOf(), swapWeeks: !!S.cal.swapWeeks, useBase: sem2,
    });
    // годовой учёт охватывает все группы педагога (грант и коммерцию вместе, по страницам типов нагрузки)
    S.resultAll = sem2 ? buildTimesheet({
      rows: S.rows, graph: S.graph, calendar: makeCalendar(), from: period.from, to: period.to,
      carryOver: false, transfers: transfersOf(), swapWeeks: !!S.cal.swapWeeks, useBase: true,
    }) : null;
    $('rowsTable').tBodies[0].querySelectorAll('input[data-f="plan"]').forEach((input, i) => { input.placeholder = `по расписанию ${S.result.totals[i]}`; });
    for (const g of new Set(S.rows.map(r => r.group))) {
      if (!S.graph.hasGroup(g)) problems.push(`Группа ${g} не найдена в графике: практика и каникулы для неё не учитываются, теория идёт весь семестр. Проверьте название группы (поле подсвечено красным).`);
    }
  }
  $('swapWeeks').textContent = S.cal.swapWeeks
    ? '⇄ Недели 1, 3, 5… = правая колонка (изменено)' : '⇄ Недели 1, 3, 5… = левая колонка';
  if (S.schedule?.semester && period && S.schedule.semester !== period.idx + 1) {
    problems.push(`Загружено расписание ${S.schedule.semester} семестра, а выбран ${period.idx + 1} семестр — часы считаются по загруженному расписанию.`);
  }
  if (sem2 && S.prev?.fullName && S.teacher && S.prev.fullName.split(/\s+/)[0].toLowerCase() !== S.teacher.split(/\s+/)[0].toLowerCase()) {
    problems.push(`Ведомость 1 семестра составлена на «${S.prev.fullName}», а выбран педагог «${S.teacher}» — проверьте, тот ли файл.`);
  }
  showMessage(problems.join('\n'));
  renderTransfers();
  renderAnnual();
  refreshPracticeSummaries();
  if (!$('tab-plan').hidden) renderPlan();
  S.activeMonth = Math.min(S.activeMonth, Math.max(0, (S.result?.sheets.length || 1) - 1));
  refreshPreview();
}

function refreshPreview() {
  markRequired();
  const sheets = S.result?.sheets || [];
  $('download').disabled = !sheets.length;
  $('tabs').innerHTML = sheets.map((s, i) => `<button type="button" class="${i === S.activeMonth ? 'on' : ''}" data-i="${i}">${MONTH_NAMES[s.month]} ${s.year}</button>`).join('');
  for (const b of $('tabs').children) b.onclick = () => { S.activeMonth = Number(b.dataset.i); refreshPreview(); };
  $('preview').innerHTML = sheets.length
    ? renderSheet(meta(), sheets[S.activeMonth])
    : '<p class="hint">Предпросмотр появится, когда загружены график и расписание и выбран педагог.</p>';
}

const meta = () => ({
  college: $('college').value.trim(),
  academicYear: S.graph ? `${S.graph.startYear}/${S.graph.startYear + 1}` : '',
  teacherFullName: $('fullName').value.trim() || 'ФИО ПЕДАГОГА',
  specialty: $('specialty').value.trim(),
  deputy: $('deputy').value.trim(),
  extras: { dept: $('dept').value.trim(), cmk: $('cmk').value.trim(), position: $('position').value.trim() },
});

async function download() {
  if (!$('fullName').value.trim()) { $('fullName').classList.add('bad'); $('fullName').focus(); return; }
  const blob = await buildDocx(JSZip, meta(), S.result.sheets);
  const y = S.graph.startYear;
  const a = Object.assign(document.createElement('a'), {
    href: URL.createObjectURL(blob),
    download: `Учет_времени_${String(y).slice(2)}-${String(y + 1).slice(2)}_${S.teacher.split(/\s+/)[0]}_${Number($('semester').value) + 1}сем_${FUNDING.find(f => f.key === S.funding).file}.docx`,
  });
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 10000);
}

// ───────── вкладка «График» ─────────

function renderGraphTab() {
  if (!S.graph) return;
  const groups = [...new Set(S.rows.map(r => r.group))];
  renderGraph($('graphView'), S.graph, {
    highlight: highlightSet(groups), query: $('graphQuery').value,
    onlyHighlighted: $('graphMine').checked && groups.length > 0, semesters: S.graph.semesters(groups),
    fundingOf: key => S.graph.fundingOf(key, S.fundingOverrides),
    onFunding: key => {
      S.fundingOverrides[key] = S.graph.fundingOf(key, S.fundingOverrides) === 'grant' ? 'comm' : 'grant';
      markDirty(); renderGraphTab(); renderRows(); renderPracticeBlocks(); recompute();
    },
  });
}

// ───────── вкладка «Календарь» ─────────

function calendarStart() {
  return S.calStart || (S.graph ? ymOf(S.graph.weeks[0].start) : { y: new Date().getFullYear(), m: 0 });
}

function renderCalendarTab() {
  const calendar = makeCalendar();
  const start = calendarStart();
  $('workSat').checked = workSaturday();
  $('calStart').value = `${start.y}-${String(start.m + 1).padStart(2, '0')}`;
  renderCalendar($('calendarView'), {
    first: start, calendar,
    onToggle(date) {
      if (S.cal.overrides[date]) delete S.cal.overrides[date];
      else S.cal.overrides[date] = calendar.dayOff(date) ? { type: 'work' } : { type: 'off', name: 'Нерабочий день' };
      afterCalendarChanged();
    },
  });
}

function shiftCalendar(months) {
  const s = calendarStart();
  const i = s.y * 12 + s.m + months;
  S.calStart = { y: Math.floor(i / 12), m: i % 12 };
  markDirty();
  renderCalendarTab();
}

// Выгрузка/загрузка календаря: свои праздники, рабочие дни и переносы одним файлом, независимо от графика
function exportCalendar() {
  const data = { type: 'calendar', version: 1, overrides: S.cal.overrides, transfers: S.cal.transfers, workSaturday: S.cal.workSaturday };
  download_(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' }), 'календарь-рабочих-дней.json');
}

async function importCalendar(file) {
  if (!file) return;
  try {
    const d = JSON.parse(await file.text());
    if (d.type !== 'calendar' || typeof d.overrides !== 'object' || !Array.isArray(d.transfers)) throw new Error('Это не файл календаря');
    S.cal = { ...S.cal, overrides: d.overrides, transfers: d.transfers, workSaturday: d.workSaturday ?? null };
    afterCalendarChanged();
  } catch (e) { showMessage(`Календарь не загружен: ${e.message}`, true); }
}

function download_(blob, name) {
  const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: name });
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 10000);
}

function addDayTransfer() {
  const from = $('dtFrom').value;
  const to = $('dtTo').value;
  if (!from || !to) return;
  const calendar = makeCalendar();
  const name = $('dtName').value.trim() || calendar.info(from).off?.name || 'Перенос дня';
  S.cal.transfers.push({ from, to, name });
  $('dtFrom').value = $('dtTo').value = $('dtName').value = '';
  afterCalendarChanged();
}

function renderDayTransfers() {
  $('dtList').innerHTML = S.cal.transfers.map((t, i) =>
    `<li><span class="grow">${esc(t.name)}: выходной с <b>${ruDate(t.from)}</b> на <b>${ruDate(t.to)}</b></span><button type="button" data-i="${i}">✕</button></li>`).join('');
  $('dtList').querySelectorAll('button').forEach(b => {
    b.onclick = () => { S.cal.transfers.splice(Number(b.dataset.i), 1);  afterCalendarChanged(); };
  });
}

function afterCalendarChanged() {
  renderDayTransfers();
  renderCalendarTab();
  recompute();
}

// ───────── мелочи ─────────

function showMessage(text, isError = false) {
  const html = text ? text.split('\n').map(t => `<p class="${isError ? 'err' : ''}">${esc(t)}</p>`).join('') : '';
  $('messages').innerHTML = html;
  $('planMessages').innerHTML = html;
}

// kind — 'graph' | 'schedule' | 'load': подпись «ошибка» у файла (на обеих вкладках) и текст ошибки
function fail(kind, error) {
  setInfo(kind, 'ошибка', false);
  showMessage(error.message, true);
}

// ───────── рабочие области: сохранение хода работы ─────────

function markDirty() {
  if (S.loading) return;
  S.dirty = true;
  updateSaveState();
}

function updateSaveState() {
  const el = $('saveState');
  el.className = `savestate ${S.dirty ? 'dirty' : 'clean'}`;
  el.textContent = S.dirty ? '● Не сохранено' : (S.savedAt ? `✓ Сохранено ${new Date(S.savedAt).toLocaleTimeString('ru-RU')}` : 'Нет изменений');
  $('wsName').textContent = S.wsName || 'Без названия';
}

function snapshot() {
  stashTeacher();
  return {
    version: 1,
    year: $('year').value,
    files: { graph: S.files.graph, load: S.files.load, prev: S.files.prev },
    schedules: S.schedules.map(s => ({ file: s.file, sheet: s.sheet, buffer: s.buffer })),
    signers: S.signers,
    plan: S.plan,
    extraLoads: S.extraLoads,
    annual: S.annual,
    practices: S.practices.map(p => ({ file: p.file, buffer: p.buffer, teacher: p.teacher })),
    teacher: S.teacher,
    teacherStates: S.teacherStates,
    transfers: S.transfers,
    funding: S.funding,
    fundingOverrides: S.fundingOverrides,
    cal: S.cal,
    calStart: S.calStart,
    global: { college: $('college').value, deputy: $('deputy').value, carry: $('carry').checked },
  };
}

async function applySnapshot(d) {
  S.loading = true;
  try {
    $('year').value = d.year || '';
    if (d.files.graph) await loadGraphBuffer(d.files.graph.name, d.files.graph.buffer);
    S.practices = [];
    for (const p of d.practices || []) await addPracticeBuffer(p.file, p.buffer, p.teacher);
    S.schedules = [];
    // рабочие области версии 1.0 хранят одно расписание в files.schedule (первый лист)
    for (const s of d.schedules || (d.files.schedule ? [{ file: d.files.schedule.name, sheet: null, buffer: d.files.schedule.buffer }] : [])) {
      try { addScheduleBuffer(s.file, s.buffer, s.sheet ? { onlySheet: s.sheet } : { first: true }); } catch (e) { showMessage(e.message, true); }
    }
    S.signers = d.signers || S.signers;
    S.plan = { ...S.plan, ...(d.plan || {}) };
    if (d.files.load) await loadLoadBuffer(d.files.load.name, d.files.load.buffer);
    rebuildSchedule();
    S.teacherStates = d.teacherStates || {};
    S.extraLoads = d.extraLoads || {};
    S.annual = d.annual || { fullGen: '', collegeGen: '', notes: {} };
    if (d.files.prev) await loadPrevBuffer(d.files.prev.name, d.files.prev.buffer);
    S.transfers = d.transfers || {};
    S.funding = d.funding || 'grant';
    S.fundingOverrides = d.fundingOverrides || {};
    S.cal = d.cal || S.cal;
    S.calStart = d.calStart || null;
    $('college').value = d.global?.college ?? '';
    $('deputy').value = d.global?.deputy ?? '';
    $('carry').checked = d.global?.carry ?? true;
    renderDayTransfers();
    if (d.teacher) selectTeacher(d.teacher); else recompute();
  } finally { S.loading = false; }
}

async function saveWorkspace() {
  if (!S.wsId) {
    const suggested = `${S.teacher ? S.teacher.split(/\s+/)[0] : 'Ведомость'} ${S.graph ? `${S.graph.startYear}/${S.graph.startYear + 1}` : ''}`.trim();
    const name = prompt('Название рабочей области', suggested);
    if (!name) return;
    S.wsId = crypto.randomUUID();
    S.wsName = name.trim();
  }
  try {
    const updatedAt = Date.now();
    await putWorkspace({ id: S.wsId, name: S.wsName, updatedAt, data: snapshot() });
    S.dirty = false;
    S.savedAt = updatedAt;
    history.replaceState(null, '', `#ws=${S.wsId}`);
    updateSaveState();
  } catch (e) { showMessage(`Не удалось сохранить: ${e.message}`, true); }
}

function leaveConfirmed() {
  return !S.dirty || confirm('Есть несохранённые изменения. Продолжить без сохранения?');
}

function openWorkspace(id) {
  if (!leaveConfirmed()) return;
  S.dirty = false;
  location.hash = `#ws=${id}`;
  location.reload();
}

async function renderWorkspaceList() {
  let items = [];
  try { items = await listWorkspaces(); } catch (e) { $('wsList').innerHTML = `<li>${esc(e.message)}</li>`; return; }
  $('wsList').innerHTML = items.length ? items.map(w => `<li data-id="${w.id}">
    <span class="grow"><b>${esc(w.name)}</b>${w.id === S.wsId ? ' <em>(открыта)</em>' : ''}<br><small>${new Date(w.updatedAt).toLocaleString('ru-RU')}</small></span>
    <button type="button" data-act="open">Открыть</button>
    <button type="button" data-act="file" title="Скачать файл для переноса на другое устройство">В файл</button>
    <button type="button" data-act="del" title="Удалить">✕</button></li>`).join('')
    : '<li>Сохранённых рабочих областей пока нет.</li>';

  $('wsList').querySelectorAll('li[data-id]').forEach(li => {
    const id = li.dataset.id;
    li.querySelector('[data-act="open"]').onclick = () => openWorkspace(id);
    li.querySelector('[data-act="file"]').onclick = async () => {
      const ws = await getWorkspace(id);
      download_(new Blob([workspaceToJson(ws)], { type: 'application/json' }), `${ws.name}.workspace.json`);
    };
    li.querySelector('[data-act="del"]').onclick = async () => {
      if (!confirm('Удалить рабочую область безвозвратно?')) return;
      await deleteWorkspace(id);
      if (id === S.wsId) { S.wsId = null; S.wsName = ''; S.dirty = true; history.replaceState(null, '', location.pathname); updateSaveState(); }
      renderWorkspaceList();
    };
  });
}

async function importWorkspaceFile(file) {
  if (!file) return;
  try {
    const ws = workspaceFromJson(await file.text());
    const id = crypto.randomUUID();
    await putWorkspace({ ...ws, id, name: `${ws.name} (из файла)`, updatedAt: Date.now() });
    openWorkspace(id);
  } catch (e) { showMessage(`Файл не загружен: ${e.message}`, true); }
}

async function initWorkspaces() {
  $('wsBtn').onclick = () => { renderWorkspaceList(); $('wsDialog').showModal(); };
  $('wsClose').onclick = () => $('wsDialog').close();
  $('wsNew').onclick = () => { if (!leaveConfirmed()) return; S.dirty = false; history.replaceState(null, '', location.pathname); location.reload(); };
  $('wsImport').onchange = e => importWorkspaceFile(e.target.files[0]);
  $('saveBtn').onclick = saveWorkspace;
  $('calExport').onclick = exportCalendar;
  $('calImport').onchange = e => { importCalendar(e.target.files[0]); e.target.value = ''; };
  $('calStart').onchange = e => {
    const m = /^(\d{4})-(\d{2})$/.exec(e.target.value);
    if (m) { S.calStart = { y: Number(m[1]), m: Number(m[2]) - 1 }; markDirty(); renderCalendarTab(); }
  };
  $('calPrev').onclick = () => shiftCalendar(-12);
  $('calNext').onclick = () => shiftCalendar(12);

  for (const type of ['input', 'change']) document.addEventListener(type, e => { if (!e.target.closest('dialog')) markDirty(); });
  document.addEventListener('keydown', e => { if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); saveWorkspace(); } });
  window.addEventListener('beforeunload', e => { if (S.dirty) { e.preventDefault(); e.returnValue = ''; } });

  const id = /#ws=([\w-]+)/.exec(location.hash)?.[1];
  if (id) {
    try {
      const ws = await getWorkspace(id);
      if (!ws) throw new Error('рабочая область не найдена в этом браузере — загрузите её из файла');
      S.wsId = ws.id; S.wsName = ws.name;
      await applySnapshot(ws.data);
      S.dirty = false; S.savedAt = ws.updatedAt;
    } catch (e) { showMessage(`Рабочая область не открыта: ${e.message}`, true); }
  }
  S.dirty = false;
  updateSaveState();
}

// ───────── 2 семестр: ведомость 1 семестра ─────────

async function loadPrevFile(file) {
  if (file) await loadPrevBuffer(file.name, await file.arrayBuffer());
}

async function loadPrevBuffer(name, buffer) {
  try {
    S.prev = await parsePrevSemester(JSZip, buffer);
    S.files.prev = { name, buffer };
    $('prevInfo').textContent = `${name}: ${S.prev.fullName || 'без ФИО'}, строк: ${S.prev.rows.length}`;
    $('prevInfo').className = 'hint ok';
    if (!$('fullName').value && S.prev.fullName) $('fullName').value = S.prev.fullName;
    if (!$('specialty').value && S.prev.specialty) $('specialty').value = S.prev.specialty;
    recompute();
  } catch (e) { S.prev = null; S.files.prev = null; $('prevInfo').textContent = e.message; $('prevInfo').className = 'hint'; recompute(); }
}

// ───────── годовой учёт часов ─────────

function annualMeta() {
  return {
    yearLabel: S.graph ? `${S.graph.startYear}-${S.graph.startYear + 1}` : '',
    fullGen: S.annual.fullGen.trim(), collegeGen: S.annual.collegeGen.trim(),
    shortName: S.teacher, notes: S.annual.notes,
  };
}

function renderAnnual() {
  if ($('annualWrap').hidden || !S.resultAll || !S.prev) {
    if (!$('annualWrap').hidden) { $('annualTable').tBodies[0].innerHTML = ''; $('annualMissing').textContent = 'Загрузите ведомость 1 семестра (блок 2) — без неё годовой учёт собрать нельзя.'; $('annualDownload').disabled = true; }
    return;
  }
  const period = currentPeriod();
  S.entries = buildAnnualEntries({
    rows: S.rows, sheets: S.resultAll.sheets, prev: S.prev, sem2Start: period.from, extraLoads: S.extraLoads,
    defaultLoad: group => S.graph.fundingOf(group, S.fundingOverrides) || '',
  });
  if (!S.annual.collegeGen && $('college').value) { S.annual.collegeGen = genitiveCollege($('college').value); }
  if ($('collegeGen') !== document.activeElement) $('collegeGen').value = S.annual.collegeGen;
  if ($('fullGen') !== document.activeElement) $('fullGen').value = S.annual.fullGen;

  const body = $('annualTable').tBodies[0];
  body.innerHTML = S.entries.map((e, i) => `<tr data-i="${i}">
    <td>${e.kind === 'practice' ? 'практика' : 'теория'}${e.ref ? '' : ' <small>(только 1 сем.)</small>'}</td>
    <td>${esc(e.group)}</td><td>${esc(e.name)}</td>
    <td><select><option value="">— выбрать —</option>${LOAD_TYPES.map(t => `<option value="${t.key}"${t.key === e.load ? ' selected' : ''}>${esc(t.option)}</option>`).join('')}</select></td>
    <td>${e.plan ?? ''}</td><td>${e.fact}</td></tr>`).join('');
  body.querySelectorAll('tr').forEach(tr => {
    const e = S.entries[Number(tr.dataset.i)];
    const select = tr.querySelector('select');
    select.classList.toggle('need', !select.value);
    select.onchange = () => {
      if (e.ref) e.ref.load = select.value; else S.extraLoads[e.key] = select.value;
      markDirty();
      renderAnnual();
    };
  });

  // причина невыполнения — обязательна, если по типу нагрузки часов не добрано
  $('annualNotes').innerHTML = groupByLoad(S.entries).filter(g => g.notDone > 0).map(g => `<div class="noteline">
    <span>${esc(g.type.label)}: не выполнено ${g.notDone} ч — причина${STARHTML}</span>
    <input data-type="${g.type.key}" value="${esc(S.annual.notes[g.type.key] || '')}" placeholder="например: командировка"></div>`).join('');
  $('annualNotes').querySelectorAll('input').forEach(input => {
    input.classList.toggle('need', !input.value.trim());
    input.oninput = () => { S.annual.notes[input.dataset.type] = input.value; input.classList.toggle('need', !input.value.trim()); markDirty(); renderAnnualStatus(); };
  });
  $('fullGen').classList.toggle('need', !S.annual.fullGen.trim());
  $('collegeGen').classList.toggle('need', !S.annual.collegeGen.trim());
  renderAnnualStatus();
}

const STARHTML = ' <b class="star">*</b>';

function renderAnnualStatus() {
  const problems = annualProblems(S.entries, { fullGen: S.annual.fullGen, collegeGen: S.annual.collegeGen, notes: S.annual.notes });
  $('fullGen').classList.toggle('need', !S.annual.fullGen.trim());
  $('collegeGen').classList.toggle('need', !S.annual.collegeGen.trim());
  $('annualDownload').disabled = problems.length > 0;
  $('annualMissing').textContent = problems.length ? `Чтобы скачать, заполните: ${problems.slice(0, 6).join('; ')}${problems.length > 6 ? ` … и ещё ${problems.length - 6}` : ''}.` : 'Всё заполнено — документ можно скачать.';
}

async function downloadAnnual() {
  const blob = await buildAnnualDocx(JSZip, S.entries, annualMeta());
  const y = S.graph.startYear;
  download_(blob, `Годовой_учет_${String(y).slice(2)}-${String(y + 1).slice(2)}_${S.teacher.split(/\s+/)[0]}.docx`);
}

// ───────── грант / коммерция ─────────

// Финансирование строки: по группе (цвет в графике или правка). Неизвестное (нет группы в графике) видно в обоих режимах.
const fundingOfRow = row => (S.graph ? S.graph.fundingOf(row.group, S.fundingOverrides) : null);
const isVisible = row => { const f = fundingOfRow(row); return !f || f === S.funding; };
const visibleRows = () => S.rows.filter(isVisible);

function setFunding(key) {
  if (S.funding === key) return;
  S.funding = key;
  S.activeMonth = 0;
  markDirty();
  renderRows(); renderPracticeBlocks(); recompute();
}

function updateFundSwitch() {
  for (const bt of $('fundSwitch').querySelectorAll('button')) {
    const n = S.rows.filter(r => fundingOfRow(r) === bt.dataset.fund).length;
    bt.classList.toggle('on', bt.dataset.fund === S.funding);
    bt.querySelector('small').textContent = S.rows.length ? `(${n})` : '';
  }
}

// Сводка для помощника: состояние страницы без файлов целиком
function assistantContext() {
  const lines = [];
  lines.push(`Версия сайта: ${$('appVersion').textContent}. Режим: ${S.funding === 'grant' ? 'грант' : 'коммерция'}. Семестр: ${Number($('semester').value) + 1}.`);
  lines.push(`График: ${S.graph ? `${S.graph.weeks.length} нед., год ${S.graph.startYear}` : 'не загружен'}. Расписаний: ${S.schedules.length}${S.schedule ? `, групп ${S.schedule.groups.length}` : ''}. Нагрузка: ${S.load ? `${S.load.rows.length} строк` : 'нет'}. Документов практики: ${S.practices.length}.`);
  lines.push(`Педагог: ${S.teacher || 'не выбран'}.`);
  const warn = $('messages').innerText.trim();
  if (warn) lines.push(`Предупреждения на странице: ${warn.replace(/\s+/g, ' ').slice(0, 600)}`);
  for (const r of visibleRows().slice(0, 15)) {
    lines.push(`- ${r.kind === 'practice' ? 'практика' : 'теория'} ${r.group}: «${r.name.slice(0, 50)}», план ${r.plan ?? 'не задан'}${r.moduleIndex ? `, модуль ${r.moduleIndex}` : ''}`);
  }
  if (S.result) lines.push(`Часы по месяцам: ${S.result.sheets.map(s => `${s.name} ${s.total}`).join(', ')}. Занятий на праздниках без переноса: ${S.result.missed.length}.`);
  const empty = document.querySelectorAll('.need').length;
  if (empty) lines.push(`Пустых обязательных полей (красные): ${empty}.`);
  return lines.join('\n');
}
