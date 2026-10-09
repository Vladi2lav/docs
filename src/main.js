// Точка входа: связывает загрузку документов → расчёт → предпросмотр → Word.
// Вся предметная логика лежит в parsers/, domain/, export/; здесь — состояние и обработчики интерфейса.

import { readSheet } from './util/sheet.js';
import { parseGraph, guessStartYear, normGroup } from './parsers/graph.js';
import { PRACTICE_PREFIX } from './config.js';
import { parseSchedule } from './parsers/schedule.js';
import { parsePractice } from './parsers/practice.js';
import { buildRows, emptyPracticeRow, graphPeriods, matchSupervisor, teacherNames } from './domain/rows.js';
import { createCalendar } from './domain/calendar.js';
import { buildTimesheet } from './domain/timesheet.js';
import { buildDocx } from './export/docx.js';
import { esc, renderSheet } from './ui/preview.js';
import { renderGraph, highlightSet } from './ui/graph-view.js';
import { renderCalendar } from './ui/calendar-view.js';
import { ruDate, weekday, ymOf } from './util/dates.js';
import { MONTH_NAMES } from './config.js';
import { deleteWorkspace, getWorkspace, listWorkspaces, putWorkspace, workspaceFromJson, workspaceToJson } from './util/workspaces.js';

const $ = id => document.getElementById(id);
const S = {
  graphSheet: null, graph: null, schedule: null, holidays: null,
  practices: [],                 // { id, file, data, teacher }
  teacher: '', rows: [], warnings: [],
  transfers: {},                 // педагог → [{ id, rowId, from, to, hours }]
  cal: { overrides: {}, transfers: [], workSaturday: null, swapWeeks: false },   // ручные правки календаря (до обновления страницы)
  result: null, activeMonth: 0,
  files: { graph: null, schedule: null },   // исходные файлы (имя + содержимое) — нужны для сохранения рабочей области
  teacherStates: {},             // педагог → { rows, warnings, fields, semester }: правки, сделанные для каждого педагога
  calStart: null,                // первый месяц, показанный в календаре ({ y, m })
  wsId: null, wsName: '', dirty: false, loading: false, savedAt: null,
};
const FIELD_IDS = ['fullName', 'specialty', 'dept', 'cmk', 'position'];

init();

async function init() {
  resetForm();
  S.holidays = await (await fetch('data/holidays-kz.json')).json();

  document.querySelectorAll('#nav button').forEach(b => { b.onclick = () => showTab(b.dataset.tab); });
  $('graphFile').onchange = e => loadGraphFile(e.target.files[0]);
  $('scheduleFile').onchange = e => loadScheduleFile(e.target.files[0]);
  $('practiceFile').onchange = e => loadPracticeFiles([...e.target.files]).then(() => { e.target.value = ''; });
  $('year').onchange = reparseGraph;
  $('teacher').onchange = e => selectTeacher(e.target.value);
  $('semester').onchange = recompute;
  $('fullName').oninput = () => { $('fullName').classList.remove('bad'); refreshPreview(); };
  for (const id of ['specialty', 'college', 'deputy', 'dept', 'cmk', 'position']) $(id).oninput = refreshPreview;
  $('fillAll').onclick = fillAll;
  $('prAdd').onclick = addPractice;
  $('carry').onchange = recompute;
  $('swapWeeks').onclick = () => { S.cal.swapWeeks = !S.cal.swapWeeks;  recompute(); };
  $('download').onclick = download;
  $('trAdd').onclick = addTransfer;
  $('graphQuery').oninput = $('graphMine').onchange = renderGraphTab;
  $('workSat').onchange = () => { S.cal.workSaturday = $('workSat').checked;  afterCalendarChanged(); };
  $('dtAdd').onclick = addDayTransfer;
  renderDayTransfers();
  initWorkspaces();
}

// После обновления страницы всё начинается с чистого листа (браузер иначе подставляет прежние значения полей)
function resetForm() {
  document.querySelectorAll('input, select').forEach(el => {
    if (el.type === 'file') el.value = '';
    else if (el.type === 'checkbox') el.checked = el.defaultChecked;
    else if (el.tagName === 'SELECT') el.selectedIndex = 0;
    else el.value = el.defaultValue;
  });
}

function showTab(name) {
  document.querySelectorAll('.tab').forEach(t => { t.hidden = t.id !== `tab-${name}`; });
  document.querySelectorAll('#nav button').forEach(b => b.classList.toggle('on', b.dataset.tab === name));
  if (name === 'graph') renderGraphTab();
  if (name === 'calendar') renderCalendarTab();
}

// ───────── загрузка документов ─────────

async function loadGraphFile(file) {
  if (file) await loadGraphBuffer(file.name, await file.arrayBuffer());
}

async function loadGraphBuffer(name, buffer) {
  try {
    S.graphSheet = readSheet(XLSX, buffer);
    S.files.graph = { name, buffer };
    const guessed = guessStartYear(name);
    if (guessed && !S.loading) $('year').value = guessed;
    $('graphInfo').dataset.name = name;
    reparseGraph();
  } catch (e) { fail('graphInfo', e); }
}

function reparseGraph() {
  if (!S.graphSheet) return;
  try {
    S.graph = parseGraph(S.graphSheet, { startYear: Number($('year').value) });
    const g = S.graph;
    $('graphInfo').textContent = `${$('graphInfo').dataset.name}: ${g.weeks.length} нед., ${Object.keys(g.groups).length} групп`;
    $('graphInfo').className = 'ok';
    if (g.college) $('college').value = g.college;
    if (g.deputy) $('deputy').value = g.deputy;
    $('specialties').innerHTML = g.specialties.map(s => `<option value="${esc(s)}">`).join('');
    afterInputsChanged();
  } catch (e) { S.graph = null; fail('graphInfo', e); }
}

async function loadScheduleFile(file) {
  if (file) await loadScheduleBuffer(file.name, await file.arrayBuffer());
}

async function loadScheduleBuffer(name, buffer) {
  try {
    S.schedule = parseSchedule(readSheet(XLSX, buffer));
    S.files.schedule = { name, buffer };
    S.schedule.semester ||= Number(/([12])\s*сем/i.exec(name)?.[1] || 0);
    $('scheduleInfo').textContent = `${name}: ${S.schedule.groups.length} групп, ${teacherNames(S.schedule.entries).length} педагогов`;
    $('scheduleInfo').className = 'ok';
    $('teacher').innerHTML = '<option value="">— выберите педагога —</option>' +
      teacherNames(S.schedule.entries).map(n => `<option>${esc(n)}</option>`).join('');
    S.teacher = '';
    afterInputsChanged();
  } catch (e) { S.schedule = null; fail('scheduleInfo', e); }
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
  S.teacher = name;
  $('teacher').value = name;
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
    }
  }
  renderRows();
  renderPracticeBlocks();
  recompute();
}

function renderRows() {
  const body = $('rowsTable').tBodies[0];
  body.innerHTML = S.rows.map((r, i) => {
    const legend = S.graph && S.graph.legendFor(r.group);
    const mods = legend ? Object.values(legend) : [];
    const known = mods.some(m => m.index === r.moduleIndex);
    const indexCell = legend
      ? `<select data-f="moduleIndex"><option value="">— без модуля —</option>${r.moduleIndex && !known ? `<option value="${esc(r.moduleIndex)}" selected>${esc(r.moduleIndex)}</option>` : ''}${mods.map(m =>
          `<option value="${esc(m.index)}"${m.index === r.moduleIndex ? ' selected' : ''}>${esc(m.index)} ${esc(m.title)}</option>`).join('')}</select>`
      : `<input data-f="moduleIndex" value="${esc(r.moduleIndex)}" placeholder="нет в графике">`;
    return `<tr data-i="${i}">
      <td>${r.kind === 'practice' ? 'практика' : 'теория'}</td>
      <td>${indexCell}</td>
      <td><input data-f="name" value="${esc(r.name)}"></td>
      <td><input data-f="group" value="${esc(r.group)}"></td>
      <td><input data-f="plan" type="number" min="0" value="${r.plan ?? ''}"></td>
      <td><button type="button" title="Убрать из ведомости">✕</button></td></tr>`;
  }).join('');

  body.querySelectorAll('tr[data-i]').forEach(tr => {
    const row = S.rows[Number(tr.dataset.i)];
    const nameInput = tr.querySelector('[data-f="name"]');
    tr.querySelectorAll('input[data-f]').forEach(input => {
      input.oninput = () => {
        row[input.dataset.f] = input.dataset.f === 'plan' ? (input.value === '' ? null : Number(input.value)) : input.value;
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
    tr.querySelector('button').onclick = () => { S.rows.splice(Number(tr.dataset.i), 1); renderRows(); renderPracticeBlocks(); recompute(); };
  });
  $('trRow').innerHTML = S.rows.filter(r => r.kind === 'theory').map(r => `<option value="${r.id}">${esc(r.name)} (${esc(r.group)})</option>`).join('');
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
  for (const id of ['fullName', 'specialty', 'college', 'deputy']) $(id).classList.toggle('need', ready && !$(id).value.trim());
  $('rowsTable').querySelectorAll('tr[data-i]').forEach(tr => {
    const row = S.rows[Number(tr.dataset.i)];
    tr.querySelectorAll('input, select').forEach(el => {
      const optional = el.dataset.f === 'moduleIndex' && row.kind === 'theory';   // у теории (например, физкультура) модуля может не быть
      el.classList.toggle('need', ready && !optional && el.value.trim() === '');
    });
  });
  $('practiceBlocks').querySelectorAll('input[data-f="hoursPerDay"]').forEach(input => {
    input.classList.toggle('need', !(Number(input.value) > 0));
  });
}

// Блоки практики: период × часов в день либо точные часы по датам (из docx)
function renderPracticeBlocks() {
  const rows = S.rows.filter(r => r.kind === 'practice');
  $('practiceBlocks').innerHTML = rows.map(r => `<div class="block" data-id="${r.id}">
    <b>${esc(r.name)}</b> <small>(${esc(r.group)})</small>
    ${r.blocks.map((b, i) => b.days
      ? `<div class="row" data-b="${i}">по документу: ${ruDate(b.from)} – ${ruDate(b.to)} <button type="button" data-act="del">✕</button></div>`
      : `<div class="row" data-b="${i}">с <input type="date" data-f="from" value="${b.from}"> по <input type="date" data-f="to" value="${b.to}">
          <input type="number" data-f="hoursPerDay" min="0" value="${b.hoursPerDay}" style="width:70px"> ч/день <b class="star">*</b>
          <label class="inline"><input type="checkbox" data-f="saturday"${b.saturday ? ' checked' : ''}> и суббота</label>
          <button type="button" data-act="del">✕</button></div>`).join('')}
    <div class="row"><button type="button" data-act="add">+ период</button>
      <button type="button" data-act="graph">Периоды из графика</button></div></div>`).join('');

  $('practiceBlocks').querySelectorAll('.block').forEach(box => {
    const row = S.rows.find(r => r.id === box.dataset.id);
    box.querySelectorAll('.row[data-b]').forEach(line => {
      const b = row.blocks[Number(line.dataset.b)];
      line.querySelectorAll('input').forEach(input => {
        input.oninput = () => {
          b[input.dataset.f] = input.type === 'checkbox' ? input.checked : (input.dataset.f === 'hoursPerDay' ? input.value : input.value);
          recompute();
        };
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
  });
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
  const problems = [...S.warnings];
  const period = currentPeriod();
  $('workSat').checked = workSaturday();

  if (S.graph && S.rows.length && period) {
    S.result = buildTimesheet({
      rows: S.rows, graph: S.graph, calendar: makeCalendar(), from: period.from, to: period.to,
      carryOver: $('carry').checked && period.idx === 0, transfers: transfersOf(), swapWeeks: !!S.cal.swapWeeks,
    });
    $('rowsTable').tBodies[0].querySelectorAll('input[data-f="plan"]').forEach((input, i) => { input.placeholder = `по расписанию ${S.result.totals[i]}`; });
    for (const g of new Set(S.rows.map(r => r.group))) {
      if (!S.graph.groups[normGroup(g)]) problems.push(`Группа ${g} не найдена в графике: практика и каникулы для неё не учитываются.`);
    }
  }
  $('swapWeeks').textContent = S.cal.swapWeeks
    ? '⇄ Недели 1, 3, 5… = правая колонка (изменено)' : '⇄ Недели 1, 3, 5… = левая колонка';
  if (S.schedule?.semester && period && S.schedule.semester !== period.idx + 1) {
    problems.push(`Загружено расписание ${S.schedule.semester} семестра, а выбран ${period.idx + 1} семестр — часы считаются по загруженному расписанию.`);
  }
  showMessage(problems.join('\n'));
  renderTransfers();
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
    download: `Учет_времени_${String(y).slice(2)}-${String(y + 1).slice(2)}_${S.teacher.split(/\s+/)[0]}_${Number($('semester').value) + 1}сем.docx`,
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
  $('messages').innerHTML = text ? text.split('\n').map(t => `<p class="${isError ? 'err' : ''}">${esc(t)}</p>`).join('') : '';
}

function fail(infoId, error) {
  $(infoId).textContent = 'ошибка';
  $(infoId).className = '';
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
    files: { graph: S.files.graph, schedule: S.files.schedule },
    practices: S.practices.map(p => ({ file: p.file, buffer: p.buffer, teacher: p.teacher })),
    teacher: S.teacher,
    teacherStates: S.teacherStates,
    transfers: S.transfers,
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
    if (d.files.schedule) await loadScheduleBuffer(d.files.schedule.name, d.files.schedule.buffer);
    S.teacherStates = d.teacherStates || {};
    S.transfers = d.transfers || {};
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
