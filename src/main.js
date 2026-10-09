// Точка входа: связывает загрузку документов → расчёт → предпросмотр → Word.
// Вся предметная логика лежит в parsers/, domain/, export/; здесь — состояние и обработчики интерфейса.

import { readSheet } from './util/sheet.js';
import { parseGraph, guessStartYear, normGroup, familyOf } from './parsers/graph.js';
import { PRACTICE_PREFIX } from './config.js';
import { parseSchedule } from './parsers/schedule.js';
import { parsePractice } from './parsers/practice.js';
import { fillRow } from './domain/autofill.js';
import { buildRows, makePracticeRow, matchSupervisor, teacherNames } from './domain/rows.js';
import { createCalendar } from './domain/calendar.js';
import { buildTimesheet } from './domain/timesheet.js';
import { buildDocx } from './export/docx.js';
import { esc, renderSheet } from './ui/preview.js';
import { renderGraph, highlightSet } from './ui/graph-view.js';
import { renderCalendar } from './ui/calendar-view.js';
import { ruDate, weekday, ymOf } from './util/dates.js';
import { MONTH_NAMES } from './config.js';

const $ = id => document.getElementById(id);
const S = {
  graphSheet: null, graph: null, schedule: null, holidays: null,
  practices: [],                 // { id, file, data, teacher }
  teacher: '', rows: [], warnings: [], rowReports: {},
  transfers: {},                 // педагог → [{ id, rowId, from, to, hours }]
  cal: { overrides: {}, transfers: [], workSaturday: null, swapWeeks: false },   // ручные правки календаря (до обновления страницы)
  result: null, activeMonth: 0,
};

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
  $('prGroup').onchange = fillPracticeModules;
  $('prAdd').onclick = addPracticeFromGraph;
  $('carry').onchange = recompute;
  $('swapWeeks').onclick = () => { S.cal.swapWeeks = !S.cal.swapWeeks;  recompute(); };
  $('download').onclick = download;
  $('trAdd').onclick = addTransfer;
  $('graphQuery').oninput = $('graphMine').onchange = renderGraphTab;
  $('workSat').onchange = () => { S.cal.workSaturday = $('workSat').checked;  afterCalendarChanged(); };
  $('dtAdd').onclick = addDayTransfer;
  renderDayTransfers();
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
  if (!file) return;
  try {
    S.graphSheet = readSheet(XLSX, await file.arrayBuffer());
    const guessed = guessStartYear(file.name);
    if (guessed) $('year').value = guessed;
    $('graphInfo').dataset.name = file.name;
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
  if (!file) return;
  try {
    S.schedule = parseSchedule(readSheet(XLSX, await file.arrayBuffer()));
    S.schedule.semester ||= Number(/([12])\s*сем/i.exec(file.name)?.[1] || 0);
    $('scheduleInfo').textContent = `${file.name}: ${S.schedule.groups.length} групп, ${teacherNames(S.schedule.entries).length} педагогов`;
    $('scheduleInfo').className = 'ok';
    $('teacher').innerHTML = '<option value="">— выберите педагога —</option>' +
      teacherNames(S.schedule.entries).map(n => `<option>${esc(n)}</option>`).join('');
    S.teacher = '';
    afterInputsChanged();
  } catch (e) { S.schedule = null; fail('scheduleInfo', e); }
}

async function loadPracticeFiles(files) {
  for (const file of files) {
    try {
      const data = await parsePractice(JSZip, await file.arrayBuffer());
      S.practices.push({ id: `${Date.now()}${S.practices.length}`, file: file.name, data, teacher: '' });
    } catch (e) { showMessage(`${file.name}: ${e.message}`, true); }
  }
  afterInputsChanged();
}

function afterInputsChanged() {
  if (S.schedule) {
    const names = teacherNames(S.schedule.entries);
    for (const p of S.practices) if (!p.teacher) p.teacher = matchSupervisor(names, p.data.supervisor) || '';
  }
  renderPractices();
  if (S.teacher) selectTeacher(S.teacher, { keepProfile: true });
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

function selectTeacher(name, { keepProfile = false } = {}) {
  S.teacher = name;
  $('teacher').value = name;
  S.rows = []; S.warnings = [];
  if (name && S.graph && S.schedule) {
    const practiceDocs = S.practices.filter(p => p.teacher === name).map(p => p.data);
    ({ rows: S.rows, warnings: S.warnings } = buildRows({ teacher: name, entries: S.schedule.entries, graph: S.graph, practiceDocs }));
    if (!keepProfile) {
      for (const id of ['fullName', 'specialty', 'dept', 'cmk', 'position']) $(id).value = '';
      $('semester').value = '0';
      S.rowReports = {};
    }
  }
  renderRows();
  renderPracticeBlocks();
  recompute();
}

function renderRows() {
  const body = $('rowsTable').tBodies[0];
  body.innerHTML = S.rows.map((r, i) => {
    const rep = S.rowReports[r.id];
    const status = rep ? [...rep.filled.map(t => `<span class="ok">✓ ${esc(t)}</span>`), ...rep.missing.map(t => `<span class="warn">⚠ ${esc(t)}</span>`)].join('<br>') : '';
    return `<tr class="head" data-h="${i}"><td colspan="6"><b>${esc(r.moduleIndex || 'модуль не выбран')} · ${r.kind === 'practice' ? 'практика' : 'теория'} · ${esc(r.group)}</b>
        <button type="button" data-fill>Заполнить по документам</button>
        <select data-module>${moduleOptions(r)}</select>${status ? `<div class="status">${status}</div>` : ''}</td></tr>
      <tr data-i="${i}">
        <td>${r.kind === 'practice' ? 'практика' : 'теория'}</td>
        <td><input data-f="moduleIndex" value="${esc(r.moduleIndex)}"></td>
        <td><input data-f="name" value="${esc(r.name)}"></td>
        <td><input data-f="group" value="${esc(r.group)}"></td>
        <td><input data-f="plan" type="number" min="0" value="${r.plan ?? ''}"></td>
        <td><button type="button" title="Убрать из ведомости">✕</button></td></tr>`;
  }).join('');

  body.querySelectorAll('tr.head').forEach(tr => {
    tr.querySelector('[data-fill]').onclick = () => fillOneRow(Number(tr.dataset.h));
    tr.querySelector('[data-module]').onchange = e => pickModule(Number(tr.dataset.h), e.target.value);
  });
  body.querySelectorAll('tr[data-i]').forEach(tr => {
    const row = S.rows[Number(tr.dataset.i)];
    tr.querySelectorAll('input').forEach(input => {
      input.oninput = () => {
        row[input.dataset.f] = input.dataset.f === 'plan' ? (input.value === '' ? null : Number(input.value)) : input.value;
        recompute();
      };
    });
    tr.querySelector('button').onclick = () => { S.rows.splice(Number(tr.dataset.i), 1); renderRows(); renderPracticeBlocks(); recompute(); };
  });
  const groups = [...new Set(S.rows.map(r => r.group))];
  $('prGroup').innerHTML = groups.map(g => `<option>${esc(g)}</option>`).join('');
  fillPracticeModules();
  $('prRow').hidden = !groups.length;
  $('trRow').innerHTML = S.rows.filter(r => r.kind === 'theory').map(r => `<option value="${r.id}">${esc(r.name)} (${esc(r.group)})</option>`).join('');
}

// ───────── выбор модуля из «Обозначений» графика (вручную, без догадок) ─────────

// Модули всех «семейств» графика; семейство группы строки идёт первым и отмечено
function moduleOptions(row) {
  if (!S.graph) return '<option value="">модуль из графика…</option>';
  const fam = familyOf(row.group);
  const mods = S.graph.legendModules();
  const groups = [...new Set(mods.map(m => m.family))].sort((a, b) => (b === fam) - (a === fam));
  return '<option value="">Модуль из графика — выбрать…</option>' + groups.map(f =>
    `<optgroup label="${esc(f)}${f === fam ? ' — семейство группы ' + esc(row.group) : ''}">${mods.filter(m => m.family === f).map(m =>
      `<option value="${esc(f)}|${esc(m.code)}"${row.moduleRef === `${f}|${m.code}` ? ' selected' : ''}>${esc(m.index)} ${esc(m.title)}</option>`).join('')}</optgroup>`).join('');
}

function pickModule(i, value) {
  const row = S.rows[i];
  const [family, code] = value.split('|');
  const mod = value && S.graph.legend[family] && S.graph.legend[family][code];
  if (!mod) return;
  row.moduleRef = value;
  row.moduleIndex = mod.index;
  row.name = row.kind === 'practice' ? PRACTICE_PREFIX + mod.title : mod.title;
  renderRows(); renderPracticeBlocks(); recompute();
}

function fillPracticeModules() {
  const group = $('prGroup').value;
  const legend = S.graph && group ? S.graph.legendFor(group) : null;
  $('prModule').innerHTML = legend
    ? Object.entries(legend).map(([code, m]) => `<option value="${esc(code)}">${esc(m.index)} ${esc(m.title)}</option>`).join('')
    : '<option value="">в графике нет «Обозначений» для этой группы</option>';
  $('prAdd').disabled = !legend;
}

function addPracticeFromGraph() {
  const group = $('prGroup').value;
  const code = $('prModule').value;
  const legend = S.graph && S.graph.legendFor(group);
  if (!legend || !legend[code]) return;
  const module = { code, ...legend[code] };
  const row = makePracticeRow({ id: `p${Date.now()}`, group, module, graph: S.graph });
  if (!row.blocks.length) { showMessage(`В графике у группы ${group} нет недель с кодом «${code}» (${module.index}).`, true); return; }
  S.rows.push(row);
  renderRows(); renderPracticeBlocks(); recompute();
}

// ───────── «Заполнить по документам» ─────────

function fillOneRow(i) {
  const row = S.rows[i];
  S.rowReports[row.id] = fillRow(row, { graph: S.graph });
  renderRows(); renderPracticeBlocks(); recompute();
}

function fillAll() {
  const ok = [];
  const bad = [];
  if (!S.graph || !S.schedule) {
    bad.push(`${!S.graph ? 'график учебного процесса' : 'расписание занятий'} не загружен (шаг 1) — без него заполнять нечего`);
    return showReport(ok, bad);
  }

  if (!S.teacher) {
    bad.push('педагог: не выбран — выберите в списке (блок 2), затем нажмите кнопку снова');
    return showReport(ok, bad);
  }

  const practiceSpecialty = S.practices.filter(p => p.teacher === S.teacher).map(p => p.data.specialty).find(Boolean);
  const attrs = [
    ['fullName', 'ФИО полностью', '', 'в расписании только фамилия с инициалами — введите полное ФИО в блоке 2'],
    ['specialty', 'специальность', practiceSpecialty,
      `в графике специальности не привязаны к группам (вариантов: ${S.graph.specialties.length}), в docx практики её нет — выберите в поле «Специальность» (блок 2)`],
    ['college', 'название колледжа', S.graph.college, 'в графике не найдена строка «Учреждение образования «…»» — введите в блоке 2'],
    ['deputy', 'подпись зам. директора', S.graph.deputy, 'в графике нет строки «Заместитель директора по учебной работе … Фамилия» — введите в блоке 2'],
  ];
  for (const [id, label, value, hint] of attrs) {
    if (value) { $(id).value = value; ok.push(`${label}: ${value}`); } else bad.push(`${label}: ${hint}`);
  }

  S.rows.forEach(row => {
    const r = fillRow(row, { graph: S.graph });
    S.rowReports[row.id] = r;
    r.missing.forEach(t => bad.push(`${row.moduleIndex || 'дисциплина'} ${row.group} (${row.kind === 'practice' ? 'практика' : 'теория'}): ${t}`));
  });
  ok.push(`дисциплин заполнено по документам: ${S.rows.length}`);
  renderRows(); renderPracticeBlocks(); recompute();
  showReport(ok, bad);
}

function showReport(ok, bad) {
  $('fillReport').innerHTML = [...ok.map(t => `<p class="okline">✓ ${esc(t)}</p>`), ...bad.map(t => `<p class="warnline">⚠ Не хватает — ${esc(t)}</p>`)].join('');
}

// Блоки практики: период × часов в день (из графика) либо точные часы по датам (из docx)
function renderPracticeBlocks() {
  const rows = S.rows.filter(r => r.kind === 'practice');
  $('practiceBlocks').innerHTML = rows.length ? rows.map(r => `<div class="block" data-id="${r.id}">
    <b>${esc(r.name)}</b> <small>(${esc(r.group)})</small>
    ${r.blocks.map((b, i) => b.days
      ? `<div class="row" data-b="${i}">по документу: ${ruDate(b.from)} – ${ruDate(b.to)} <button type="button" data-act="del">✕</button></div>`
      : `<div class="row" data-b="${i}">с <input type="date" data-f="from" value="${b.from}"> по <input type="date" data-f="to" value="${b.to}">
          <input type="number" data-f="hoursPerDay" min="0" value="${b.hoursPerDay}" style="width:70px"> ч/день
          <label class="inline"><input type="checkbox" data-f="saturday"${b.saturday ? ' checked' : ''}> и суббота</label>
          <button type="button" data-act="del">✕</button></div>`).join('')}
    <button type="button" data-act="add">+ период</button></div>`).join('')
    : '<p class="hint">Практика не добавлена: нет docx практики с этим руководителем. Добавьте вручную кнопкой «+ Практика по графику» ниже.</p>';

  $('practiceBlocks').querySelectorAll('.block').forEach(box => {
    const row = S.rows.find(r => r.id === box.dataset.id);
    box.querySelectorAll('.row').forEach(line => {
      const b = row.blocks[Number(line.dataset.b)];
      line.querySelectorAll('input').forEach(input => {
        input.oninput = () => {
          b[input.dataset.f] = input.type === 'checkbox' ? input.checked : (input.dataset.f === 'hoursPerDay' ? Number(input.value) : input.value);
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

function makeCalendar() {
  return createCalendar({
    holidays: S.holidays, state: S.cal, workSaturday: workSaturday(),
    years: S.graph ? [S.graph.startYear - 1, S.graph.startYear + 2] : [2025, 2028],
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
  if (S.rows.some(r => r.plan === null)) problems.push('План часов не задан у части дисциплин: в Word колонки «Запланировано» и «Остаток» останутся пустыми. Введите план в блоке 3.');
  if (S.rows.some(r => r.kind === 'practice' && r.blocks.some(b => !b.days && !(Number(b.hoursPerDay) > 0)))) problems.push('У практики не указаны часы в день — такие периоды дают 0 часов. Заполните блок «Практика».');
  showMessage(problems.join('\n'));
  renderTransfers();
  S.activeMonth = Math.min(S.activeMonth, Math.max(0, (S.result?.sheets.length || 1) - 1));
  refreshPreview();
}

function refreshPreview() {
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

function renderCalendarTab() {
  const calendar = makeCalendar();
  $('workSat').checked = workSaturday();
  if (!S.graph) {
    $('calendarView').innerHTML = '<p class="hint">Загрузите график учебного процесса на вкладке «Ведомость»: по нему определяется учебный год.</p>';
    return;
  }
  renderCalendar($('calendarView'), {
    first: ymOf(S.graph.weeks[0].start), calendar,
    onToggle(date) {
      if (S.cal.overrides[date]) delete S.cal.overrides[date];
      else S.cal.overrides[date] = calendar.dayOff(date) ? { type: 'work' } : { type: 'off', name: 'Нерабочий день' };
      afterCalendarChanged();
    },
  });
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
