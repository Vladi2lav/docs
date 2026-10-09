// Сквозная проверка на реальных документах из исходники/: парсеры → расчёт → .docx.
// Запуск: npm install && npm test

import assert from 'node:assert/strict';
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const { DOMParser } = require('@xmldom/xmldom');
const JSZip = require('jszip');
const XLSX = require('xlsx');

globalThis.DOMParser = DOMParser;

import { readSheet } from '../src/util/sheet.js';
import { parseGraph } from '../src/parsers/graph.js';
import { parseSchedule } from '../src/parsers/schedule.js';
import { parsePractice } from '../src/parsers/practice.js';
import { buildRows, matchSupervisor, teacherNames } from '../src/domain/rows.js';
import { createCalendar } from '../src/domain/calendar.js';
import { buildTimesheet } from '../src/domain/timesheet.js';
import { buildDocx } from '../src/export/docx.js';
import { fillRow } from '../src/domain/autofill.js';

const src = name => readFileSync(new URL(`../исходники/${name}`, import.meta.url));
const holidays = JSON.parse(readFileSync(new URL('../data/holidays-kz.json', import.meta.url), 'utf8'));

const graph = parseGraph(readSheet(XLSX, src('График учебного процесса 2026-2027  14 09.xlsx')), { startYear: 2026 });
const schedule = parseSchedule(readSheet(XLSX, src('2 курс 1 семестр 26-27 13.09.xlsx')));

// график
assert.equal(graph.weeks[0].start, '2026-09-01');
assert.equal(graph.weeks.find(w => w.num === 5).start, '2026-09-28');
assert.equal(graph.parityAt('2026-09-08'), 'right');
assert.equal(graph.codeAt('П2А', '2026-09-08'), '');          // теория
assert.ok(graph.codeAt('П2А', '2026-11-03'));                  // практика
assert.equal(graph.moduleOf('П2А', 'ПМ4').title.startsWith('Разработка десктопных'), true);
console.log('график:', graph.college, '|', graph.deputy, '|', graph.semesters(['П2А']));

// расписание
const names = teacherNames(schedule.entries);
assert.ok(names.includes('Жаксыбаева Н.Н.'));
console.log('расписание:', schedule.entries.length, 'занятий,', names.length, 'педагогов, суббота:', schedule.hasSaturday);

// практика: docx из исходники/ (2024 год — вне учебного года, поэтому берутся периоды из графика)
const practiceDocs = [await parsePractice(JSZip, src('+2024 П2А УП Прог 1 семестр.docx'))];
assert.equal(practiceDocs[0].total, 144);
assert.equal(matchSupervisor(names, practiceDocs[0].supervisor), 'Жаксыбаева Н.Н.');

// расчёт и Word
const teacher = 'Жаксыбаева Н.Н.';
const { rows, warnings } = buildRows({ teacher, entries: schedule.entries, graph, practiceDocs });
console.log('предупреждения:', warnings);
const practice = rows.find(r => r.kind === 'practice');
assert.ok(practice, 'практика берётся из docx, где руководитель назван прямо');
assert.ok(warnings.some(w => w.includes('другой год')), 'документ 2024 года должен применяться с предупреждением');
assert.ok(practice.blocks.some(b => b.fromDoc), 'часы из документа перенесены на период графика');
console.log('практика из графика:', practice.name.slice(0, 40), practice.blocks);

const sems = graph.semesters(rows.map(r => r.group));
console.log('семестры:', JSON.stringify(sems));
assert.equal(sems.length, 2);
assert.deepEqual(sems[0].to, { y: 2027, m: 0 });

// без догадок: теория не получает индекс модуля и практику сама
assert.equal(rows[0].moduleIndex, '');
assert.equal(rows.filter(r => r.kind === 'practice').length, 1);
assert.equal(buildRows({ teacher: 'Зеленин В.А.', entries: schedule.entries, graph, practiceDocs: [] }).rows.filter(r => r.kind === 'practice').length, 0);

const calendar = createCalendar({ holidays, state: { overrides: {}, transfers: [] }, workSaturday: false, years: [2025, 2028] });
const theory = rows[0];
const transfers = [{ rowId: theory.id, from: '2026-09-01', to: '2026-09-05', hours: 2 }];
// «Заполнить по документам»: плана в документах нет — он вводится вручную (эталон Жаксыбаевой: 96 / 264)
const report = rows.map(r => fillRow(r, { graph }));
assert.ok(report.every(r => r.missing.some(t => t.startsWith('план часов'))));
assert.equal(buildTimesheet({ rows, graph, calendar, from: sems[0].from, to: sems[0].to, carryOver: true }).sheets[0].rows[0].remainingText, '');   // плана нет — остаток не выдумываем
rows[0].plan = 96;
rows[1].plan = 264;

const { sheets, missed } = buildTimesheet({ rows, graph, calendar, from: sems[0].from, to: sems[0].to, carryOver: true, transfers });
const sep = sheets[0];
console.log(sep.name, sep.rows.map(r => [r.name.slice(0, 30), r.total, r.plan, r.remainingText]));
assert.equal(sep.rows[0].remainingText, '68');
assert.equal(sep.rows[0].total, 28);                       // эталон «Жаксыбаева», сентябрь: 28 часов (перенос внутри месяца)
assert.equal(sep.rows[0].cells[4].text, '2 отраб за 01.09');   // 5 сентября — подпись отработки
assert.equal(sep.rows[0].cells[0].text, '2');              // 1 сентября: было 4, перенесено 2
console.log('сорванные занятия:', missed.map(m => m.date + ' ' + m.reason));
const dec = sheets.find(s => s.month === 11);
assert.ok(dec.rows[1].total > 0, 'практика в декабре');
console.log('практика декабрь/январь:', dec.rows[1].total, sheets[4].rows[1].total, 'итого', sheets.map(s => s.total));

const blob = await buildDocx(JSZip, {
  college: graph.college, academicYear: '2026/2027', teacherFullName: 'Жаксыбаева Наталья Николаевна',
  specialty: graph.specialties[0] || '', deputy: graph.deputy,
}, sheets);
mkdirSync(new URL('../tests/out/', import.meta.url), { recursive: true });
writeFileSync(new URL('../tests/out/result.docx', import.meta.url), Buffer.from(await blob.arrayBuffer()));
console.log('OK, месяцев в документе:', sheets.length);
