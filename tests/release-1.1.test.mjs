// Проверка возможностей версии 1.1: несколько расписаний, нагрузка, дни практики, документ практики.
// Берёт документы из releases/<версия>/тест-документы/ (папка не публикуется); без неё — пропускает проверку.
// Запуск: RELEASE=1.1.0 node tests/release-1.1.test.mjs

import assert from 'node:assert/strict';
import { existsSync, readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const { DOMParser } = require('@xmldom/xmldom');
const JSZip = require('jszip');
const XLSX = require('xlsx');
globalThis.DOMParser = DOMParser;

const { releases } = JSON.parse(readFileSync(new URL('../releases.json', import.meta.url), 'utf8'));
const RELEASE = process.env.RELEASE || releases[0].version;
const BASE = new URL(`../releases/${RELEASE}/`, import.meta.url);
const DOCS = new URL('тест-документы/', BASE);
if (!existsSync(new URL('src/domain/schedules.js', BASE))) { console.log(`версия ${RELEASE}: новых возможностей нет — пропуск`); process.exit(0); }
if (!existsSync(new URL('1 График учебного процесса 2026-2027.xlsx', DOCS))) { console.log('нет папки тест-документы — пропуск'); process.exit(0); }

const load = async p => import(new URL(`src/${p}`, BASE).href);
const { readSheets } = await load('util/sheet.js');
const { parseGraph } = await load('parsers/graph.js');
const { parseSchedule } = await load('parsers/schedule.js');
const { parseLoad } = await load('parsers/load.js');
const { parsePractice } = await load('parsers/practice.js');
const { buildRows, teacherNames, emptyPracticeRow } = await load('domain/rows.js');
const { mergeSchedules, bellTable, buildGrid, parseBells, weeklySummary } = await load('domain/schedules.js');
const { matchLoad, applyLoad, loadCandidates, sameGroup, sameTitle } = await load('domain/load.js');
const { applyPattern, describe, documentDays, expandBlock, pairsText, parsePairsText, patternOf, setDay, totalHours } = await load('domain/practice-days.js');
const { buildPracticeDocx, shortSigner } = await load('export/practice-docx.js');
const { createCalendar } = await load('domain/calendar.js');
const { buildTimesheet } = await load('domain/timesheet.js');

const file = name => readFileSync(new URL(name, DOCS));
const ab = buf => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
const sheet0 = name => readSheets(XLSX, ab(file(name)))[0].sheet;
const holidays = JSON.parse(readFileSync(new URL('data/holidays-kz.json', BASE), 'utf8'));

// ───── расписания: время звонков, курс, семестр, год
const s2 = parseSchedule(sheet0('2 Расписание 2 курса 3 семестр 2026-2027.xlsx'));
const s3 = parseSchedule(sheet0('3 Расписание 3 курса 5 семестр 2026-2027.xlsx'));
assert.deepEqual([s2.course, s2.semesterNo, s2.semester, s2.startYear], [2, 3, 1, 2026]);
assert.deepEqual([s3.course, s3.semesterNo, s3.semester, s3.startYear], [3, 5, 1, 2026]);
assert.ok(s2.entries.every(e => /^\d\d:\d\d-\d\d:\d\d$/.test(e.time)), 'у каждого занятия есть время звонка');

// книга из нескольких листов: каждый лист-расписание находится сам
const whole = readSheets(XLSX, ab(file('доп/Книга Зиаятдинова целиком (все листы, 2026-27).xlsx')));
const found = whole.filter(({ sheet }) => { try { parseSchedule(sheet); return true; } catch { return false; } });
assert.equal(found.length, 2, 'в книге два листа с расписаниями: 2 курс за 2025-26 и 3 курс за 2026-27');

// ───── объединение расписаний
const merged = mergeSchedules([{ name: '2 курс', parsed: s2 }, { name: '3 курс', parsed: s3 }], { graphYear: 2026 });
assert.equal(merged.entries.length, s2.entries.length + s3.entries.length);
assert.equal(merged.semester, 1);
assert.deepEqual(merged.warnings, []);
const bells = bellTable(merged.entries);
assert.equal(bells.length, 7);
assert.equal(bells[0], '08:00-09:30');
assert.equal(bells[6], '18:30-20:00');
const old = parseSchedule(found.find(f => f.name === 'Рас-е 1 семестр').sheet);
const withOld = mergeSchedules([{ name: '2 курс', parsed: s2 }, { name: 'старое', parsed: old }, { name: '3 курс', parsed: s3 }], { graphYear: 2026 });
assert.ok(withOld.warnings.some(w => w.includes('2025–2026')), 'расписание другого года — предупреждение');
assert.ok(withOld.warnings.some(w => w.startsWith('Группа П2А есть в двух расписаниях')), 'одна группа в двух файлах — предупреждение, берётся поздний');
assert.deepEqual(parseBells('8:00-9:30\r\n09.40 – 11.10'), ['08:00-09:30', '09:40-11:10']);

// ───── педагог: сетка недели и пары в неделю
const teacher = 'Зиаятдинов В.Р.';
assert.ok(teacherNames(merged.entries).includes(teacher));
const summary = weeklySummary(merged.entries, teacher);
assert.equal(summary.length, 6);
assert.equal(summary.reduce((a, r) => a + r.left, 0), 19);
assert.equal(summary.find(r => r.group === 'П3У/Г').left, 4);
const mine = [...new Set(merged.entries.filter(e => e.teacher === teacher).map(e => e.group))];
const grid = buildGrid({ entries: merged.entries, bells, teacher, groups: mine, onlyMine: true });
const used = grid.rows.filter(r => r.used);
assert.ok(used.length > 0 && used.every(r => r.number >= 1));
assert.ok(grid.cells.left.flat(2).every(x => x.mine), 'в режиме «только пары педагога» чужих пар нет');
const wide = buildGrid({ entries: merged.entries, bells, teacher, groups: mine });
assert.ok(wide.cells.left.flat(2).some(x => !x.mine) && wide.cells.left.flat(2).some(x => x.mine), 'пары педагога подсвечены среди остальных');

// ───── нагрузка
const loadDoc = parseLoad(sheet0('4 Нагрузка Зиаятдинов 2026-2027.xlsx'));
assert.equal(loadDoc.rows.length, 15);
const pm10 = loadDoc.rows.find(r => r.group === 'П3А(д)' && r.moduleIndex === 'ПМ10');
assert.deepEqual([pm10.sems[0].weeks, pm10.sems[0].perWeek, pm10.sems[0].total], [3, 36, 108]);
assert.ok(sameGroup('П3У/Г', loadDoc.rows.find(r => r.group === 'П3У/Г').group));
assert.ok(sameGroup('П3А', 'П3А(д)'));

const graph = parseGraph(sheet0('1 График учебного процесса 2026-2027.xlsx'), { startYear: 2026 });
const { rows } = buildRows({ teacher, entries: merged.entries, graph, practiceDocs: [] });
assert.equal(applyLoad(rows, loadDoc, 0) >= 4, true);
const planOf = g => rows.find(r => r.group === g).plan;
assert.deepEqual(['ВТ2А', 'П3А', 'П3К', 'РОБ3А'].map(planOf), [72, 48, 48, 72]);
assert.equal(planOf('РОБ2А'), null, 'у РОБ2А в нагрузке две строки без модуля — выбирает пользователь, а не программа');
assert.equal(loadCandidates(loadDoc, 'РОБ2А', 0).length, 2);

// практика ↔ нагрузка: индекс в графике и нагрузке разный (ПМ6 / ПМ10), совпадает название
const mod = graph.moduleOf('П3А', 'ПМ6');
assert.ok(sameTitle(mod.title, pm10.title));
const prRow = { kind: 'practice', group: 'П3А', moduleIndex: 'ПМ6', name: `УП - ${mod.title}` };
assert.equal(matchLoad(prRow, rows, loadDoc, 0).row, pm10);

// ───── дни практики
assert.equal(pairsText([4, 5, 6]), '4-6 пара');
assert.equal(pairsText([1, 3]), '1, 3 пара');
assert.equal(pairsText([1, 2, 4, 5, 6]), '1-2, 4-6 пара');
assert.deepEqual(parsePairsText('1-2, 4-6 пара'), [1, 2, 4, 5, 6]);
assert.equal(parsePairsText('праздник'), null);
const cal = createCalendar({ holidays, state: { overrides: {}, transfers: [] }, workSaturday: true, years: [2025, 2028] });
const block = { from: '2026-11-16', to: '2026-11-28', days: {}, slots: {} };
applyPattern(block, { 1: [2, 3], 2: [4, 5, 6], 3: [1, 2, 3], 4: [1, 2, 3], 5: [], 6: [1, 2, 3] }, cal);
assert.equal(block.days['2026-11-16'], 4);                  // пн 2-3 пара
assert.equal(block.days['2026-11-17'], 6);                  // вт 4-6 пара
assert.equal(block.days['2026-11-20'], undefined, 'в пятницу практики нет');
assert.equal(totalHours([block]), 56);                       // две недели: пн 4 + вт 6 + ср 6 + чт 6 + сб 6 = 28 в каждой
assert.match(describe([block]), /^пн 2-3 пара, вт 4-6 пара, ср 1-3 пара, чт 1-3 пара, сб 1-3 пара · практики нет: пт/);
setDay(block, '2026-11-24', { pairs: [1, 2] });             // отдельный день иначе
assert.match(describe([block]), /иначе — 1-2 пара: 24\.11/);
assert.deepEqual(patternOf(block)[1], [2, 3]);
// простой блок «часов в день» разворачивается в дни; нерабочие дни календаря пропускаются
const simple = expandBlock({ from: '2026-12-14', to: '2026-12-19', hoursPerDay: 6, saturday: true }, cal);
assert.equal(Object.keys(simple.days).length, 5);          // 16 декабря — День Независимости, суббота считается
// строки документа: праздники на днях практики помечаются, выходные — нет
const rowsDoc = documentDays({ from: '2026-12-09', to: '2026-12-19', days: { '2026-12-09': 6, '2026-12-14': 6, '2026-12-15': 6, '2026-12-17': 6, '2026-12-18': 6, '2026-12-19': 6 }, slots: {} }, cal);
assert.ok(rowsDoc.some(d => d.date === '2026-12-16' && d.text === 'праздник' && d.hours === 0));
assert.ok(!rowsDoc.some(d => d.date === '2026-12-13'), 'воскресенье в документ не попадает');

// ведомость считает часы практики из дней с парами
const pr = emptyPracticeRow({ id: 'px', group: 'П3К', start: '2026-11-16' });
pr.moduleIndex = 'ПМ6'; pr.blocks = [block];
const sheets = buildTimesheet({ rows: [pr], graph, calendar: cal, from: { y: 2026, m: 10 }, to: { y: 2026, m: 10 }, carryOver: false }).sheets;
assert.equal(totalHours([block]), 54);                      // 24.11 — вторник: вместо 6 ч стало 4 ч (пары 1-2)
assert.equal(sheets[0].rows[0].total, 54);

// ───── документ практики: сформированный сайтом читается обратно, документ колледжа — тоже
const generated = await parsePractice(JSZip, file('5б Практика П3К ПМ6 2026-27 (сформирована приложением).docx'));
assert.equal(generated.group, 'П3К');
assert.equal(generated.moduleIndex, 'ПМ6');
assert.equal(generated.supervisor, 'В.Зиаятдинов');
assert.equal(generated.from, '2026-10-26');
assert.ok(Object.keys(generated.slots).length > 10);
assert.ok(Object.values(generated.days).includes(0), 'праздничный день остаётся с нулевыми часами');
const college = await parsePractice(JSZip, file('5а Практика П2А 2024-25 (чужой год, настоящий документ).docx'));
assert.equal(college.group, 'П2А');
assert.equal(college.total, 144);
assert.equal(Object.keys(college.slots).length, 21);
assert.deepEqual(college.slots['2025-01-03'], [3, 4, 5, 6]);
assert.deepEqual(college.signers, { approver: 'Н.Жаксыбаева', upr: 'А.Карсыбаева', head: 'А.Изекенова' });
assert.equal(college.specialtyOnly, '06130100 Программное обеспечение (по видам)');
assert.equal(college.qualification, '4S06130103 Разработчик программного обеспечения');
assert.equal(college.course, 2);
assert.equal(matchSupervisorShort(teacherNames(merged.entries), generated.supervisor), teacher);

const blob = await buildPracticeDocx(JSZip, {
  college: graph.college, yearLabel: '2026-2027', startYear: 2026, group: 'П3К', course: 3, specialty: '06130100', qualification: '4S06130103',
  moduleLine: 'ПМ6', supervisor: shortSigner('Зиаятдинов Владимир Рашидович'), approver: graph.deputy, upr: 'А.Карсыбаева', head: 'А.Изекенова',
}, [block, simple], cal);
mkdirSync(new URL('../tests/out/', import.meta.url), { recursive: true });
writeFileSync(new URL('../tests/out/practice.docx', import.meta.url), Buffer.from(await blob.arrayBuffer()));
const back = await parsePractice(JSZip, Buffer.from(await blob.arrayBuffer()));
assert.equal(back.from, '2026-11-16');
assert.equal(back.supervisor, 'В.Зиаятдинов');
assert.equal(shortSigner('Зиаятдинов В.Р.'), 'В.Зиаятдинов');

console.log('OK 1.1: расписаний', 2, 'звонков', bells.length, 'нагрузка', loadDoc.rows.length, 'практика', totalHours([block]), 'ч');

function matchSupervisorShort(names, supervisor) {
  const m = /^([^.]+)\.\s*(.+)$/.exec(supervisor || '');
  return m ? names.find(n => n.toLowerCase().startsWith(m[2].toLowerCase()) && n.split(/\s+/)[1]?.toLowerCase().startsWith(m[1].toLowerCase())) : null;
}
