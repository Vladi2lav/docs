// Годовой учёт часов: 1 семестр (из загруженной ведомости) + 2 семестр (расчёт) → записи по типам нагрузки.
// Ничего не угадывается: тип нагрузки, план, ФИО и название колледжа в родительном падеже должны быть заполнены.

import { LOAD_TYPES, PRACTICE_PREFIX } from '../config.js';
import { matchPrev } from './carryover.js';

// Столбцы формы — сентябрь … июнь (индексы месяцев Date): порядок учебного года
export const YEAR_MONTHS = [8, 9, 10, 11, 0, 1, 2, 3, 4, 5];
const pos = m => (m - 8 + 12) % 12;

const isPractice = name => name.trim().toLowerCase().startsWith(PRACTICE_PREFIX.trim().toLowerCase());
const sum = list => list.reduce((a, b) => a + b, 0);

// Записи для годовой формы. Строки 2 семестра дополняются часами 1 семестра из ведомости; строки, которые были
// только в 1 семестре, добавляются как есть. months: { индекс месяца: часы } за сентябрь–июнь.
export function buildAnnualEntries({ rows, sheets, prev, sem2Start, extraLoads = {} }) {
  const sem2Pos = pos(sem2Start.m);
  const used = new Set();
  const entries = rows.map((row, i) => {
    const p = prev && matchPrev(row, prev);
    if (p) used.add(p.key);
    const months = {};
    if (p) for (const [m, h] of Object.entries(p.months)) if (pos(Number(m)) < sem2Pos) months[m] = h;
    for (const sh of sheets) if (pos(sh.month) >= sem2Pos) months[sh.month] = sh.rows[i].total;
    return { id: row.id, ref: row, kind: row.kind, group: row.group, moduleIndex: row.moduleIndex, name: row.name, plan: row.plan, load: row.load || '', months, hasPrev: !!p };
  });
  for (const p of prev ? prev.rows : []) {
    if (used.has(p.key)) continue;
    const months = {};
    for (const [m, h] of Object.entries(p.months)) if (pos(Number(m)) < sem2Pos) months[m] = h;
    entries.push({ id: `prev:${p.key}`, ref: null, key: p.key, kind: isPractice(p.name) ? 'practice' : 'theory', group: p.group, moduleIndex: p.moduleIndex, name: p.name,
      plan: p.plan, load: extraLoads[p.key] || '', months, hasPrev: true });
  }
  for (const e of entries) e.fact = sum(YEAR_MONTHS.map(m => e.months[m] || 0));
  return entries;
}

// Страницы формы: по одной группе записей на тип нагрузки (в порядке LOAD_TYPES)
export function groupByLoad(entries) {
  return LOAD_TYPES.map(type => {
    const list = entries.filter(e => e.load === type.key);
    const plan = sum(list.map(e => e.plan || 0));
    const fact = sum(list.map(e => e.fact));
    return {
      type, entries: list, plan, fact,
      notDone: sum(list.map(e => Math.max(0, (e.plan || 0) - e.fact))),
      over: sum(list.map(e => Math.max(0, e.fact - (e.plan || 0)))),
    };
  }).filter(g => g.entries.length);
}

// Чего не хватает для формы (пустой список — документ можно получать)
export function annualProblems(entries, { fullGen, collegeGen, notes = {} }) {
  const out = [];
  if (!entries.length) out.push('нет дисциплин');
  for (const e of entries) {
    const title = `${e.group} ${e.name.slice(0, 28)}`;
    if (!e.load) out.push(`тип нагрузки: ${title}`);
    if (e.plan == null) out.push(`план часов: ${title}`);
    if (!e.hasPrev) out.push(`нет данных 1 семестра: ${title}`);
  }
  if (!fullGen.trim()) out.push('ФИО в родительном падеже');
  if (!collegeGen.trim()) out.push('колледж в родительном падеже');
  for (const g of groupByLoad(entries)) {
    if (g.notDone > 0 && !(notes[g.type.key] || '').trim()) out.push(`причина невыполнения (${g.type.label})`);
  }
  return out;
}

// «Центральноазиатский технико-экономический колледж» → «…ого … ого колледжа»: подсказка по типовому окончанию, поле редактируется
export function genitiveCollege(name) {
  return name.split(' ').map(w => {
    if (/^колледж$/i.test(w)) return `${w}а`;
    return /(ий|ый)$/i.test(w) ? `${w.slice(0, -2)}ого` : w;
  }).join(' ');
}
