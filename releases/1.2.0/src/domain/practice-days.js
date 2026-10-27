// Дни практики: какие пары в какой день. Хранится в блоке практики двумя словарями по датам:
//   days  — часы в день (их читает расчёт ведомости),
//   slots — номера пар в день (необязательно: у практики из старого документа или «ч/день» пар нет).
// Часы дня с парами всегда = число пар × HOURS_PER_PAIR.

import { HOURS_PER_PAIR } from '../config.js';
import { addDays, weekday } from '../util/dates.js';

export const DOW_SHORT = ['вс', 'пн', 'вт', 'ср', 'чт', 'пт', 'сб'];
export const DOW_FULL = ['Воскресенье', 'Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота'];

export const hoursOfPairs = pairs => pairs.length * HOURS_PER_PAIR;

// [4,5,6] → «4-6 пара»; [3] → «3 пара»; [1,3] → «1, 3 пара»; [1,2,4,5,6] → «1-2, 4-6 пара»
export function pairsText(pairs) {
  const p = [...new Set(pairs)].sort((a, b) => a - b);
  if (!p.length) return '';
  const parts = [];
  for (let i = 0; i < p.length;) {
    let j = i;
    while (j + 1 < p.length && p[j + 1] === p[j] + 1) j++;
    parts.push(j > i ? `${p[i]}-${p[j]}` : String(p[i]));
    i = j + 1;
  }
  return `${parts.join(', ')} пара`;
}

// «4-6 пара», «1-2, 4 пара», «3 пара» → [4,5,6] …; всё остальное («праздник», «08:00 – 12:55») → null
export function parsePairsText(text) {
  const m = /^\s*([\d\s,\-–]+?)\s*пар/i.exec(text || '');
  if (!m) return null;
  const out = [];
  for (const part of m[1].split(',')) {
    const r = /^\s*(\d+)\s*(?:[-–]\s*(\d+))?\s*$/.exec(part);
    if (!r) return null;
    const a = Number(r[1]), b = r[2] ? Number(r[2]) : a;
    if (b < a || b - a > 12) return null;
    for (let n = a; n <= b; n++) out.push(n);
  }
  return out.length ? out : null;
}

export function datesBetween(from, to) {
  const out = [];
  for (let d = from; d <= to; d = addDays(d, 1)) out.push(d);
  return out;
}

// Задать день: пары (часы считаются сами) либо только часы. Пустые пары и 0 часов — дня практики нет.
export function setDay(block, date, { pairs = null, hours = null } = {}) {
  block.days ||= {}; block.slots ||= {};
  const h = pairs && pairs.length ? hoursOfPairs(pairs) : Number(hours) || 0;
  if (!(h > 0)) { delete block.days[date]; delete block.slots[date]; return; }
  block.days[date] = h;
  if (pairs && pairs.length) block.slots[date] = [...pairs].sort((a, b) => a - b); else delete block.slots[date];
}

// Блок «период × часов в день» (простой режим) → блок по дням. Нерабочие дни календаря пропускаются.
export function expandBlock(b, calendar) {
  if (b.days) return { from: b.from, to: b.to, days: { ...b.days }, slots: { ...(b.slots || {}) } };
  const out = { from: b.from, to: b.to, days: {}, slots: {} };
  const h = Number(b.hoursPerDay) || 0;
  for (const date of datesBetween(b.from, b.to)) {
    const dow = weekday(date);
    if (h > 0 && ((dow >= 1 && dow <= 5) || (dow === 6 && b.saturday)) && !calendar.dayOff(date)) out.days[date] = h;
  }
  return out;
}

// Шаблон недели {1: [4,5,6], 2: [4,5,6], 5: []} применяется ко всем датам периода; праздники и выходные пропускаются.
// Дни недели, которых нет в шаблоне, не меняются; пустой список — практики в этот день недели нет.
export function applyPattern(block, pattern, calendar) {
  for (const date of datesBetween(block.from, block.to)) {
    const dow = weekday(date);
    if (!(dow in pattern)) continue;
    if (calendar.dayOff(date)) { setDay(block, date, {}); continue; }
    setDay(block, date, { pairs: pattern[dow] });
  }
}

// Шаблон недели из уже заданных дней: для каждого дня недели — самый частый набор пар (null — пар нет, только часы)
export function patternOf(block) {
  const stat = {};
  for (const date of Object.keys(block.days || {})) {
    if (!(block.days[date] > 0)) continue;
    const dow = weekday(date);
    const key = block.slots?.[date] ? block.slots[date].join(',') : '';
    ((stat[dow] ||= {})[key] = (stat[dow][key] || 0) + 1);
  }
  const out = {};
  for (const dow of Object.keys(stat)) {
    const best = Object.entries(stat[dow]).sort((a, b) => b[1] - a[1])[0][0];
    out[dow] = best ? best.split(',').map(Number) : null;
  }
  return out;
}

const dayValue = (block, date) => (block.slots?.[date] ? pairsText(block.slots[date]) : `${block.days[date]} ч`);
const shortDate = date => date.split('-').reverse().slice(0, 2).join('.');

// Текст «в какие дни»: «пн 2-3 пара, вт 4-6 пара … · практики нет: пт · иначе: 13.10 — 3-6 пара»
export function describe(blocks) {
  const byDow = {};
  const dates = blocks.flatMap(b => Object.keys(b.days || {}).filter(d => b.days[d] > 0).map(d => ({ d, b })));
  for (const { d, b } of dates) {
    const dow = weekday(d);
    ((byDow[dow] ||= {})[dayValue(b, d)] ||= []).push(d);
  }
  const main = [];
  const other = {};                 // текст → даты, где день отличается от обычного для своего дня недели
  for (const dow of [1, 2, 3, 4, 5, 6]) {
    const values = byDow[dow];
    if (!values) continue;
    const ranked = Object.entries(values).sort((a, b) => b[1].length - a[1].length);
    main.push(`${DOW_SHORT[dow]} ${ranked[0][0]}`);
    for (const [text, ds] of ranked.slice(1)) (other[text] ||= []).push(...ds);
  }
  const none = [1, 2, 3, 4, 5].filter(dow => !byDow[dow]).map(dow => DOW_SHORT[dow]);
  const parts = [main.join(', ')];
  if (main.length && none.length) parts.push(`практики нет: ${none.join(', ')}`);
  const exceptions = Object.entries(other).map(([text, ds]) => {
    const list = ds.sort().map(shortDate);
    return `${text}: ${list.slice(0, 6).join(', ')}${list.length > 6 ? ` и ещё ${list.length - 6}` : ''}`;
  });
  if (exceptions.length) parts.push(`иначе — ${exceptions.join('; ')}`);
  return main.length ? parts.join(' · ') : '';
}

export const totalHours = blocks => blocks.reduce((s, b) => s + Object.values(b.days || {}).reduce((a, h) => a + h, 0), 0);

// Строки таблицы документа практики: рабочие дни периода + праздники на днях недели, в которые практика идёт
export function documentDays(block, calendar) {
  const used = new Set(Object.keys(block.days || {}).filter(d => block.days[d] > 0).map(weekday));
  const out = [];
  for (const date of datesBetween(block.from, block.to)) {
    const dow = weekday(date);
    const hours = block.days?.[date] || 0;
    if (hours > 0) out.push({ date, dow, text: block.slots?.[date] ? pairsText(block.slots[date]) : '', hours });
    else {
      const off = used.has(dow) && calendar.dayOff(date);
      if (off && off.kind !== 'weekend') out.push({ date, dow, text: off.kind === 'holiday' ? 'праздник' : 'нерабочий день', hours: 0 });
    }
  }
  return out;
}
