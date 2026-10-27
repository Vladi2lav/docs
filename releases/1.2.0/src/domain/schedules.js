// Несколько расписаний (по одному на курс) → одно общее; сетка недели «пара × день» для вкладки «Расписание и практики».
// Ничего не угадывается: если одна группа встретилась в двух файлах — берётся тот, что загружен позже, с предупреждением.

import { groupKeys } from '../parsers/graph.js';

const groupId = name => groupKeys(name).join('/') || String(name).toUpperCase();

// list: [{ name, parsed }] в порядке загрузки → { entries, groups, semester, hasSaturday, warnings }
export function mergeSchedules(list, { graphYear = 0 } = {}) {
  const warnings = [];
  const owner = new Map();               // группа → индекс расписания, которому она принадлежит
  list.forEach((s, i) => {
    for (const g of s.parsed.groups) {
      const id = groupId(g);
      if (owner.has(id)) warnings.push(`Группа ${g} есть в двух расписаниях («${list[owner.get(id)].name}» и «${s.name}») — взята из «${s.name}».`);
      owner.set(id, i);
    }
    if (graphYear && s.parsed.startYear && s.parsed.startYear !== graphYear) {
      warnings.push(`Расписание «${s.name}» составлено на ${s.parsed.startYear}–${s.parsed.startYear + 1} учебный год, а график — на ${graphYear}–${graphYear + 1}. Проверьте, тот ли файл.`);
    }
  });

  const entries = list.flatMap((s, i) => s.parsed.entries.filter(e => owner.get(groupId(e.group)) === i));
  const groups = [...new Set(entries.map(e => e.group))];
  const sems = [...new Set(list.map(s => s.parsed.semester).filter(Boolean))];
  if (sems.length > 1) warnings.push(`Расписания относятся к разным семестрам (${sems.join(' и ')}). Загрузите расписания одного семестра.`);

  return {
    entries,
    groups: list.flatMap(s => s.parsed.groups).filter((g, i, a) => a.indexOf(g) === i && groups.includes(g)),
    semester: sems[0] || 0,
    hasSaturday: entries.some(e => e.day === 6),
    warnings,
  };
}

export const DAY_NAMES = ['Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];

// Время начала для сортировки: '08:00-09:30' → 480
const startMinutes = time => {
  const m = /^(\d\d):(\d\d)/.exec(time);
  return m ? Number(m[1]) * 60 + Number(m[2]) : Infinity;
};

// Звонки: все различные интервалы времени из расписаний по возрастанию. Номер звонка (1, 2, …) — общая нумерация
// пар для практики: номера пар в самих расписаниях у курсов разные (у 2 курса 1-я пара в 13:25, у 3 курса — в 08:00).
// Если загружены расписания не всех курсов, список неполный — его можно поправить вручную (см. parseBells).
export const bellTable = entries => [...new Set(entries.map(e => e.time).filter(Boolean))].sort((a, b) => startMinutes(a) - startMinutes(b));

// Текст «по строке на звонок» → список интервалов; '8:00-9:30', '08.00 – 09.30' приводятся к '08:00-09:30'
export function parseBells(text) {
  return String(text || '').split(/\r?\n/).map(line => {
    const m = /(\d{1,2})[:.](\d{2})\s*[-–—]\s*(\d{1,2})[:.](\d{2})/.exec(line);
    return m ? `${m[1].padStart(2, '0')}:${m[2]}-${m[3].padStart(2, '0')}:${m[4]}` : '';
  }).filter(Boolean);
}

// Сетка недели: cells[side][row][day] = [{ group, subject, teacher, mine }]; строки — звонки (+ строки «пара N»
// для занятий без времени). groups — какие группы показывать (null — все); onlyMine — только пары педагога.
export function buildGrid({ entries, bells, teacher = '', groups = null, onlyMine = false }) {
  const shown = entries.filter(e => (!groups || groups.includes(e.group)) && (!onlyMine || e.teacher === teacher));
  const rows = bells.map((time, i) => ({ key: time, number: i + 1, time, pairs: [] }));
  for (const e of shown) {
    const key = e.time || `пара ${e.pair}`;
    let row = rows.find(r => r.key === key);
    if (!row) { row = { key, number: 0, time: e.time || '', label: e.time ? '' : `пара ${e.pair}, время не указано`, pairs: [] }; rows.push(row); }
    if (!row.pairs.includes(e.pair)) row.pairs.push(e.pair);
  }
  const days = [1, 2, 3, 4, 5, 6];
  const blank = () => rows.map(() => days.map(() => []));
  const cells = { left: blank(), right: blank() };
  for (const e of shown) {
    const r = rows.findIndex(x => x.key === (e.time || `пара ${e.pair}`));
    const col = days.indexOf(e.day);
    if (r < 0 || col < 0) continue;
    const item = { group: e.group, subject: e.subject, teacher: e.teacher, mine: !!teacher && e.teacher === teacher };
    for (const side of e.week === 'both' ? ['left', 'right'] : [e.week]) cells[side][r][col].push(item);
  }
  for (const side of ['left', 'right']) for (const r of cells[side]) for (const c of r) c.sort((a, b) => Number(b.mine) - Number(a.mine) || a.group.localeCompare(b.group, 'ru'));
  rows.forEach(r => { r.pairs.sort((a, b) => a - b); r.used = days.some((_, d) => cells.left[rows.indexOf(r)][d].length || cells.right[rows.indexOf(r)][d].length); });
  return { rows, days, cells };
}

// Пары педагога в неделю по группам и дисциплинам: [{ group, subject, left, right }] (both считается в обе недели)
export function weeklySummary(entries, teacher) {
  const map = new Map();
  for (const e of entries.filter(x => x.teacher === teacher)) {
    const key = `${e.group}|${e.subject}`;
    if (!map.has(key)) map.set(key, { group: e.group, subject: e.subject, left: 0, right: 0 });
    const row = map.get(key);
    if (e.week !== 'right') row.left++;
    if (e.week !== 'left') row.right++;
  }
  return [...map.values()];
}
