// Строки ведомости педагога.
//   теория   — из расписания: группа, дисциплина (текст как в расписании), пары по дням и неделям;
//   практика — только из docx практики, где руководитель назван прямо, либо добавляется вручную кнопкой.
// Ничего не угадывается: индекс модуля, план и часы практики в день заполняются, только если они прямо указаны.

import { PRACTICE_PREFIX } from '../config.js';
import { addDays, ruDate, toDate, weekday } from '../util/dates.js';

export const teacherNames = entries => [...new Set(entries.map(e => e.teacher))].sort((a, b) => a.localeCompare(b, 'ru'));

// «Н.Жаксыбаева» → «Жаксыбаева Н.Н.»: совпадает фамилия и первая буква имени; при неоднозначности — никто
export function matchSupervisor(names, supervisor) {
  const m = /^([^.]+)\.\s*(.+)$/.exec(supervisor || '');
  if (!m) return null;
  const initial = m[1].trim().toLowerCase();
  const surname = m[2].trim().toLowerCase();
  const hits = names.filter(n => {
    const [sn, ini = ''] = n.toLowerCase().split(/\s+/);
    return sn === surname && ini.startsWith(initial);
  });
  return hits.length === 1 ? hits[0] : null;
}

export function buildRows({ teacher, entries, graph, practiceDocs }) {
  const rows = [];
  const warnings = [];
  const byKey = new Map();

  for (const e of entries.filter(x => x.teacher === teacher)) {
    const key = `${e.group}|${e.subject}`;
    let row = byKey.get(key);
    if (!row) {
      row = { id: `t${rows.length}`, kind: 'theory', moduleIndex: '', name: e.subject, group: e.group, plan: null, subject: e.subject, pairs: { both: {}, left: {}, right: {} } };
      byKey.set(key, row);
      rows.push(row);
    }
    row.pairs[e.week][e.day] = (row.pairs[e.week][e.day] || 0) + 1;
  }

  const range = { from: graph.weeks[0].start, to: graph.weeks[graph.weeks.length - 1].end };
  for (const doc of practiceDocs) {
    const mod = graph.moduleOf(doc.group, doc.moduleIndex);
    let blocks;
    if (Object.keys(doc.days).some(d => d >= range.from && d <= range.to)) {
      blocks = [{ from: doc.from, to: doc.to, days: doc.days, fromDoc: true }];
    } else {
      const placed = mod && shiftToGraph(doc, graph.runsOf(doc.group, mod.code));
      if (!placed) {
        warnings.push(`Практика ${doc.group} ${doc.moduleIndex}: документ другого года (${ruDate(doc.from)} – ${ruDate(doc.to)}), а в графике нет периода практики этого модуля — документ не применён.`);
        continue;
      }
      blocks = placed.blocks;
      warnings.push(`Практика ${doc.group} ${doc.moduleIndex}: документ составлен на другой год (${ruDate(doc.from)} – ${ruDate(doc.to)}). Часы по дням перенесены на период практики из графика (${ruDate(placed.target.from)} – ${ruDate(placed.target.to)}), праздники берутся из календаря.`);
    }
    rows.push({
      id: `p${rows.length}`, kind: 'practice', moduleIndex: doc.moduleIndex,
      name: PRACTICE_PREFIX + (mod ? mod.title : doc.moduleTitle), group: doc.group, plan: null, blocks,
    });
  }
  snapshot(rows);
  return { rows, warnings };
}

// Практика, добавленная вручную: всё, кроме префикса «УП - », заполняет пользователь
export function emptyPracticeRow({ id, group, start }) {
  const row = {
    id, kind: 'practice', moduleIndex: '', name: PRACTICE_PREFIX, group, plan: null,
    blocks: [{ from: start, to: start, hoursPerDay: '', saturday: false }],
  };
  snapshot([row]);
  return row;
}

// Периоды практики из недель графика, где у группы стоит код модуля с этим индексом (точное совпадение), иначе []
export function graphPeriods(graph, group, moduleIndex) {
  const mod = moduleIndex && graph.moduleOf(group, moduleIndex);
  return mod ? graph.runsOf(group, mod.code).map(r => ({ ...r, hoursPerDay: '', saturday: false })) : [];
}

// Исходные значения из документов — к ним возвращает «Заполнить по документам»
function snapshot(rows) {
  for (const r of rows) r.defaults = { moduleIndex: r.moduleIndex, name: r.name, group: r.group, plan: r.plan, blocks: r.blocks ? r.blocks.map(b => ({ ...b })) : [] };
}

// Документ практики другого года: период документа совмещается с ближайшим по календарю периодом из графика
// (сдвиг кратен неделе — дни недели сохраняются), подряд идущие дни с равными часами сворачиваются в блоки
// «период × часов в день»; нулевые дни документа (праздники) отбрасываются — праздники берутся из календаря.
function shiftToGraph(doc, runs) {
  if (!runs.length) return null;
  const yearDelta = run => Math.abs((toDate(`${run.from.slice(0, 4)}${doc.from.slice(4)}`) - toDate(run.from)) / 864e5);
  const target = runs.reduce((best, r) => (yearDelta(r) < yearDelta(best) ? r : best));
  const shift = Math.round(((toDate(target.from) - toDate(doc.from)) / 864e5) / 7) * 7;

  const segments = [];
  for (const date of Object.keys(doc.days).sort().filter(d => doc.days[d] > 0)) {
    const last = segments[segments.length - 1];
    const gap = last ? (toDate(date) - toDate(last.to)) / 864e5 : 0;
    if (last && last.hoursPerDay === doc.days[date] && gap <= 7) {
      last.to = date;
      last.saturday ||= weekday(date) === 6;
    } else {
      segments.push({ from: date, to: date, hoursPerDay: doc.days[date], saturday: weekday(date) === 6, fromDoc: true });
    }
  }
  return { target, blocks: segments.map(s => ({ ...s, from: addDays(s.from, shift), to: addDays(s.to, shift) })) };
}
