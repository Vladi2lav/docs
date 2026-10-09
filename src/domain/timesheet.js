// Расчёт учёта времени: строки ведомости × дни месяцев → листы по месяцам.
// Чистая функция: на входе строки, график, календарь и переносы; на выходе данные для Word и предпросмотра.

import { CARRY_NOTE, FILL, HOURS_PER_PAIR, MONTH_NAMES, VACATION_CODE } from '../config.js';
import { daysInMonth, iso, monthsBetween, ruDate, weekday, ymIndex, ymOf } from '../util/dates.js';

// transfers: [{ rowId, from, to, hours }] — занятие за дату `from` проведено (отработано) в дату `to`
export function buildTimesheet({ rows, graph, calendar, from, to, carryOver, transfers = [], swapWeeks = false }) {
  const graphStart = ymOf(graph.weeks[0].start);
  const first = ymIndex(from) < ymIndex(graphStart) ? from : graphStart;   // накопительный итог — с начала графика
  // Считаем до конца графика: план по умолчанию = все часы года, тогда в конце семестра виден остаток «на 2 сем»
  const graphEnd = ymOf(graph.weeks[graph.weeks.length - 1].end);
  const lastMonth = ymIndex(to) > ymIndex(graphEnd) ? to : graphEnd;
  const months = monthsBetween(first, lastMonth).map(({ y, m }) => {
    const days = Array.from({ length: daysInMonth(y, m) }, (_, i) => {
      const date = iso(y, m, i + 1);
      return { n: i + 1, date, off: calendar.dayOff(date) };
    });
    return { y, m, days, cells: rows.map(row => days.map(day => cellFor(row, day, graph, swapWeeks))) };
  });

  const cellAt = (rowIndex, date) => {
    const mo = months.find(x => x.y === Number(date.slice(0, 4)) && x.m === Number(date.slice(5, 7)) - 1);
    return mo ? mo.cells[rowIndex][Number(date.slice(8)) - 1] : null;
  };

  const missed = findMissed(rows, months, graph, transfers, from, to, swapWeeks);
  const notes = [];
  for (const t of transfers) {
    const i = rows.findIndex(r => r.id === t.rowId);
    if (i < 0) continue;
    const src = cellAt(i, t.from);
    const dst = cellAt(i, t.to);
    if (src) {
      src.hours = Math.max(0, src.hours - t.hours);
      if (!src.hours && !src.fill) src.fill = FILL.moved;
    }
    if (dst) {
      const regular = dst.hours;
      dst.hours += t.hours;
      dst.fill = FILL.makeup;
      const label = `${t.hours} отраб за ${shortDate(t.from)}`;
      dst.label = dst.label ? `${dst.label}; ${label}` : (regular ? `${dst.hours} (${regular}+${label})` : label);
    }
    notes.push({ date: t.to, from: t.from, text: `${t.hours} ч — занятие за ${ruDate(t.from)} проведено ${ruDate(t.to)} (${rows[i].name})` });
  }

  const totals = rows.map((_, i) => sum(months.flatMap(mo => mo.cells[i].map(c => c.hours))));   // справка: все часы по расписанию
  const planned = rows.map(row => row.plan);                                                     // null — план не задан
  const cumulative = rows.map(() => 0);
  const sheets = [];

  months.forEach(mo => {
    const isLast = ymIndex(mo) === ymIndex(to);
    const sheetRows = rows.map((row, i) => {
      const total = sum(mo.cells[i].map(c => c.hours));
      cumulative[i] += total;
      const remaining = planned[i] === null ? null : planned[i] - cumulative[i];
      const carry = carryOver && isLast && remaining > 0;
      return {
        moduleIndex: row.moduleIndex,
        name: row.name,
        group: row.group,
        cells: mo.cells[i].map(c => ({ text: c.label || (c.hours > 0 ? String(c.hours) : ''), fill: c.fill || null, vertical: !!c.label })),
        total,
        plan: planned[i] ?? '',
        cumulative: cumulative[i],
        remainingText: remaining === null ? '' : (carry ? `${remaining} (${CARRY_NOTE})` : String(remaining)),
        remainingFill: carry ? FILL.mark : null,
      };
    });
    if (ymIndex(mo) < ymIndex(from) || ymIndex(mo) > ymIndex(to)) return;
    const prefix = `${mo.y}-${String(mo.m + 1).padStart(2, '0')}`;
    sheets.push({
      year: mo.y,
      month: mo.m,
      name: MONTH_NAMES[mo.m].toUpperCase(),
      days: mo.days.map(d => ({ n: d.n, fill: d.off ? d.off.fill : null, title: d.off ? d.off.name : '' })),
      rows: sheetRows.map((r, i) => ({ ...r, num: i + 1 })),
      total: sum(sheetRows.map(r => r.total)),
      notes: notes.filter(n => n.date.startsWith(prefix) || n.from.startsWith(prefix)).map(n => n.text),
    });
  });

  // «ост на 2 сем» — ещё и в последнем месяце, где у дисциплины были часы (дальше занятий нет)
  if (carryOver) {
    rows.forEach((_, i) => {
      const lastActive = sheets.filter(sh => sh.rows[i].total > 0).pop();
      const r = lastActive && lastActive.rows[i];
      const remaining = r && r.plan !== '' ? r.plan - r.cumulative : 0;
      if (remaining > 0) { r.remainingText = `${remaining} (${CARRY_NOTE})`; r.remainingFill = FILL.mark; }
    });
  }

  return { sheets, totals, missed };
}

// Часы по расписанию без учёта выходных/праздников (нужны, чтобы найти сорванные праздником занятия)
function scheduledHours(row, day, graph, swap) {
  if (graph.codeAt(row.group, day.date)) return 0;
  const parity = graph.parityAt(day.date, swap);
  if (!parity) return 0;
  const dow = weekday(day.date);
  return ((row.pairs.both[dow] || 0) + (row.pairs[parity][dow] || 0)) * HOURS_PER_PAIR;
}

function cellFor(row, day, graph, swap) {
  const idle = () => ({
    hours: 0,
    fill: day.off ? day.off.fill : (graph.codeAt(row.group, day.date) === VACATION_CODE ? FILL.vacation : null),
  });
  if (row.kind === 'practice') {
    const hours = practiceHours(row, day);
    return hours > 0 ? { hours } : idle();
  }
  if (day.off) return idle();
  const hours = scheduledHours(row, day, graph, swap);
  return hours > 0 ? { hours } : idle();
}

// Блок практики: либо точные часы по датам (из docx), либо период × часов в день по будням
function practiceHours(row, day) {
  for (const b of row.blocks) {
    if (b.days) {
      if (day.date in b.days) return b.days[day.date];
      continue;
    }
    if (day.date < b.from || day.date > b.to || day.off) continue;
    const dow = weekday(day.date);
    if ((dow >= 1 && dow <= 5) || (dow === 6 && b.saturday)) return Number(b.hoursPerDay) || 0;
  }
  return 0;
}

// Занятия теории, выпавшие на праздник/нерабочий день и ещё не перенесённые
function findMissed(rows, months, graph, transfers, from, to, swap) {
  const out = [];
  rows.forEach(row => {
    if (row.kind !== 'theory') return;
    for (const mo of months) {
      if (ymIndex(mo) < ymIndex(from) || ymIndex(mo) > ymIndex(to)) continue;
      for (const day of mo.days) {
        if (!day.off || day.off.kind === 'weekend') continue;
        const hours = scheduledHours(row, day, graph, swap);
        if (hours > 0 && !transfers.some(t => t.rowId === row.id && t.from === day.date)) {
          out.push({ rowId: row.id, name: row.name, date: day.date, hours, reason: day.off.name });
        }
      }
    }
  });
  return out;
}

const shortDate = date => ruDate(date).slice(0, 5);   // 31.08
const sum = list => list.reduce((a, b) => a + b, 0);
