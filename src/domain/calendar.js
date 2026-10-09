// Календарь рабочих дней. Источники, по возрастанию приоритета:
//   1. воскресенье (и суббота, если неделя пятидневная);
//   2. праздники РК из data/holidays-kz.json — ежегодные + плавающие даты; праздник на выходной
//      переносится на ближайший рабочий день (если у праздника не стоит transfer:false);
//   3. ручные правки пользователя: «нерабочий» / «рабочий» день и переносы «с даты на дату».

import { FILL } from '../config.js';
import { iso, addDays, weekday } from '../util/dates.js';

// Ручные правки = клики по календарю + результат переносов
export function effectiveOverrides({ overrides = {}, transfers = [] }) {
  const out = { ...overrides };
  for (const t of transfers) {
    out[t.from] = { type: 'work', name: t.name };
    out[t.to] = { type: 'off', name: t.name };
  }
  return out;
}

export function createCalendar({ holidays, state, workSaturday, years }) {
  const manual = effectiveOverrides(state || {});
  const isWeekend = date => {
    const d = weekday(date);
    return d === 0 || (d === 6 && !workSaturday);
  };

  const off = new Map();   // date → { name, fill }
  for (let y = years[0]; y <= years[1]; y++) {
    for (const h of holidays.fixed) {
      let date = iso(y, h.month - 1, h.day);
      if (h.transfer !== false) while (isWeekend(date) || off.has(date)) date = addDays(date, 1);
      off.set(date, { name: h.name, fill: h.fill || FILL.holiday });
    }
  }
  for (const h of holidays.dates) off.set(h.date, { name: h.name, fill: h.fill || FILL.holiday });

  // → { kind: 'weekend'|'holiday'|'custom', name, fill, manual } для нерабочего дня или null
  function dayOff(date) {
    const m = manual[date];
    if (m) return m.type === 'work' ? null : { kind: 'custom', name: m.name || 'Нерабочий день', fill: FILL.holiday, manual: true };
    if (isWeekend(date)) return { kind: 'weekend', name: 'Выходной', fill: FILL.weekend };
    const h = off.get(date);
    return h ? { kind: 'holiday', name: h.name, fill: h.fill } : null;
  }

  // Для окна календаря: как день выглядел бы без ручных правок и есть ли правка
  const info = date => ({ off: dayOff(date), manual: manual[date] || null });

  return { dayOff, info };
}
