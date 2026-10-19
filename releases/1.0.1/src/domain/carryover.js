// Перенос данных 1 семестра во 2-й: «выполнено с начала года» (base) и план — из ведомости 1 семестра.
// Строки сопоставляются только при однозначном совпадении: та же группа, тот же вид (теория/практика)
// и либо тот же индекс модуля, либо то же название. Не нашлось или неоднозначно — поля остаются пустыми.

import { PRACTICE_PREFIX } from '../config.js';
import { normGroup } from '../parsers/graph.js';

const norm = s => s.toLowerCase().replace(/[^a-zа-яёәіңғүұқөһ0-9]+/gi, ' ').trim();
const isPractice = name => name.trim().toLowerCase().startsWith(PRACTICE_PREFIX.trim().toLowerCase());

export function matchPrev(row, prev) {
  const same = prev.rows.filter(p => normGroup(p.group) === normGroup(row.group) && isPractice(p.name) === (row.kind === 'practice'));
  const hits = same.filter(p => (row.moduleIndex && p.moduleIndex === row.moduleIndex) || norm(p.name) === norm(row.name));
  return hits.length === 1 ? hits[0] : null;
}

// Заполняет только пустые base / plan; введённое вручную не трогает. Возвращает число найденных строк.
export function applyPrevSemester(rows, prev) {
  let found = 0;
  for (const row of rows) {
    const p = matchPrev(row, prev);
    if (!p) continue;
    found++;
    if (row.base == null) row.base = p.cumulative;
    if (row.plan == null && p.plan != null) row.plan = p.plan;
  }
  return found;
}
