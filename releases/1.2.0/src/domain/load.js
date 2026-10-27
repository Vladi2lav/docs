// Сопоставление строк ведомости со строками нагрузки.
// Подставляется только то, что выбирается однозначно; иначе пользователь выбирает строку нагрузки сам.

import { groupKeys } from '../parsers/graph.js';
import { normIndex } from '../parsers/load.js';

// Та же группа (учитываются объединённые «П3У/Г» и пометки «(д)», «заоч»)
export function sameGroup(a, b) {
  const x = groupKeys(a), y = groupKeys(b);
  return x.length > 0 && x.length === y.length && x.every(k => y.includes(k));
}

// Названия модулей в документах бывают обрезаны по-разному — совпадением считается начало одного в другом
const normTitle = t => String(t || '').toLowerCase().replace(/[^a-zа-яёәіңғүұқөһ0-9]+/gi, ' ').trim();
export function sameTitle(a, b) {
  const x = normTitle(a), y = normTitle(b);
  return x.length >= 15 && y.length >= 15 && (x.startsWith(y) || y.startsWith(x));
}

// Строки нагрузки группы, где в выбранном семестре (0 — первый, 1 — второй) есть часы
export function loadCandidates(load, group, semIdx) {
  return load.rows.map((row, i) => ({ row, i })).filter(({ row }) => sameGroup(row.group, group) && row.sems[semIdx]);
}

// Однозначное соответствие или null. rows — все строки ведомости педагога (чтобы теория по группе была одна).
export function matchLoad(row, rows, load, semIdx) {
  const cands = loadCandidates(load, row.group, semIdx);
  if (row.kind === 'practice') {
    const idx = normIndex(row.moduleIndex);
    let hits = idx ? cands.filter(c => c.row.moduleIndex === idx) : [];
    // номера модулей в графике и нагрузке бывают разными — тогда сопоставляется по названию модуля
    if (!hits.length) hits = cands.filter(c => c.row.moduleIndex && sameTitle(c.row.title, row.name.replace(/^УП\s*-\s*/i, '')));
    return hits.length === 1 ? hits[0] : null;
  }
  const theory = cands.filter(c => !c.row.moduleIndex);
  const mine = rows.filter(r => r.kind === 'theory' && sameGroup(r.group, row.group));
  return theory.length === 1 && mine.length === 1 ? theory[0] : null;
}

// Подпись строки нагрузки для списка выбора: «РО 10.1. Определять … — 48 ч»
export function loadLabel({ row }, semIdx) {
  const s = row.sems[semIdx];
  const weeks = s.weeks && s.perWeek ? ` (${s.weeks} нед × ${s.perWeek})` : '';
  return `${row.moduleIndex ? `${row.moduleIndex} · ` : ''}${row.title.slice(0, 60)}${row.title.length > 60 ? '…' : ''} — ${s.total} ч${weeks}`;
}

// Подставляет план из нагрузки в пустые поля плана. Возвращает число заполненных строк.
export function applyLoad(rows, load, semIdx) {
  let filled = 0;
  for (const row of rows) {
    if (row.loadIndex == null) {
      const m = matchLoad(row, rows, load, semIdx);
      if (m) row.loadIndex = m.i;
    }
    const src = row.loadIndex != null && load.rows[row.loadIndex];
    const sem = src && src.sems[semIdx];
    if (sem && row.plan == null) { row.plan = sem.total; row.planFromLoad = true; filled++; }
  }
  return filled;
}
