// Нагрузка преподавателя (xlsx, лист вида «Группа | Модуль | Название модуля/РО | 1 СЕМЕСТР | 2 СЕМЕСТР …»)
// → строки нагрузки: группа, индекс модуля, название и часы по семестрам.
//
// Столбцы ищутся по заголовкам двух верхних строк, а не по буквам: в одних годах таблица короче, в других
// справа дописаны «Преподаватель», «МБ / РБ / КОММЕРЧ» и примечания — всё это необязательно.

import { groupKeys } from './graph.js';

const num = text => {
  const n = Number(String(text || '').replace(',', '.'));
  return Number.isFinite(n) && String(text).trim() !== '' ? n : null;
};

// «ПМ 10» → «ПМ10»; «29» → «29»
export const normIndex = s => String(s || '').replace(/\s+/g, '').toUpperCase();

export function parseLoad(sheet) {
  const head = findHeader(sheet);
  if (!head) throw new Error('Не найдена шапка нагрузки (колонки «Группа», «Модуль», «1 СЕМЕСТР»). Это точно файл нагрузки?');
  const { top, cols } = head;

  // границы блоков семестров: от колонки «N СЕМЕСТР» до следующего такого же
  const semCols = [];
  for (let c = 0; c < sheet.colCount; c++) {
    const m = /(\d)\s*семестр/i.exec(sheet.get(top, c));
    if (m) semCols.push({ no: Number(m[1]), from: c });
  }
  const sub = top + 1;
  const blocks = semCols.map((s, i) => {
    const to = i + 1 < semCols.length ? semCols[i + 1].from : sheet.colCount;
    const find = re => { for (let c = s.from; c < to; c++) if (re.test(sheet.get(sub, c))) return c; return -1; };
    const totals = [];
    for (let c = s.from; c < to; c++) if (/^всего$/i.test(sheet.get(sub, c))) totals.push(c);
    return {
      no: s.no, weeks: find(/^нед/i), perWeek: find(/^часы/i), total: totals[totals.length - 1] ?? -1,
      lab: find(/^лаб/i), kp: find(/^кп$/i), kpAccept: find(/прием/i), exam: find(/конс|экзамен/i),
    };
  });

  const colOf = re => { for (let c = 0; c < sheet.colCount; c++) if (re.test(sheet.get(sub, c)) || re.test(sheet.get(top, c))) return c; return -1; };
  const extra = { teacher: colOf(/^преподаватель/i), mb: colOf(/^мб$/i), rb: colOf(/^рб$/i), comm: colOf(/коммерч/i) };
  const cell = (r, c) => (c >= 0 ? sheet.get(r, c) : '');

  const rows = [];
  for (let r = sub + 1; r < sheet.rowCount; r++) {
    const group = sheet.get(r, cols.group);
    if (!group || !groupKeys(group).length) continue;       // строки итогов и пустые
    const sems = [0, 1].map(i => {
      const b = blocks[i];
      if (!b) return null;
      const total = num(cell(r, b.total));
      return total == null ? null : {
        weeks: num(cell(r, b.weeks)), perWeek: num(cell(r, b.perWeek)), total,
        lab: num(cell(r, b.lab)), kp: num(cell(r, b.kp)), kpAccept: num(cell(r, b.kpAccept)), exam: num(cell(r, b.exam)),
      };
    });
    if (!sems[0] && !sems[1]) continue;
    const title = sheet.get(r, cols.title);
    const budget = cell(r, extra.mb) || cell(r, extra.rb);
    rows.push({
      group,
      moduleIndex: moduleIndexOf(sheet.get(r, cols.module), title),
      title,
      sems,
      teacher: cell(r, extra.teacher),
      funding: cell(r, extra.comm) ? 'comm' : (budget ? 'grant' : ''),
    });
  }
  if (!rows.length) throw new Error('В нагрузке не найдено ни одной строки с группой и часами');
  return { rows };
}

// Индекс модуля: колонка «Модуль» («ПМ 10», «29») либо «ПМ8.» в начале названия; «РО …» — не модуль
function moduleIndexOf(moduleCell, title) {
  const idx = normIndex(moduleCell);
  if (idx) return idx;
  const m = /^(ПМ)\s*(\d{1,2})\s*[.\s]/i.exec(title);
  return m ? normIndex(m[1] + m[2]) : '';
}

function findHeader(sheet) {
  for (let r = 0; r < Math.min(sheet.rowCount, 15); r++) {
    let group = -1, module = -1, title = -1, sem = -1;
    for (let c = 0; c < sheet.colCount; c++) {
      const t = sheet.get(r, c);
      if (/^группа$/i.test(t)) group = c;
      else if (/^модуль$/i.test(t)) module = c;
      else if (/^название/i.test(t)) title = c;
      else if (/\d\s*семестр/i.test(t) && sem < 0) sem = c;
    }
    if (group >= 0 && module >= 0 && sem >= 0) return { top: r, cols: { group, module, title: title >= 0 ? title : module + 1 } };
  }
  return null;
}
