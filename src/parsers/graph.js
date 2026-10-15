// График учебного процесса (xlsx) → недели, коды по группам, расшифровка модулей, реквизиты колледжа.
//
// Структура листа:
//   • строка «Дата/ группа»: названия месяцев (и диапазоны вида «28.09 -3.10» для недель на стыке месяцев);
//   • ниже: числа начала недели, числа конца недели, номера недель 1…N;
//   • далее строки групп: пустая ячейка недели = теория, иначе код (практика, К — каникулы, Э …);
//   • блок «Обозначения»: код → модуль/практика, для каждой группы-семейства («п2», «вм2» …).

import { LEFT_WEEK_IS_ODD, VACATION_CODE } from '../config.js';
import { iso } from '../util/dates.js';

const LETTERS = 'А-ЯЁӘІҢҒҮҰҚӨҺA-Z';
const MONTHS = {
  сентябр: 8, октябр: 9, ноябр: 10, декабр: 11, январ: 0, феврал: 1,
  март: 2, апрел: 3, май: 4, июн: 5, июл: 6, август: 7,
};
// Шифр группы в начале подписи: «П2К-д», «П1А(д)», «П4А (WEB)», «ИС2У заоч» → «П2К», «П1А», «П4А», «ИС2У»
const GROUP_TOKEN = new RegExp(`^([${LETTERS}]{1,6})\\s*(\\d{1,2})\\s*([${LETTERS}]{0,2})(?![${LETTERS}])`);
const GROUP_EXTRA = new RegExp(`^\\s*/\\s*([${LETTERS}]{1,2})(?![${LETTERS}])`);

export class AcademicGraph {
  constructor(data) { Object.assign(this, data); }

  weekAt(date) { return this.weeks.find(w => date >= w.start && date <= w.end) || null; }

  // swap — поменять местами левую и правую недели (если расписание начинается с «правой»)
  parityAt(date, swap = false) {
    const w = this.weekAt(date);
    if (!w) return null;
    const odd = w.num % 2 === 1;
    return odd === (LEFT_WEEK_IS_ODD !== swap) ? 'left' : 'right';
  }

  // Код недели группы: '' — теория, строка — практика/каникулы/…, null — группа или неделя неизвестны.
  // Объединённая группа расписания («П2У/Г») — код известен, только если у всех её групп он одинаковый.
  codeAt(group, date) {
    const w = this.weekAt(date);
    const gs = groupKeys(group).map(k => this.groups[k]);
    if (!w || !gs.length || gs.some(g => !g)) return null;
    const codes = [...new Set(gs.map(g => g.codes[w.num] || ''))];
    if (codes.length > 1) return null;
    return this.theoryCodes.includes(codes[0]) ? '' : codes[0];      // «праздничная неделя» — обычная учебная неделя
  }

  // Есть ли группа (все части объединённой) в графике
  hasGroup(group) {
    const keys = groupKeys(group);
    return keys.length > 0 && keys.every(k => this.groups[k]);
  }

  // Расшифровка кодов для группы — только при точном совпадении «семейства» (П2А → «п2»), без подбора похожих
  legendFor(group) { return this.legend[familyOf(group)] || null; }


  // Модуль группы по индексу («ПМ4»): { code, index, title } или null
  moduleOf(group, index) {
    const legend = this.legendFor(group);
    if (!legend) return null;
    const code = Object.keys(legend).find(c => legend[c].index === index);
    return code ? { code, ...legend[code] } : null;
  }

  // Непрерывные периоды недель группы с кодом code → [{ from, to }] (от начала первой до конца последней недели)
  runsOf(group, code) {
    const g = this.groups[normGroup(group)];
    if (!g) return [];
    const runs = [];
    let cur = null;
    for (const w of this.weeks) {
      if (g.codes[w.num] === code) {
        if (cur) cur.to = w.end; else { cur = { from: w.start, to: w.end }; runs.push(cur); }
      } else cur = null;
    }
    return runs;
  }

  // Семестры по графику: 1-й заканчивается месяцем зимних каникул (первый период «К» у групп педагога),
  // 2-й начинается сразу после них и идёт до следующих каникул либо до конца графика.
  semesters(groups) {
    const names = (groups && groups.length ? groups : Object.keys(this.groups)).map(normGroup).filter(n => this.groups[n]);
    const vacation = this.weeks.map(w => names.some(n => this.groups[n].codes[w.num] === VACATION_CODE));
    const runs = [];
    vacation.forEach((v, i) => {
      if (!v) return;
      if (i > 0 && vacation[i - 1]) runs[runs.length - 1].last = i;
      else if (i > 0) runs.push({ first: i, last: i });      // каникулы в самом начале графика не считаем
    });

    const ym = date => ({ y: Number(date.slice(0, 4)), m: Number(date.slice(5, 7)) - 1 });
    const last = this.weeks.length - 1;
    const sem1End = runs[0] ? this.weeks[runs[0].last].end : this.weeks[last].end;
    const sem1 = { from: ym(this.weeks[0].start), to: ym(sem1End) };
    if (!runs[0] || runs[0].last >= last) return [sem1];
    const sem2End = runs[1] ? this.weeks[runs[1].last].end : this.weeks[last].end;
    return [sem1, { from: ym(this.weeks[runs[0].last + 1].start), to: ym(sem2End) }];
  }
}

// «П2У/Г» → ['П2У', 'П2Г'];  «П2К-д» → ['П2К'];  не похоже на группу → []
export function groupKeys(name) {
  const s = String(name).toUpperCase().trim();
  const m = GROUP_TOKEN.exec(s);
  if (!m) return [];
  const base = m[1] + m[2];
  const keys = [base + m[3]];
  let rest = s.slice(m[0].length);
  for (let x = GROUP_EXTRA.exec(rest); x; x = GROUP_EXTRA.exec(rest)) {
    keys.push(base + x[1]);
    rest = rest.slice(x[0].length);
  }
  return keys;
}
export const normGroup = name => groupKeys(name)[0] || String(name).replace(/\s+/g, '').toUpperCase();
export const familyOf = group => {
  const m = /^([^\d]+)(\d+)/.exec(normGroup(group));
  return m ? (m[1] + m[2]).toLowerCase() : '';
};

export function guessStartYear(fileName) {
  const m = /(20\d\d)\s*[-–/]\s*(20\d\d)/.exec(fileName || '');
  return m ? Number(m[1]) : null;
}

export function parseGraph(sheet, { startYear }) {
  if (!startYear) throw new Error('Укажите учебный год (например 2026) — в графике он не указан');

  const anchor = findAnchor(sheet);
  if (!anchor) throw new Error('Не найдена шапка графика (ячейка «Дата/ группа» и номера недель)');

  const weeks = readWeeks(sheet, anchor, startYear);
  const groups = readGroups(sheet, anchor);
  const legendRow = findRow(sheet, anchor.weekRow, /Обозначения/i);

  return new AcademicGraph({
    startYear,
    weeks,
    groups,
    legend: readLegend(sheet, legendRow >= 0 ? legendRow : anchor.weekRow),
    theoryCodes: readTheoryCodes(sheet, legendRow >= 0 ? legendRow : anchor.weekRow),
    college: readCollege(sheet, anchor.monthRow),
    deputy: readDeputy(sheet),
    specialties: readSpecialties(sheet),
  });
}

function findRow(sheet, from, re) {
  for (let r = from; r < sheet.rowCount; r++) {
    for (let c = 0; c < sheet.colCount; c++) if (re.test(sheet.get(r, c))) return r;
  }
  return -1;
}

function findAnchor(sheet) {
  const monthRow = findRow(sheet, 0, /Дата\s*\/\s*групп/i);
  if (monthRow < 0) return null;
  for (let r = monthRow + 1; r < Math.min(monthRow + 10, sheet.rowCount); r++) {
    const nums = [];
    for (let c = 0; c < sheet.colCount; c++) if (/^\d+$/.test(sheet.get(r, c))) nums.push(Number(sheet.get(r, c)));
    if (nums.length > 10 && nums[0] === 1 && nums[1] === 2) {
      const numeric = [];
      for (let k = monthRow + 1; k < r; k++) {
        const filled = Array.from({ length: sheet.colCount }, (_, c) => sheet.get(k, c)).filter(t => /^\d+$/.test(t)).length;
        if (filled > 5) numeric.push(k);
      }
      return { monthRow, weekRow: r, startRow: numeric[0], endRow: numeric[numeric.length - 1] };
    }
  }
  return null;
}

function readWeeks(sheet, { monthRow, weekRow, startRow, endRow }, startYear) {
  // Год увеличивается, когда номер месяца в шапке «переходит» через декабрь → январь (начало года не привязано к сентябрю)
  let year = startYear;
  let last = -1;
  const yearFor = m => { if (last !== -1 && m < last) year++; last = m; return year; };
  const weeks = [];
  let month = null;
  let monthYear = startYear;

  for (let c = 0; c < sheet.colCount; c++) {
    const head = sheet.get(monthRow, c);
    let range = null;
    if (head) {
      const r = /(\d{1,2})\.(\d{1,2})\s*[-–]\s*(\d{1,2})\.(\d{1,2})/.exec(head);
      if (r) range = { s: [Number(r[1]), Number(r[2]) - 1], e: [Number(r[3]), Number(r[4]) - 1] };
      else {
        const key = Object.keys(MONTHS).find(k => head.toLowerCase().startsWith(k));
        if (key) { month = MONTHS[key]; monthYear = yearFor(month); }
      }
    }
    const num = /^\d+$/.test(sheet.get(weekRow, c)) ? Number(sheet.get(weekRow, c)) : 0;
    if (!num) continue;

    let start; let end;
    if (range) {
      start = iso(yearFor(range.s[1]), range.s[1], range.s[0]);
      end = iso(yearFor(range.e[1]), range.e[1], range.e[0]);
      monthYear = year;
    } else if (month !== null && sheet.get(startRow, c) && sheet.get(endRow, c)) {
      const startDay = Number(sheet.get(startRow, c));
      const endDay = Number(sheet.get(endRow, c));
      start = iso(monthYear, month, startDay);
      // неделя на стыке месяцев без подписи «27.10-1.11»: конец меньше начала — значит, это уже следующий месяц
      end = endDay < startDay
        ? iso(month === 11 ? monthYear + 1 : monthYear, (month + 1) % 12, endDay)
        : iso(monthYear, month, endDay);
    } else continue;

    weeks.push({ num, col: c, start, end });
  }
  if (!weeks.length) throw new Error('В графике не удалось определить даты недель');
  return weeks;
}

function readGroups(sheet, { weekRow }) {
  const nameCol = findNameCol(sheet, weekRow);
  const weekCols = [];
  for (let c = 0; c < sheet.colCount; c++) if (/^\d+$/.test(sheet.get(weekRow, c))) weekCols.push({ c, num: Number(sheet.get(weekRow, c)) });

  const groups = {};
  for (let r = weekRow + 1; r < sheet.rowCount; r++) {
    if (/Обозначения/i.test(sheet.get(r, 1)) || /Обозначения/i.test(sheet.get(r, 2))) break;
    const name = sheet.get(r, nameCol).trim();
    if (!groupKeys(name).length) continue;
    const codes = {};
    // число в ячейке недели — пометка «кол-во недель теории», а не код
    for (const { c, num } of weekCols) codes[num] = /^\d+$/.test(sheet.get(r, c)) ? '' : sheet.get(r, c);
    groups[normGroup(name)] = { name, codes };
  }
  return groups;
}

// Колонка с названиями групп — та, где больше всего значений вида «П2А»
function findNameCol(sheet, weekRow) {
  let best = 2; let bestCount = -1;
  for (let c = 0; c < 6; c++) {
    let n = 0;
    for (let r = weekRow + 1; r < sheet.rowCount; r++) if (groupKeys(sheet.get(r, c)).length) n++;
    if (n > bestCount) { best = c; bestCount = n; }
  }
  return best;
}

// Легенда: [семейство группы] [код] [«ПМ4.Название …»] — три ячейки подряд
function readLegend(sheet, fromRow) {
  const legend = {};
  const RE = /^([A-ZА-ЯЁ]{2,3}\d{1,2})\s*\.?\s*(.+)$/;
  for (let r = fromRow; r < sheet.rowCount; r++) {
    for (let c = 2; c < sheet.colCount; c++) {
      const m = RE.exec(sheet.get(r, c));
      const code = sheet.get(r, c - 1);
      const family = sheet.get(r, c - 2).toLowerCase();
      if (!m || !code || code.length > 4 || !/^[a-zа-яё]{1,5}\d$/i.test(family)) continue;
      (legend[family] ||= {})[code] = { index: m[1], title: m[2].trim() };
    }
  }
  return legend;
}

// Коды, которые «Обозначения» графика расшифровывают как «Праздничная неделя»: занятия по расписанию идут,
// а сами праздничные дни выпадают по календарю.
function readTheoryCodes(sheet, fromRow) {
  const codes = new Set();
  for (let r = fromRow; r < sheet.rowCount; r++) {
    for (let c = 0; c < sheet.colCount - 1; c++) {
      const code = sheet.get(r, c);
      if (code && code.length <= 3 && /праздничн/i.test(sheet.get(r, c + 1))) codes.add(code);
    }
  }
  return [...codes];
}

function readCollege(sheet, untilRow) {
  for (let r = 0; r < untilRow; r++) {
    for (let c = 0; c < sheet.colCount; c++) {
      const m = /колледж|училищ/i.test(sheet.get(r, c)) && /["«“]([^"»”]+)["»”]/.exec(sheet.get(r, c));
      if (m) return m[1].trim();
    }
  }
  return '';
}

function readDeputy(sheet) {
  const r = findRow(sheet, 0, /Заместитель директора по учебной работе/i);
  if (r < 0) return '';
  for (let c = 0; c < sheet.colCount; c++) {
    const m = /Заместитель директора по учебной работе[\s_]*(.*)$/i.exec(sheet.get(r, c));
    if (m) return m[1].replace(/^_+\s*/, '').trim();
  }
  return '';
}

function readSpecialties(sheet) {
  const found = [];
  const RE = /^[«"“\s]*((?:\d{8})|(?:[34][SW]\d{8}))\s*[-–—.]*\s*(.+?)[»"”\s]*$/;
  for (let r = 0; r < sheet.rowCount; r++) {
    for (let c = 0; c < sheet.colCount; c++) {
      const m = RE.exec(sheet.get(r, c));
      if (m) found.push(`${m[1]} ${m[2].trim()}`);
    }
  }
  return [...new Set(found)];
}
