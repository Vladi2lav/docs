// Расписание занятий (xlsx, сетка «пары × группы») → плоский список занятий.
//
// Структура сетки:
//   • строка заголовка: для каждой группы ячейка «№ пары» и правее — название группы;
//     группа занимает две колонки: левая и правая неделя (числитель / знаменатель);
//   • блок пары: строка с номером пары, внутри блока — предмет, преподаватель(и), аудитория;
//   • ячейка, объединённая на обе колонки, означает «каждую неделю».

const L = 'А-ЯЁӘІҢҒҮҰҚӨҺа-яёәіңғүұқөһA-Za-z';
// Ячейка педагога: «Иванов И.И.», а также «Иванов И.И. 303 ауд.» / «Иванов И.И. Спортивный зал» (фамилия и аудитория в одной ячейке)
const TEACHER = new RegExp(`^([${L}][${L}\\-]+\\s+[${L}]\\.\\s*(?:[${L}]\\.)?)(?:\\s+(.+))?$`);
// Два педагога на одну пару (подгруппы: языки) — «Даулетбеккызы Н./Кикбаева Ж.Ж.»: пара засчитывается обоим.
// Инициалы без последней точки («Нарбаева Г.И») дописываются. Если хоть одна часть не похожа на педагога — ячейка не педагог.
const teachersOf = text => {
  const parts = String(text).split('/').map(s => s.trim().replace(/^([^\s]+\s+[^\s.]\.[^\s.])(?=\s|$)/, '$1.'));
  const names = parts.map(teacherOf);
  return names.every(Boolean) ? [...new Set(names)] : [];
};
const teacherOf = text => {
  const m = TEACHER.exec(text);
  return m && (!m[2] || ROOM.test(m[2]) || /^\d+\b/.test(m[2])) ? m[1].trim() : null;
};
const ROOM = /ауд|зал/i;
const DAYS = [
  ['понедельник', 1], ['вторник', 2], ['среда', 3],
  ['четверг', 4], ['пятница', 5], ['суббота', 6],
];

export function parseSchedule(sheet) {
  const header = findGroupHeader(sheet);
  if (!header) throw new Error('Не найдена строка с группами («№ пары» + название группы). Это точно расписание занятий?');

  const pairRows = [];
  let day = 0;
  for (let r = header.row + 1; r < sheet.rowCount; r++) {
    day = dayAt(sheet, r, header.pairCol) || day;
    if (/^\d+$/.test(sheet.get(r, header.pairCol))) pairRows.push({ r, day, pair: Number(sheet.get(r, header.pairCol)), time: timeAt(sheet, r, header.pairCol) });
  }

  const entries = [];
  pairRows.forEach((block, i) => {
    const end = i + 1 < pairRows.length ? pairRows[i + 1].r - 1 : Math.min(sheet.rowCount - 1, block.r + 8);
    for (const group of header.groups) {
      entries.push(...readBlock(sheet, block, end, group));
    }
  });

  return {
    entries,
    ...detectMeta(sheet, header.row),
    groups: header.groups.map(g => g.name),
    hasSaturday: entries.some(e => e.day === 6),
  };
}

// «11:25-12:55» в колонке звонков слева от номера пары → '11:25-12:55'; время — ключ строки в таблице расписания,
// потому что у курсов своя нумерация пар (у 2 курса 1-я пара в 13:25, у 3 курса — в 08:00)
const TIME_RE = /(\d{1,2})[:.](\d{2})\s*[-–—]\s*(\d{1,2})[:.](\d{2})/;
function timeAt(sheet, r, pairCol) {
  for (let c = 0; c < pairCol; c++) {
    const m = TIME_RE.exec(sheet.get(r, c));
    if (m) return `${m[1].padStart(2, '0')}:${m[2]}-${m[3].padStart(2, '0')}:${m[4]}`;
  }
  return '';
}

function findGroupHeader(sheet) {
  for (let r = 0; r < Math.min(sheet.rowCount, 40); r++) {
    const groups = [];
    for (let c = 1; c < sheet.colCount - 1; c++) {
      const text = sheet.get(r, c);
      if (text && !/пар/i.test(text) && /пар/i.test(sheet.get(r, c - 1))) {
        groups.push({ name: text.replace(/\s+/g, ''), left: c, right: c + 1 });
      }
    }
    if (groups.length >= 2) return { row: r, pairCol: groups[0].left - 1, groups };
  }
  return null;
}

function dayAt(sheet, r, pairCol) {
  for (let c = 0; c < pairCol; c++) {
    const text = sheet.get(r, c).toLowerCase();
    const hit = text && DAYS.find(([name]) => text.includes(name));
    if (hit) return hit[1];
  }
  return 0;
}

function readBlock(sheet, block, endRow, group) {
  const cols = [group.left, group.right];
  const subjects = [null, null];
  const teachers = [];

  for (let r = block.r; r <= endRow; r++) {
    cols.forEach((c, side) => {
      const text = sheet.get(r, c);
      if (!text) return;
      const names = teachersOf(text);
      if (names.length) names.forEach(name => teachers.push({ name, side, both: sheet.span(r, c).cols >= 2 }));
      else if (!ROOM.test(text) && !subjects[side]) subjects[side] = text;
    });
  }

  return teachers
    .map(t => ({
      day: block.day,
      pair: block.pair,
      time: block.time,
      group: group.name,
      teacher: t.name,
      week: t.both ? 'both' : (t.side === 0 ? 'left' : 'right'),
      subject: t.both ? (subjects[0] || subjects[1]) : (subjects[t.side] || subjects[1 - t.side]),
    }))
    .filter(e => e.subject);
}

// Шапка расписания: «Расписание учебных занятий 3 курса на 5 семестр 2026-2027 учебного года».
// semester — семестр учебного года (1 — осенний, 2 — весенний): 1, 3, 5, 7 → 1; 2, 4, 6, 8 → 2; 0 — не указан.
function detectMeta(sheet, headerRow) {
  const text = [];
  for (let r = 0; r < headerRow; r++) for (let c = 0; c < sheet.colCount; c++) text.push(sheet.get(r, c));
  const head = text.join(' ');
  const sem = /(\d)\s*[-–]?\s*(?:семестр|жартыжыл)/i.exec(head);
  const course = /(\d)\s*[-–]?\s*курс/i.exec(head);
  const year = /(20\d{2})\s*[-–/]\s*(20\d{2})/.exec(head);
  return {
    semester: sem ? (Number(sem[1]) % 2 === 1 ? 1 : 2) : 0,
    semesterNo: sem ? Number(sem[1]) : 0,
    course: course ? Number(course[1]) : 0,
    startYear: year ? Number(year[1]) : 0,
  };
}
