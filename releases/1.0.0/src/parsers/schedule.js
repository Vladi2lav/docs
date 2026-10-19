// Расписание занятий (xlsx, сетка «пары × группы») → плоский список занятий.
//
// Структура сетки:
//   • строка заголовка: для каждой группы ячейка «№ пары» и правее — название группы;
//     группа занимает две колонки: левая и правая неделя (числитель / знаменатель);
//   • блок пары: строка с номером пары, внутри блока — предмет, преподаватель(и), аудитория;
//   • ячейка, объединённая на обе колонки, означает «каждую неделю».

const L = 'А-ЯЁӘІҢҒҮҰҚӨҺа-яёәіңғүұқөһA-Za-z';
const TEACHER = new RegExp(`^[${L}][${L}\\-]+\\s+[${L}]\\.\\s*(?:[${L}]\\.)?$`);
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
    if (/^\d+$/.test(sheet.get(r, header.pairCol))) pairRows.push({ r, day, pair: Number(sheet.get(r, header.pairCol)) });
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
    semester: detectSemester(sheet, header.row),
    groups: header.groups.map(g => g.name),
    hasSaturday: entries.some(e => e.day === 6),
  };
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
      if (TEACHER.test(text)) teachers.push({ name: text, side, both: sheet.span(r, c).cols >= 2 });
      else if (!ROOM.test(text) && !subjects[side]) subjects[side] = text;
    });
  }

  return teachers
    .map(t => ({
      day: block.day,
      pair: block.pair,
      group: group.name,
      teacher: t.name,
      week: t.both ? 'both' : (t.side === 0 ? 'left' : 'right'),
      subject: t.both ? (subjects[0] || subjects[1]) : (subjects[t.side] || subjects[1 - t.side]),
    }))
    .filter(e => e.subject);
}

// «… 1 семестр …» в шапке расписания → 1 или 2 (0 — не указано)
function detectSemester(sheet, headerRow) {
  for (let r = 0; r < headerRow; r++) {
    for (let c = 0; c < sheet.colCount; c++) {
      const m = /([12])\s*[-–]?\s*(?:семестр|жартыжыл)/i.exec(sheet.get(r, c));
      if (m) return Number(m[1]);
    }
  }
  return 0;
}
