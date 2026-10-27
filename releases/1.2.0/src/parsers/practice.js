// Расписание учебной практики (.docx) → группа, модуль, руководитель и часы по датам.
//
// Документ целиком состоит из таблиц; строки разбираются по подписям в первой колонке
// («Группа», «Модуль», «Руководитель практики») и по датам вида дд.мм.гггг.

import { parseRuDate } from '../util/dates.js';
import { parsePairsText } from '../domain/practice-days.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export async function parsePractice(JSZip, arrayBuffer) {
  let xml;
  try {
    const zip = await JSZip.loadAsync(arrayBuffer);
    xml = await zip.file('word/document.xml').async('string');
  } catch {
    throw new Error('Не удалось прочитать файл. Нужен формат .docx — старый .doc пересохраните в Word как .docx');
  }

  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const rows = Array.from(doc.getElementsByTagNameNS(W, 'tr')).map(tr =>
    Array.from(tr.childNodes).filter(n => n.localName === 'tc').map(cellText));

  const labelled = label => {
    for (const cells of rows) {
      const i = cells.findIndex(t => t.toLowerCase().startsWith(label.toLowerCase()));
      if (i >= 0) {
        const rest = cells[i].slice(label.length).trim();
        return rest || cells.slice(i + 1).find(Boolean) || '';
      }
    }
    return '';
  };

  const days = {};
  const slots = {};              // «4-6 пара» в колонке времени → [4, 5, 6]
  let total = null;
  for (const cells of rows) {
    const dateCell = cells.find(t => /^\d{1,2}\.\d{1,2}\.\d{4}$/.test(t));
    if (dateCell) {
      const hours = Number(cells.slice(cells.indexOf(dateCell) + 1).reverse().find(t => /^\d+([.,]\d+)?$/.test(t))?.replace(',', '.'));
      const date = parseRuDate(dateCell);
      days[date] = Number.isFinite(hours) ? hours : 0;
      const pairs = cells.map(parsePairsText).find(Boolean);
      if (pairs && days[date] > 0) slots[date] = pairs;
    } else if (/^итого/i.test(cells.find(Boolean) || '')) {
      const n = Number([...cells].reverse().find(t => /^\d+$/.test(t)));
      if (Number.isFinite(n)) total = n;
    }
  }

  const dates = Object.keys(days).sort();
  if (!dates.length) throw new Error('В документе не найдена таблица с датами практики');

  const moduleText = labelled('Модуль');
  const mod = /^([A-ZА-ЯЁ]{2,3}\d{1,2})\s*\.?\s*(.*)$/.exec(moduleText);
  const supervisorCells = rows.find(cells => cells.some(t => /^руководитель практики/i.test(t))) || [];
  const NAME = /[A-ZА-ЯЁӘІҢҒҮҰҚӨҺ]\.\s*[A-Za-zА-Яа-яЁёӘәІіҢңҒғҮүҰұҚқӨөҺһ-]+/;
  const nameAfter = label => {
    const cells = rows.find(cs => cs.some(t => t.toLowerCase().startsWith(label)));
    return (cells && cells.map(t => NAME.exec(t)?.[0]).find(Boolean)) || '';
  };
  const approver = (rows.flat().find(t => /^утверждаю/i.test(t)) || '').match(NAME)?.[0] || '';

  return {
    group: labelled('Группа').replace(/\s+/g, ''),
    moduleIndex: mod ? mod[1] : '',
    moduleTitle: mod ? mod[2].trim() : moduleText,
    specialty: [labelled('Специальность'), labelled('Квалификация')].filter(Boolean).join(', '),
    specialtyOnly: labelled('Специальность'),
    qualification: labelled('Квалификация'),
    course: Number(labelled('Курс')) || 0,
    signers: { approver, upr: nameAfter('заместитель директора по учебно-производственной'), head: nameAfter('заведующий учебной частью') },
    supervisor: supervisorCells.find(t => /^[A-ZА-ЯЁӘІҢҒҮҰҚӨҺ]\.\s*[A-Za-zА-Яа-яЁёӘәІіҢңҒғҮүҰұҚқӨөҺһ-]+$/.test(t)) || '',
    days,
    slots,
    from: dates[0],
    to: dates[dates.length - 1],
    total: total ?? Object.values(days).reduce((a, b) => a + b, 0),
  };
}

function cellText(tc) {
  const parts = [];
  const walk = n => {
    if (n.localName === 't') parts.push(n.textContent);
    else if (n.localName === 'tab') parts.push(' ');
    Array.from(n.childNodes || []).forEach(ch => walk(ch));
  };
  walk(tc);
  return parts.join('').replace(/\s+/g, ' ').trim();
}
