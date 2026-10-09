// Ведомость 1 семестра (.docx формата «Форма Приказ №130», в том числе созданная этим приложением) →
// то, что переносится во 2 семестр: ФИО, специальность и по каждой строке — план, выполнено с начала года, остаток.
// Берётся последняя встреченная строка (последний месяц документа), поэтому нужны страницы до конца семестра.

import { MONTH_NAMES } from '../config.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export async function parsePrevSemester(JSZip, arrayBuffer) {
  let xml;
  try {
    const zip = await JSZip.loadAsync(arrayBuffer);
    xml = await zip.file('word/document.xml').async('string');
  } catch {
    throw new Error('Не удалось прочитать файл. Нужна ведомость 1 семестра в формате .docx');
  }
  const doc = new DOMParser().parseFromString(xml, 'application/xml');

  const textOf = (node, br) => {
    const parts = [];
    const walk = n => {
      if (n.localName === 't') parts.push(n.textContent);
      else if (n.localName === 'br') parts.push(br);
      Array.from(n.childNodes || []).forEach(walk);
    };
    walk(node);
    return parts.join('');
  };

  // ФИО: строка после «за … учебный год» в том же абзаце
  let fullName = '';
  for (const p of Array.from(doc.getElementsByTagNameNS(W, 'p'))) {
    const lines = textOf(p, '\n').split('\n').map(l => l.trim()).filter(Boolean);
    const i = lines.findIndex(l => /учебный год/i.test(l));
    if (i >= 0 && lines[i + 1]) { fullName = titleCase(lines[i + 1]); break; }
  }

  const rows = Array.from(doc.getElementsByTagNameNS(W, 'tr')).map(tr =>
    Array.from(tr.childNodes).filter(n => n.localName === 'tc').map(tc => textOf(tc, ' ').replace(/\s+/g, ' ').trim()));

  let specialty = '';
  for (const cells of rows) {
    const c = cells.find(t => /^Специальность/i.test(t));
    if (c) { specialty = c.replace(/^Специальность\s*\(Квалификация\)\s*/i, '').replace(/Учебная группа.*$/i, '').trim(); break; }
  }

  // Таблица 1 (по месяцам): «Месяц СЕНТЯБРЬ» | № | индекс | наименование | группа | дни… | итого;
  // Таблица 2: группа | наименование | план | за месяц | с начала года | остаток
  const modules = {};
  const monthly = {};            // группа|наименование → { индекс месяца: часы }
  const byKey = new Map();
  let month = -1;
  for (const c of rows) {
    const head = c.find(t => /^Месяц\s/i.test(t));
    if (head) { month = MONTH_NAMES.findIndex(n => head.toLowerCase().includes(n.toLowerCase())); continue; }
    if (c.length >= 6 && /^\d+$/.test(c[0]) && /^[A-ZА-ЯЁӘІҢҒҮҰҚӨҺ]{1,6}\d/i.test(c[3]) && /^\d+([.,]\d+)?$/.test(c[c.length - 1])) {
      if (/^[A-ZА-ЯЁ]{2,3}\d+$/.test(c[1])) modules[`${c[3]}|${c[2]}`] = c[1];
      if (month >= 0) (monthly[`${c[3]}|${c[2]}`] ||= {})[month] = Number(c[c.length - 1].replace(',', '.'));
    }
    if (c.length === 6 && /^\d*$/.test(c[2]) && /^\d+$/.test(c[4]) && /^[A-ZА-ЯЁӘІҢҒҮҰҚӨҺ]{1,6}\d/i.test(c[0])) {
      byKey.set(`${c[0]}|${c[1]}`, {
        group: c[0].replace(/\s+/g, ''), name: c[1], key: `${c[0]}|${c[1]}`,
        plan: c[2] === '' ? null : Number(c[2]), cumulative: Number(c[4]),
        remaining: /^\d+/.test(c[5]) ? Number(/^\d+/.exec(c[5])[0]) : null,
      });
    }
  }
  for (const r of byKey.values()) { r.moduleIndex = modules[r.key] || ''; r.months = monthly[r.key] || {}; }
  if (!fullName && !byKey.size) throw new Error('Это не похоже на ведомость формы №130 (не найдены ФИО и таблица часов)');
  return { fullName, specialty, rows: [...byKey.values()] };
}

const titleCase = s => s.toLowerCase().replace(/(^|[\s-])(\S)/g, (_, a, b) => a + b.toUpperCase());
