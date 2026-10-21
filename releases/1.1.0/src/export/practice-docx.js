// Расписание проведения учебной практики (по образцу документа колледжа) → .docx.
// Одна страница A4 на период практики: шапка, реквизиты группы, таблица «день недели · дата · пары · часы», подписи.

import { documentDays, DOW_FULL } from '../domain/practice-days.js';
import { ruDate } from '../util/dates.js';
import { PACKAGE } from './docx-package.js';
import { br, cell, para, row, run, sum, table } from './ooxml.js';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const SZ = 24;                        // 12 pt
const BLANK = '____________';

const t = (text, o = {}) => run(text, { sz: SZ, ...o });

// «Зиаятдинов Владимир Рашидович» / «Зиаятдинов В.Р.» → «В.Зиаятдинов»; без инициала — просто фамилия
export function shortSigner(name) {
  const [surname = '', second = ''] = String(name || '').trim().split(/\s+/);
  return surname ? `${second ? `${second[0].toUpperCase()}.` : ''}${surname}` : '';
}

// meta: { college, yearLabel, startYear, group, course, specialty, qualification, moduleLine,
//         supervisor, approver, upr, head }; blocks — блоки практики с days / slots; calendar — для праздников
export async function buildPracticeDocx(JSZip, meta, blocks, calendar) {
  const body = blocks.map((b, i) => periodXml(meta, b, calendar, i > 0)).join('');
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="567" w:right="567" w:bottom="567" w:left="1134" w:header="279" w:footer="306" w:gutter="0"/></w:sectPr></w:body></w:document>`;

  const zip = new JSZip();
  zip.file('[Content_Types].xml', PACKAGE.contentTypes);
  zip.file('_rels/.rels', PACKAGE.rels);
  zip.file('word/_rels/document.xml.rels', PACKAGE.documentRels);
  zip.file('word/styles.xml', PACKAGE.styles);
  zip.file('word/document.xml', documentXml);
  return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}

function periodXml(meta, block, calendar, newPage) {
  const days = documentDays(block, calendar);
  const total = sum(days.map(d => d.hours));
  const dateLine = `____ ___________ ${meta.startYear || '____'} года`;

  return [
    para([t('МИНИСТЕРСТВО ПРОСВЕЩЕНИЯ РЕСПУБЛИКИ КАЗАХСТАН', { b: true }), br(), t(`Учреждение образования «${meta.college || ''}»`)], { jc: 'center', pageBreak: newPage }),
    para([t('Утверждаю'), br(), t('Заместитель директора по учебной работе'), br(), t(`${BLANK} ${meta.approver || ''}`), br(), t(dateLine)], { ind: 5500, before: 120 }),
    para([t('РАСПИСАНИЕ', { b: true }), br(), t('проведения учебной практики обучающихся', { b: true }), br(), t(`на ${meta.yearLabel} учебный год`)], { jc: 'center', before: 240 }),
    infoTable(meta, block),
    para([]),
    dayTable(days, total),
    para([]),
    signTable(meta, dateLine),
  ].join('');
}

// Реквизиты без рамок: подпись слева, значение справа (подчёркнуто)
function infoTable(meta, block) {
  const w = [3300, 2300, 1500, 3370];
  const plain = (text, width, o = {}) => cell(para([t(text, o)]), width, { noBorder: true });
  const wide = (label, value) => row([plain(label, w[0]), cell(para([t(value || '', { u: true })]), w[1] + w[2] + w[3], { span: 3, noBorder: true })]);
  return table(w, [
    wide('Специальность', meta.specialty),
    wide('Квалификация', meta.qualification),
    row([plain('Курс', w[0]), plain(meta.course ? String(meta.course) : '', w[1], { u: true }), plain('Группа', w[2]), plain(meta.group || '', w[3], { u: true })]),
    wide('Модуль', meta.moduleLine),
    row([plain('Сроки прохождения практики', w[0]), plain(`с ${ruDate(block.from)}`, w[1], { u: true }), plain('по', w[2]), plain(ruDate(block.to), w[3], { u: true })]),
  ]);
}

function dayTable(days, total) {
  const w = [2400, 2400, 3000, 2670];
  const c = (text, i, o = {}) => cell(para([t(text, o)], { jc: 'center' }), w[i]);
  const head = row(['День недели', 'Дата проведения', 'Время проведения', 'Количество часов'].map((h, i) => c(h, i, { b: true })));
  const rows = days.map(d => row([c(DOW_FULL[d.dow], 0), c(ruDate(d.date), 1), c(d.text, 2), c(d.hours ? String(d.hours) : '', 3)]));
  const foot = row([cell(para([t('Итого:', { b: true })], { jc: 'right' }), w[0] + w[1] + w[2], { span: 3 }), c(String(total), 3, { b: true })]);
  return table(w, [head, ...rows, foot]);
}

// Подписи: «Должность  ____________  И.Фамилия» и строка даты под каждой
function signTable(meta, dateLine) {
  const w = [5200, 2200, 3070];
  const person = (label, name) => [
    row([cell(para([t(label)]), w[0], { noBorder: true }), cell(para([t(BLANK)], { jc: 'center' }), w[1], { noBorder: true }), cell(para([t(name || '')]), w[2], { noBorder: true })]),
    row([cell(para([t(dateLine)]), w[0] + w[1] + w[2], { span: 3, noBorder: true })]),
  ];
  return table(w, [
    ...person('Руководитель практики', meta.supervisor),
    row([cell(para([t('Согласовано')]), w[0] + w[1] + w[2], { span: 3, noBorder: true })]),
    ...person('Заместитель директора по учебно-производственной работе', meta.upr),
    ...person('Заведующий учебной частью', meta.head),
  ]);
}
