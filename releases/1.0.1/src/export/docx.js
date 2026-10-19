// Ведомость (Форма Приказ №130) → .docx. Одна страница A4 (альбомная) на месяц.
// Порядок дочерних элементов соответствует схеме OOXML (иначе Word может счесть файл повреждённым).

import { FORM, FILL } from '../config.js';
import { PACKAGE } from './docx-package.js';
import { br, cell, para, row, run, sum, table } from './ooxml.js';

const PAGE_W = 16838;
const MARGIN = 567;
const BODY_W = PAGE_W - 2 * MARGIN;
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

export async function buildDocx(JSZip, meta, sheets) {
  const body = sheets.map((sheet, i) => monthXml(meta, sheet, i > 0)).join('');
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>${body}<w:sectPr><w:pgSz w:w="${PAGE_W}" w:h="11906" w:orient="landscape"/><w:pgMar w:top="${MARGIN}" w:right="${MARGIN}" w:bottom="${MARGIN}" w:left="${MARGIN}" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr></w:body></w:document>`;

  const zip = new JSZip();
  zip.file('[Content_Types].xml', PACKAGE.contentTypes);
  zip.file('_rels/.rels', PACKAGE.rels);
  zip.file('word/_rels/document.xml.rels', PACKAGE.documentRels);
  zip.file('word/styles.xml', PACKAGE.styles);
  zip.file('word/document.xml', documentXml);
  return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}

// ───────── разметка месяца ─────────

function monthXml(meta, sheet, newPage) {
  return [
    para([run(FORM.stamp)], { jc: 'right', pageBreak: newPage }),
    para([run(FORM.ministry)], { jc: 'center' }),
    para([run(`УО «${meta.college}»`, { u: true }), br(), run(FORM.orgCaption), br(), run(FORM.title, { b: true })], { jc: 'center' }),
    para([run(`за ${meta.academicYear} учебный год`), br(), run(meta.teacherFullName.toUpperCase(), { hl: 'green', u: true }), br(), run(FORM.nameCaption)], { jc: 'center' }),
    ...extrasXml(meta),
    dayTable(meta, sheet),
    para([run(FORM.cont)], { ind: 708 }),
    summaryTable(sheet),
    ...signatureXml(meta, sheet),
  ].join('');
}

// Необязательная строка «Отделение · ЦМК · Должность» — только если хоть что-то заполнено
function extrasXml(meta) {
  const line = extrasLine(meta.extras);
  return line ? [para([run(line)], { jc: 'center' })] : [];
}

export function extrasLine(extras = {}) {
  return Object.entries(FORM.extras).filter(([k]) => extras[k]).map(([k, label]) => `${label}: ${extras[k]}`).join('; ');
}

function dayTable(meta, sheet) {
  const n = sheet.days.length;
  const fixed = [440, 560, 2700, 640];
  const totalW = 640;
  const dayW = Math.floor((BODY_W - sum(fixed) - totalW) / n);
  const widths = [...fixed, ...Array(n).fill(dayW), totalW];
  widths[2] += BODY_W - sum(widths);                       // остаток — в колонку наименования

  const groups = [...new Set(sheet.rows.map(r => r.group))].join(', ');
  const topLeft = widths.length - 8;
  const info = row([
    cell(para([run(FORM.specialty + ' ', { b: true }), run(meta.specialty, { u: true })]), sum(widths.slice(0, topLeft)), { span: topLeft, noBorder: true }),
    cell(para([run(`${FORM.group} `), run(groups, { u: true })], { jc: 'right' }), sum(widths.slice(topLeft)), { span: 8, noBorder: true }),
  ]);

  const head1 = row([
    ...FORM.t1.slice(0, 4).map((t, i) => cell(para([run(t, { sz: i === 1 || i === 3 ? 18 : 20 })], { jc: 'center' }), widths[i], { vMerge: 'restart', vertical: i === 1 || i === 3 })),
    cell(para([run(`${FORM.month} `), run(sheet.name, { hl: 'yellow', u: true })], { jc: 'center' }), sum(widths.slice(4, 4 + n)), { span: n }),
    cell(para([run(FORM.t1[4])], { jc: 'center' }), totalW, { vMerge: 'restart', vertical: true }),
  ], { height: 1300 });

  const head2 = row([
    ...widths.slice(0, 4).map(w => cell(para([]), w, { vMerge: 'continue' })),
    ...sheet.days.map((d, i) => cell(para([run(String(d.n), { sz: 16 })], { jc: 'center' }), widths[4 + i], { fill: d.fill })),
    cell(para([]), totalW, { vMerge: 'continue' }),
  ]);

  const body = sheet.rows.map(r => row([
    cell(para([run(String(r.num), { sz: 18 })], { jc: 'center' }), widths[0]),
    cell(para([run(r.moduleIndex, { sz: 18 })], { jc: 'center' }), widths[1]),
    cell(para([run(r.name, { sz: 18 })], { jc: 'center' }), widths[2]),
    cell(para([run(r.group, { sz: 18 })], { jc: 'center' }), widths[3]),
    ...r.cells.map((c, i) => cell(para([run(c.text, { sz: c.vertical ? 12 : 16 })], { jc: 'center' }), widths[4 + i], { fill: c.fill, vertical: c.vertical })),
    cell(para([run(String(r.total), { sz: 18 })], { jc: 'center' }), totalW),
  ]));

  return table(widths, [info, head1, head2, ...body]);
}

function summaryTable(sheet) {
  const widths = [1400, 0, 2300, 1900, 1900, 1500];
  widths[1] = BODY_W - sum(widths);
  const c = (text, w, opts = {}) => cell(para([run(text, { sz: opts.sz })], { jc: 'center' }), w, opts);
  const [h0, h1, h2, h3, h4, h5, h6] = FORM.t2;

  const head1 = row([
    c(h0, widths[0], { vMerge: 'restart' }), c(h1, widths[1], { vMerge: 'restart' }), c(h2, widths[2], { vMerge: 'restart' }),
    c(h3, widths[3] + widths[4], { span: 2 }), c(h6, widths[5], { vMerge: 'restart' }),
  ]);
  const head2 = row([
    cell(para([]), widths[0], { vMerge: 'continue' }), cell(para([]), widths[1], { vMerge: 'continue' }), cell(para([]), widths[2], { vMerge: 'continue' }),
    c(h4, widths[3]), c(h5, widths[4]), cell(para([]), widths[5], { vMerge: 'continue' }),
  ]);
  const body = sheet.rows.map(r => row([
    c(r.group, widths[0]),
    c(r.name, widths[1]),
    c(String(r.plan), widths[2]),
    c(String(r.total), widths[3], { fill: FILL.mark }),
    c(String(r.cumulative), widths[4]),
    c(r.remainingText, widths[5], { fill: r.remainingFill }),
  ]));
  return table(widths, [head1, head2, ...body]);
}

function signatureXml(meta, sheet) {
  const line = (text, left = 708) => para([run(text)], { ind: left });
  return [
    para([run(`${FORM.total} `), run(`${sheet.total} часа(ов)`, { hl: 'yellow', u: true })], { ind: 708, before: 120 }),
    ...sheet.notes.map(text => para([run(`${FORM.makeup}: ${text}`)], { ind: 708 })),
    line(`${FORM.teacher} ___________________________`),
    line(FORM.sign, 2600),
    line(`${FORM.registrar} __________________`),
    line(`${FORM.deputy} ______________________________ ${meta.deputy}`),
    line(FORM.sign, 9000),
    para([run(FORM.note)], { ind: 354 }),
  ];
}
