// Годовой учёт часов (нагрузка ГРАНТ / комм / ВАКАНТ …) → .docx.
// На каждый тип нагрузки две страницы: «месяцы × группы» и «Дополнительные сведения к годовому учету часов».

import { MONTH_NAMES } from '../config.js';
import { YEAR_MONTHS, groupByLoad } from '../domain/annual.js';
import { PACKAGE } from './docx-package.js';
import { cell, para, row, run, sum, table } from './ooxml.js';

const PAGE_W = 16838;
const MARGIN = 720;
const BODY_W = PAGE_W - 2 * MARGIN;
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const GRAY = 'D9D9D9';
const MIN_COLUMNS = 10;      // как в бланке: не менее 10 столбцов групп
const MIN_DETAIL_ROWS = 8;

const num = n => (n ? String(Math.round(n * 100) / 100).replace('.', ',') : '');
const or0 = n => (n ? num(n) : 'нет');

// meta: { yearLabel, fullGen, collegeGen, shortName, notes: { [тип]: причина } }
export async function buildAnnualDocx(JSZip, entries, meta) {
  const pages = groupByLoad(entries);
  const body = pages.map((g, i) => monthsPage(g, meta, i > 0) + detailsPage(g, meta)).join('');
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>${body}<w:sectPr><w:pgSz w:w="${PAGE_W}" w:h="11906" w:orient="landscape"/><w:pgMar w:top="426" w:right="${MARGIN}" w:bottom="${MARGIN}" w:left="${MARGIN}" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`;

  const zip = new JSZip();
  zip.file('[Content_Types].xml', PACKAGE.contentTypes);
  zip.file('_rels/.rels', PACKAGE.rels);
  zip.file('word/_rels/document.xml.rels', PACKAGE.documentRels);
  zip.file('word/styles.xml', PACKAGE.styles);
  zip.file('word/document.xml', documentXml);
  return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}

const title = e => (e.kind === 'practice' ? `${e.group} (УП)` : e.group);
const fullName = e => {
  const base = [e.moduleIndex, e.name.replace(/^УП\s*-\s*/i, '')].filter(Boolean).join(' ');
  return e.kind === 'practice' ? `Учебная практика: ${base}` : base;
};

// ───────── страница 1: месяцы × группы ─────────

function monthsPage(g, meta, newPage) {
  const cols = Math.max(MIN_COLUMNS, g.entries.length);
  const first = 2400;
  const colW = Math.floor((BODY_W - first) / cols);
  const widths = [BODY_W - colW * cols, ...Array(cols).fill(colW)];
  const pad = list => [...list, ...Array(cols - list.length).fill('')];
  const sz = 22;
  const c = (text, w, opts = {}) => cell(para([run(text, { sz })], { jc: opts.left ? 'left' : 'center' }), w, { valign: 'top', ...opts });
  const line = (label, values, opts = {}) => row([c(label, widths[0], { left: true, fill: opts.fill }), ...pad(values).map((v, i) => c(v, widths[i + 1], { fill: opts.fill }))], { height: opts.height || 330 });

  const header = row([
    cell(para([run('Группы', { sz })], { jc: 'right' }) + para([run('Месяцы', { sz })], { jc: 'left' }), widths[0], { diag: true, valign: 'top' }),
    ...pad(g.entries.map(title)).map((t, i) => c(t, widths[i + 1], { valign: 'top' })),
  ], { height: 760 });

  const monthRows = YEAR_MONTHS.map(m => line(MONTH_NAMES[m], g.entries.map(e => num(e.months[m]))));
  const lines = [
    header, ...monthRows,
    line('Экзамены (заносятся на основании экзаменационной ведомости)', [], { fill: GRAY, height: 1000 }),
    line('Консультации', [], { fill: GRAY }),
    line('Всего запланировано, часов', g.entries.map(e => num(e.plan)), { height: 560 }),
    line('Фактически выполнено, часов', g.entries.map(e => num(e.fact)), { height: 900 }),
  ];
  // у «Экзаменов» ячейки внутри — серые, как в бланке; пустые значения тоже должны быть серыми
  const note = (g.notDone > 0 ? `${num(g.notDone)}${meta.notes[g.type.key] ? ` (${meta.notes[g.type.key].trim()})` : ''}` : 'нет');
  const text = t => para([run(t, { sz: 20 })], { ind: 720 });

  return [
    para([run('МИНИСТЕРСТВО ПРОСВЕЩЕНИЯ РЕСПУБЛИКИ КАЗАХСТАН', { sz: 20 })], { jc: 'center', pageBreak: newPage }),
    para([run('Ведомость учета учебного времени преподавателей за год ', { sz: 20 }), run(`(${g.type.label})`, { sz: 20, hl: 'yellow' })], { jc: 'center' }),
    para([run(meta.collegeGen, { sz: 20 })], { jc: 'center' }),
    para([run('_'.repeat(63), { sz: 20 })], { jc: 'center' }),
    para([run(`Годовой учет часов, данных преподавателем в ${meta.yearLabel} учебном году`, { sz: 20 })], { jc: 'center' }),
    para([run(meta.fullGen, { sz: 20, i: true })], { jc: 'center', before: 0 }),
    para([]),
    table(widths, lines),
    text(`Всего часов и (или) кредитов по плану: ${num(g.plan)}`),
    text(`Не выполнено часов и (или) кредитов: ${note}`),
    text(`Дано часов сверх плана: ${or0(g.over)}`),
    text(`Всего дано за год часов и (или) кредитов: ${num(g.fact)}`),
    text('Заместитель руководителя по учебной работе ______________________'),
    para([run('(подпись)', { sz: 20 })], { ind: 4200 }),
  ].join('');
}

// ───────── страница 2: дополнительные сведения ─────────

function detailsPage(g, meta) {
  const w = [1500, 5400, 800, 800, 1100, 1100, 1100, 1100, 1100, 1100, 1198];     // 11 столбцов, сумма = BODY_W
  w[1] += BODY_W - sum(w);
  const sz = 22;
  const c = (text, i, opts = {}) => cell(para([run(text, { sz })], { jc: opts.left ? 'left' : 'center' }), opts.width || w[i], { valign: 'top', ...opts });
  const span = (a, b) => sum(w.slice(a, b));
  const blank = i => cell(para([]), w[i], { vMerge: 'continue' });

  const head1 = row([
    c('№ учебной группы', 0, { vMerge: 'restart' }), c('Наименование дисциплины и (или) модулей', 1, { vMerge: 'restart' }),
    c('Количество часов', 2, { span: 2, width: span(2, 4), vMerge: 'restart' }), c('из них часы', 4, { span: 6, width: span(4, 10) }),
    c('Общее количество часов', 10, { vMerge: 'restart' }),
  ]);
  const head2 = row([
    blank(0), blank(1), cell(para([]), span(2, 4), { span: 2, vMerge: 'continue' }),
    c('факультатива', 4, { span: 2, width: span(4, 6) }), c('консультаций', 6, { span: 2, width: span(6, 8) }), c('экзаменов', 8, { span: 2, width: span(8, 10) }),
    blank(10),
  ]);
  const head3 = row([blank(0), blank(1), ...['план', 'факт', 'план', 'факт', 'план', 'факт', 'план', 'факт'].map((t, i) => c(t, i + 2)), blank(10)]);

  const filled = g.entries.map(e => row([
    c(e.group, 0, { left: true }), c(fullName(e), 1, { left: true }), c(num(e.plan), 2, { left: true }), c(num(e.fact), 3, { left: true }),
    c('', 4), c('', 5), c('', 6), c('', 7), c('', 8), c('', 9), c(num(e.fact), 10, { left: true }),
  ], { height: 560 }));
  const empty = Array.from({ length: Math.max(0, MIN_DETAIL_ROWS - g.entries.length) }, () => row(w.map((_, i) => c('', i)), { height: 420 }));
  const text = t => para([run(t, { sz: 20 })], { ind: 720 });

  return [
    para([run('Дополнительные сведения к годовому учету часов преподавателя', { sz: 24 })], { jc: 'center', pageBreak: true }),
    para([run(meta.fullGen, { sz: 20, i: true })], { jc: 'center' }),
    para([]),
    table(w, [head1, head2, head3, ...filled, ...empty]),
    para([]),
    para([run('Фамилия, ', { sz: 20, hl: 'yellow' }), run(`имя, отчество преподавателя : ${meta.shortName} ___________`, { sz: 20, i: true, hl: 'yellow' }), run(' (подпись)', { sz: 20 })], { ind: 720 }),
    text('Офис-регистратор _______________________________'),
    text('Проверено'),
    text('Заместитель директора по учебной работе _____________ (подпись)'),
  ].join('');
}

