// Примитивы WordprocessingML (абзацы, таблицы, ячейки) — общие для всех Word-документов приложения.
// Порядок дочерних элементов соответствует схеме OOXML.

export const esc = s => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

export function run(text, { b, i, u, hl, sz } = {}) {
  const props = (b ? '<w:b/>' : '') + (i ? '<w:i/>' : '') + (sz ? `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>` : '') +
    (hl ? `<w:highlight w:val="${hl}"/>` : '') + (u ? '<w:u w:val="single"/>' : '');
  return `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}
export const br = () => '<w:r><w:br/></w:r>';

export function para(runs, { jc, ind, pageBreak, before } = {}) {
  const props = (pageBreak ? '<w:pageBreakBefore/>' : '') +
    (before ? `<w:spacing w:before="${before}"/>` : '') +
    (ind ? `<w:ind w:left="${ind}"/>` : '') +
    (jc ? `<w:jc w:val="${jc}"/>` : '');
  return `<w:p>${props ? `<w:pPr>${props}</w:pPr>` : ''}${runs.join('')}</w:p>`;
}

const BORDERS = ['top', 'left', 'bottom', 'right'].map(s => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('');
const NO_BORDERS = ['top', 'left', 'bottom', 'right'].map(s => `<w:${s} w:val="nil"/>`).join('');

export function cell(content, width, { span, vMerge, vertical, fill, noBorder, diag, valign = 'center' } = {}) {
  const props = `<w:tcW w:w="${width}" w:type="dxa"/>` +
    (span > 1 ? `<w:gridSpan w:val="${span}"/>` : '') +
    (vMerge ? (vMerge === 'restart' ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>') : '') +
    `<w:tcBorders>${noBorder ? NO_BORDERS : BORDERS}${diag ? '<w:tl2br w:val="single" w:sz="4" w:space="0" w:color="000000"/>' : ''}</w:tcBorders>` +
    (fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>` : '') +
    (vertical ? '<w:textDirection w:val="btLr"/>' : '') +
    `<w:vAlign w:val="${valign}"/>`;
  return `<w:tc><w:tcPr>${props}</w:tcPr>${content}</w:tc>`;
}

export function row(cells, { height } = {}) {
  const props = '<w:cantSplit/>' + (height ? `<w:trHeight w:val="${height}" w:hRule="atLeast"/>` : '');
  return `<w:tr><w:trPr>${props}</w:trPr>${cells.join('')}</w:tr>`;
}

export function table(widths, rows) {
  return `<w:tbl><w:tblPr><w:tblW w:w="${sum(widths)}" w:type="dxa"/><w:tblLayout w:type="fixed"/>` +
    '<w:tblCellMar><w:left w:w="40" w:type="dxa"/><w:right w:w="40" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
    `<w:tblGrid>${widths.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows.join('')}</w:tbl>`;
}

export const sum = list => list.reduce((a, b) => a + b, 0);
