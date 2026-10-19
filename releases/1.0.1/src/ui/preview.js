// HTML-предпросмотр листа месяца (тот же набор данных, что уходит в Word).

import { FORM } from '../config.js';
import { extrasLine } from '../export/docx.js';

export const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const bg = fill => (fill ? ` style="background:#${fill}"` : '');

export function renderSheet(meta, sheet) {
  const n = sheet.days.length;
  const groups = [...new Set(sheet.rows.map(r => r.group))].join(', ');
  const [h0, h1, h2, h3, h4, h5, h6] = FORM.t2;

  const days = sheet.days.map(d => `<th${bg(d.fill)} title="${esc(d.title)}">${d.n}</th>`).join('');
  const dayRows = sheet.rows.map(r => `<tr>
    <td>${r.num}</td><td>${esc(r.moduleIndex)}</td><td class="l">${esc(r.name)}</td><td>${esc(r.group)}</td>
    ${r.cells.map(c => `<td${bg(c.fill)}${c.vertical ? ' class="v"' : ''}>${esc(c.text)}</td>`).join('')}<td><b>${r.total}</b></td></tr>`).join('');
  const sumRows = sheet.rows.map(r => `<tr>
    <td>${esc(r.group)}</td><td class="l">${esc(r.name)}</td><td>${r.plan}</td>
    <td style="background:#FFFF00">${r.total}</td><td>${r.cumulative}</td><td${bg(r.remainingFill)}>${esc(r.remainingText)}</td></tr>`).join('');

  return `<div class="sheet">
    <div class="r">${FORM.stamp}</div>
    <div class="c">${FORM.ministry}<br><u>УО «${esc(meta.college)}»</u><br>${FORM.orgCaption}<br><b>${FORM.title}</b><br>
      за ${esc(meta.academicYear)} учебный год<br><u class="hl-green">${esc(meta.teacherFullName.toUpperCase())}</u><br>${FORM.nameCaption}${extrasLine(meta.extras) ? `<br>${esc(extrasLine(meta.extras))}` : ''}</div>
    <div class="info"><span><b>${FORM.specialty}</b> <u>${esc(meta.specialty)}</u></span><span>${FORM.group} <u>${esc(groups)}</u></span></div>
    <div class="scroll"><table class="grid days">
      <tr><th rowspan="2">${FORM.t1[0]}</th><th rowspan="2">${FORM.t1[1]}</th><th rowspan="2">${FORM.t1[2]}</th><th rowspan="2">${FORM.t1[3]}</th>
          <th colspan="${n}">${FORM.month} <u class="hl-yellow">${sheet.name}</u></th><th rowspan="2">${FORM.t1[4]}</th></tr>
      <tr>${days}</tr>${dayRows}</table></div>
    <div>&nbsp;&nbsp;&nbsp;${FORM.cont}</div>
    <table class="grid sum">
      <tr><th rowspan="2">${h0}</th><th rowspan="2">${h1}</th><th rowspan="2">${h2}</th><th colspan="2">${h3}</th><th rowspan="2">${h6}</th></tr>
      <tr><th>${h4}</th><th>${h5}</th></tr>${sumRows}</table>
    <p>&nbsp;&nbsp;&nbsp;${FORM.total} <u class="hl-yellow">${sheet.total} часа(ов)</u></p>
    ${sheet.notes.map(t => `<p>&nbsp;&nbsp;&nbsp;${FORM.makeup}: ${esc(t)}</p>`).join('')}
    <p class="sign">${FORM.teacher} ___________________________<br>${FORM.registrar} __________________<br>
      ${FORM.deputy} ______________________________ ${esc(meta.deputy)}</p>
    <p class="note">${FORM.note}</p>
  </div>`;
}
