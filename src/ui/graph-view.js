// Вкладка «График учебного процесса»: группы × недели с кодами, как в исходном документе.

import { VACATION_CODE } from '../config.js';
import { normGroup } from '../parsers/graph.js';
import { ruDate } from '../util/dates.js';
import { esc } from './preview.js';

const PALETTE = ['#fde68a', '#fecaca', '#bfdbfe', '#ddd6fe', '#fbcfe8', '#a7f3d0', '#fed7aa', '#bae6fd', '#e9d5ff', '#d9f99d'];

export function renderGraph(container, graph, { highlight = new Set(), query = '', onlyHighlighted = false, semesters = [] } = {}) {
  const codes = [...new Set(Object.values(graph.groups).flatMap(g => Object.values(g.codes)).filter(Boolean))].sort();
  const color = code => (code === VACATION_CODE ? '#c5e0b3' : PALETTE[codes.indexOf(code) % PALETTE.length]);
  const q = query.trim().toLowerCase();

  const head = graph.weeks.map(w => `<th title="${ruDate(w.start)} – ${ruDate(w.end)}">${w.num}<small>${w.start.slice(8)}.${w.start.slice(5, 7)}</small></th>`).join('');
  const body = Object.entries(graph.groups)
    .filter(([key, g]) => (!q || g.name.toLowerCase().includes(q)) && (!onlyHighlighted || highlight.has(key)))
    .map(([key, g]) => `<tr class="${highlight.has(key) ? 'mine' : ''}"><th class="name">${esc(g.name)}</th>${
      graph.weeks.map(w => {
        const code = g.codes[w.num];
        return `<td${code ? ` style="background:${color(code)}"` : ''} title="${esc(code || 'теория')}">${esc(code)}</td>`;
      }).join('')}</tr>`).join('');

  const legend = Object.entries(graph.legend).map(([family, items]) =>
    `<div><b>${esc(family)}</b>: ${Object.entries(items).map(([code, m]) =>
      `<span class="chip" style="background:${color(code)}">${esc(code)}</span> ${esc(m.index)} ${esc(m.title)}`).join('; ')}</div>`).join('');

  const info = semesters.map((s, i) => `${i + 1} семестр: ${s.from.m + 1}.${s.from.y} – ${s.to.m + 1}.${s.to.y}`).join(' · ');

  container.innerHTML = `${info ? `<p class="hint">Семестры по каникулам графика для групп педагога — ${info}</p>` : ''}
    <div class="scroll graph"><table><thead><tr><th class="name">Группа</th>${head}</tr></thead><tbody>${body}</tbody></table></div>
    <p class="hint">Пустая ячейка — теория; код — практика/каникулы/аттестация; <span class="chip" style="background:#c5e0b3">К</span> — каникулы.</p>
    <details><summary>Расшифровка кодов (из «Обозначений» графика)</summary>${legend || '<p class="hint">В графике не найдены.</p>'}</details>`;
}

export const highlightSet = groups => new Set(groups.map(normGroup));
