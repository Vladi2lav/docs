// Таблица недели «звонок × день» для левой и правой недели (как лист «ШАБЛОН» в Excel педагога), в интерфейсе.
// Пары выбранного педагога подсвечены, остальные (если показаны группы целиком) — приглушены.

import { DAY_NAMES } from '../domain/schedules.js';
import { esc } from './preview.js';

// grid — результат buildGrid; showEmpty — показывать и звонки без занятий; swap — подпись, какая неделя первая
export function renderGrid(container, grid, { showEmpty = false, swap = false } = {}) {
  const rows = grid.rows.map((r, i) => ({ r, i })).filter(({ r }) => showEmpty || r.used);
  if (!rows.length) { container.innerHTML = '<p class="hint">В выбранных группах нет занятий.</p>'; return; }
  const sat = grid.days.map((d, c) => ({ d, c })).filter(({ d, c }) => d !== 6 || rows.some(({ i }) => grid.cells.left[i][c].length || grid.cells.right[i][c].length));
  const side = (name, label) => `<th colspan="${sat.length}" class="week">${label}</th>`;

  const head = `<tr><th rowspan="2" class="bell">Звонок</th>
    ${side('left', `Левая неделя <small>(${swap ? 'чётные' : 'нечётные'} недели графика)</small>`)}
    ${side('right', `Правая неделя <small>(${swap ? 'нечётные' : 'чётные'} недели графика)</small>`)}</tr>
    <tr>${['left', 'right'].map(() => sat.map(({ d }) => `<th>${DAY_NAMES[d]}</th>`).join('')).join('')}</tr>`;

  const body = rows.map(({ r, i }) => {
    const label = r.number ? `<b>${r.number}</b> ${esc(r.time.replace('-', '–'))}` : esc(r.label || r.time);
    const pairs = r.pairs.length ? `<small>в расписаниях: пара ${r.pairs.join(' / ')}</small>` : '';
    const cells = ['left', 'right'].map(s => sat.map(({ c }) => {
      const items = grid.cells[s][i][c];
      return `<td class="${items.some(x => x.mine) ? 'has-mine' : ''}">${items.map(x =>
        `<div class="lesson ${x.mine ? 'mine' : 'other'}" title="${esc(`${x.group} — ${x.subject} (${x.teacher})`)}"><b>${esc(x.group)}</b> ${esc(x.subject)}${x.mine ? '' : ` <small>${esc(x.teacher)}</small>`}</div>`).join('')}</td>`;
    }).join('')).join('');
    return `<tr><th class="bell">${label}${pairs}</th>${cells}</tr>`;
  }).join('');

  container.innerHTML = `<div class="scroll"><table class="tpl"><thead>${head}</thead><tbody>${body}</tbody></table></div>`;
}
