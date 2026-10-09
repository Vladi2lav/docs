// Вкладка «Календарь рабочих дней»: учебный год по месяцам, клик по дню — нерабочий/рабочий.

import { MONTH_NAMES } from '../config.js';
import { daysInMonth, iso, ruDate, weekday } from '../util/dates.js';
import { esc } from './preview.js';

// first — {y, m} первого месяца учебного года (из графика), далее 12 месяцев подряд
export function renderCalendar(container, { first, calendar, onToggle }) {
  const months = Array.from({ length: 12 }, (_, i) => ({ y: first.y + Math.floor((first.m + i) / 12), m: (first.m + i) % 12 }));
  const names = [];

  const html = months.map(({ y, m }) => {
    const offset = (weekday(iso(y, m, 1)) + 6) % 7;      // Пн = 0
    let working = 0;
    const cells = Array(offset).fill('<i></i>');
    for (let d = 1; d <= daysInMonth(y, m); d++) {
      const date = iso(y, m, d);
      const { off, manual } = calendar.info(date);
      if (!off) working++;
      if (off && off.kind !== 'weekend') names.push({ date, text: off.name, manual: !!manual });
      const cls = !off ? (manual ? 'work-manual' : '') : off.kind === 'weekend' ? 'weekend' : off.manual ? 'custom' : 'holiday';
      cells.push(`<b class="${cls}" data-date="${date}" title="${esc(off ? off.name : 'Рабочий день')}">${d}</b>`);
    }
    return `<div class="month"><h4>${MONTH_NAMES[m]} ${y}<small>раб. дней: ${working}</small></h4>
      <div class="dow"><span>Пн</span><span>Вт</span><span>Ср</span><span>Чт</span><span>Пт</span><span>Сб</span><span>Вс</span></div>
      <div class="days">${cells.join('')}</div></div>`;
  }).join('');

  const list = names.map(n => `<li>${ruDate(n.date)} — ${esc(n.text)}${n.manual ? ' <em>(ваша правка)</em>' : ''}</li>`).join('');
  container.innerHTML = `<div class="months">${html}</div>
    <details open><summary>Нерабочие дни учебного года (${names.length})</summary><ul class="list">${list}</ul></details>`;
  container.querySelectorAll('.days b').forEach(el => { el.onclick = () => onToggle(el.dataset.date); });
}
