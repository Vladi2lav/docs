// Расширенный вид практики (всплывающее окно): шаблон недели по парам, сетка недель периода,
// правка отдельного дня, текст «пн 2-3 пара, вт 4-6 пара …» и скачивание документа практики.
// Окно работает с копией блоков; в строку ведомости они попадают только по кнопке «Применить».

import { applyPattern, describe, DOW_SHORT, expandBlock, pairsText, patternOf, setDay, totalHours } from '../domain/practice-days.js';
import { addDays, ruDate, weekday } from '../util/dates.js';
import { esc } from './preview.js';

const mondayOf = date => addDays(date, -((weekday(date) + 6) % 7));
const ddmm = date => `${date.slice(8)}.${date.slice(5, 7)}`;

// ctx: { row, graph, calendar, bells, doc, specialties, loadInfo, onSave(blocks, doc), onDownload(blocks, doc) }
//   calendar — календарь, где суббота рабочая (практика идёт и по субботам); праздники из него отмечаются в сетке
export function openPracticeDialog(dialog, ctx) {
  const { row, graph, calendar, bells } = ctx;
  const blocks = row.blocks.map(b => expandBlock(b, calendar));
  const doc = { ...ctx.doc };
  const merged = { days: Object.assign({}, ...blocks.map(b => b.days)), slots: Object.assign({}, ...blocks.map(b => b.slots)) };
  const draft = {};                                   // день недели → [пары] | [] (практики нет) | отсутствует (не менять)
  for (const [dow, pairs] of Object.entries(patternOf(merged))) if (pairs) draft[dow] = pairs;
  const maxPair = Math.max(7, bells.length, ...Object.values(merged.slots).flat());
  let selected = null;

  const pairButtons = (on, attrs) => Array.from({ length: maxPair }, (_, i) =>
    `<button type="button" class="pair ${on.includes(i + 1) ? 'on' : ''}" ${attrs} data-pair="${i + 1}" title="${esc(bells[i] ? bells[i].replace('-', '–') : `пара ${i + 1}`)}">${i + 1}</button>`).join('');
  const blockOf = date => blocks.find(b => date >= b.from && date <= b.to);

  function render() {
    const scroll = dialog.scrollTop;
    const dates = blocks.flatMap(b => { const out = []; for (let d = b.from; d <= b.to; d = addDays(d, 1)) out.push(d); return out; });
    const mondays = [...new Set(dates.map(mondayOf))].sort();
    const weekRows = mondays.map(m => {
      const days = Array.from({ length: 6 }, (_, i) => addDays(m, i));
      const total = days.reduce((s, d) => s + (blockOf(d)?.days[d] || 0), 0);
      const week = graph.weekAt(m) || graph.weekAt(addDays(m, 6));
      return `<tr><th class="wk">${week ? `нед. ${week.num}` : ''}<small>${ddmm(m)}</small></th>${days.map(d => {
        const blk = blockOf(d);
        if (!blk) return '<td class="pd out"></td>';
        const hours = blk.days[d] || 0;
        const off = calendar.dayOff(d);
        const text = hours ? (blk.slots[d] ? pairsText(blk.slots[d]).replace(' пара', '') : `${hours} ч`) : (off ? 'праздник' : '—');
        return `<td class="pd ${hours ? 'on' : ''} ${off ? 'off' : ''} ${d === selected ? 'sel' : ''}" data-date="${d}" title="${esc(off ? off.name : '')}"><i>${ddmm(d)}</i><b>${text}</b></td>`;
      }).join('')}<td class="wkt">${total || ''}</td></tr>`;
    }).join('');

    const sel = selected && blockOf(selected);
    const selOff = selected && calendar.dayOff(selected);
    const editor = sel
      ? `<div class="dayedit"><b>${ruDate(selected)}, ${DOW_SHORT[weekday(selected)]}</b>${selOff ? ` <span class="warn">— ${esc(selOff.name)}</span>` : ''}
          <span class="pairs">${pairButtons(sel.slots[selected] || [], 'data-day="1"')}</span>
          <label class="inline">или часов <input type="number" min="0" id="dayHours" value="${sel.slots[selected] ? '' : (sel.days[selected] || '')}" style="width:64px"></label>
          <button type="button" id="dayClear">Нет практики</button></div>`
      : '<p class="hint">Нажмите на день в сетке, чтобы задать пары или часы только для него.</p>';

    const desc = describe(blocks);
    const total = totalHours(blocks);
    const li = ctx.loadInfo;
    const loadLine = li
      ? `<p class="${li.total === total ? 'ok' : 'warn'}">По нагрузке: ${li.weeks && li.perWeek ? `${li.weeks} нед × ${li.perWeek} ч = ` : ''}${li.total} ч — ${li.total === total ? 'совпадает' : `сейчас ${total} ч, разница ${total - li.total} ч`}.</p>` : '';

    dialog.innerHTML = `<h2>Практика: ${esc(row.name)} <small>${esc(row.group)}</small></h2>
      <div class="pgrid">
        <label>Специальность<input id="dSpec" list="dSpecList" value="${esc(doc.specialty || '')}"></label>
        <label>Квалификация<input id="dQual" list="dSpecList" value="${esc(doc.qualification || '')}"></label>
        <label>Курс<input id="dCourse" type="number" min="1" max="6" value="${esc(doc.course || '')}"></label>
        <label class="wide2">Модуль<input id="dModule" value="${esc(doc.moduleLine || '')}"></label>
        <datalist id="dSpecList">${(ctx.specialties || []).map(s => `<option value="${esc(s)}">`).join('')}</datalist>
      </div>
      <p class="periods">${blocks.map((b, i) => `Период: <input type="date" data-i="${i}" data-f="from" value="${b.from}"> – <input type="date" data-i="${i}" data-f="to" value="${b.to}">`).join(' · ')}</p>
      <h3>Шаблон недели — какие пары</h3>
      <div class="pattern">${[1, 2, 3, 4, 5, 6].map(dow => `<div class="patrow"><b>${DOW_SHORT[dow]}</b>
        <span class="pairs">${pairButtons(draft[dow] || [], `data-dow="${dow}"`)}</span>
        <button type="button" class="none ${draft[dow] && !draft[dow].length ? 'on' : ''}" data-dow="${dow}" data-none="1">нет практики</button>
        <small>${draft[dow] && draft[dow].length ? pairsText(draft[dow]) : ''}</small></div>`).join('')}
        <button type="button" id="applyPat" class="primary">Применить ко всем неделям</button>
        <small class="hint">Дни недели без выбора не меняются. Праздники и выходные пропускаются.</small></div>
      <h3>Недели периода</h3>
      <div class="scroll"><table class="pdays"><thead><tr><th></th>${[1, 2, 3, 4, 5, 6].map(d => `<th>${DOW_SHORT[d]}</th>`).join('')}<th>Σ ч</th></tr></thead><tbody>${weekRows}</tbody></table></div>
      ${editor}
      <p class="pdesc"><b>В какие дни:</b> ${esc(desc) || '—'}<br><b>Всего:</b> ${total} ч</p>${loadLine}
      <div class="row"><button type="button" id="dDownload">Скачать документ практики (.docx)</button>
        <button type="button" id="dSave" class="primary">Применить</button><button type="button" id="dCancel">Отмена</button></div>`;
    dialog.scrollTop = scroll;
    wire();
  }

  function wire() {
    const $ = s => dialog.querySelector(s);
    const bind = (id, key) => { $(id).oninput = e => { doc[key] = e.target.value; }; };
    bind('#dSpec', 'specialty'); bind('#dQual', 'qualification'); bind('#dModule', 'moduleLine');
    $('#dCourse').oninput = e => { doc.course = e.target.value; };

    dialog.querySelectorAll('.periods input').forEach(input => {
      input.onchange = () => {
        const b = blocks[Number(input.dataset.i)];
        if (!input.value) return;
        b[input.dataset.f] = input.value;
        if (b.to < b.from) b.to = b.from;
        for (const d of Object.keys(b.days)) if (d < b.from || d > b.to) { delete b.days[d]; delete b.slots[d]; }
        render();
      };
    });

    dialog.querySelectorAll('.pattern .pair').forEach(btn => {
      btn.onclick = () => {
        const dow = btn.dataset.dow, pair = Number(btn.dataset.pair);
        const cur = draft[dow] || [];
        const next = cur.includes(pair) ? cur.filter(p => p !== pair) : [...cur, pair].sort((a, b) => a - b);
        if (next.length) draft[dow] = next; else delete draft[dow];
        render();
      };
    });
    dialog.querySelectorAll('.pattern .none').forEach(btn => {
      btn.onclick = () => { const dow = btn.dataset.dow; if (draft[dow] && !draft[dow].length) delete draft[dow]; else draft[dow] = []; render(); };
    });
    $('#applyPat').onclick = () => { for (const b of blocks) applyPattern(b, draft, calendar); render(); };

    dialog.querySelectorAll('td.pd[data-date]').forEach(td => { td.onclick = () => { selected = td.dataset.date; render(); }; });
    dialog.querySelectorAll('.dayedit .pair').forEach(btn => {
      btn.onclick = () => {
        const blk = blockOf(selected), pair = Number(btn.dataset.pair);
        const cur = blk.slots[selected] || [];
        setDay(blk, selected, { pairs: cur.includes(pair) ? cur.filter(p => p !== pair) : [...cur, pair] });
        render();
      };
    });
    const hours = $('#dayHours');
    if (hours) hours.onchange = () => { setDay(blockOf(selected), selected, { hours: hours.value }); render(); };
    const clear = $('#dayClear');
    if (clear) clear.onclick = () => { setDay(blockOf(selected), selected, {}); render(); };

    $('#dDownload').onclick = () => ctx.onDownload(blocks, doc);
    $('#dSave').onclick = () => { ctx.onSave(blocks, doc); dialog.close(); };
    $('#dCancel').onclick = () => dialog.close();
  }

  render();
  dialog.showModal();
}
