// Поле выбора из списка с несколькими значениями — как в 1С: набираете, список фильтруется, выбранное
// превращается в «плашку»; можно вписать своё значение. Результат — значения через запятую.

import { esc } from './preview.js';

// host — пустой <div>; options — варианты; allowCustom — разрешён ли свой текст; onChange(items)
export function makeCombo(host, { options = [], placeholder = '', allowCustom = true, separator = ', ', onChange = () => {} } = {}) {
  let items = [];
  let opts = options;
  let active = 0;

  host.classList.add('combo');
  host.innerHTML = `<div class="combo-box"><span class="combo-chips"></span>
    <input class="combo-input" placeholder="${esc(placeholder)}" autocomplete="off">
    <button type="button" class="combo-toggle" tabindex="-1" title="Показать список">▾</button></div>
    <ul class="combo-list" hidden></ul>`;
  const chipsEl = host.querySelector('.combo-chips');
  const input = host.querySelector('.combo-input');
  const list = host.querySelector('.combo-list');
  const toggle = host.querySelector('.combo-toggle');

  const visible = () => {
    const q = input.value.trim().toLowerCase();
    const found = opts.filter(o => !items.includes(o) && (!q || o.toLowerCase().includes(q))).map(value => ({ value }));
    const typed = input.value.trim();
    const custom = allowCustom && typed && !items.includes(typed) && !opts.some(o => o.toLowerCase() === q) ? [{ value: typed, custom: true }] : [];
    return [...found, ...custom];
  };

  function renderChips() {
    chipsEl.innerHTML = items.map((t, i) =>
      `<span class="combo-chip" title="${esc(t)}">${esc(t)}<button type="button" data-i="${i}" tabindex="-1" title="Убрать">✕</button></span>`).join('');
    chipsEl.querySelectorAll('button').forEach(b => { b.onclick = () => remove(Number(b.dataset.i)); });
  }

  function renderList() {
    const rows = visible();
    active = Math.min(active, Math.max(0, rows.length - 1));
    list.innerHTML = rows.length
      ? rows.map((r, i) => `<li data-i="${i}" class="${i === active ? 'on' : ''}">${r.custom ? `Добавить «${esc(r.value)}»` : esc(r.value)}</li>`).join('')
      : '<li class="none">Нет подходящих вариантов</li>';
    list.querySelectorAll('li[data-i]').forEach(li => {
      li.onmousedown = e => { e.preventDefault(); add(rows[Number(li.dataset.i)].value); };
    });
  }

  const open = () => { list.hidden = false; renderList(); };
  const close = () => { list.hidden = true; };
  const change = () => onChange(items.slice());

  function add(text) {
    const t = String(text).trim();
    if (!t || items.includes(t)) return;
    items.push(t);
    input.value = '';
    renderChips(); renderList(); change();
  }

  function remove(i) {
    items.splice(i, 1);
    renderChips(); renderList(); change();
  }

  input.onfocus = open;
  input.oninput = () => { active = 0; open(); };
  input.onkeydown = e => {
    const rows = visible();
    if (e.key === 'ArrowDown') { e.preventDefault(); active = Math.min(active + 1, rows.length - 1); open(); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); active = Math.max(active - 1, 0); open(); }
    else if (e.key === 'Enter') { e.preventDefault(); const r = rows[active] || rows[0]; if (r) add(r.value); }
    else if (e.key === 'Escape') close();
    else if (e.key === 'Backspace' && !input.value && items.length) remove(items.length - 1);
  };
  toggle.onclick = () => { if (list.hidden) { input.focus(); open(); } else close(); };
  document.addEventListener('mousedown', e => { if (!host.contains(e.target)) close(); });
  host.querySelector('.combo-box').onclick = e => { if (e.target.closest('.combo-box') && !e.target.closest('button')) input.focus(); };
  renderChips();

  return {
    get items() { return items.slice(); },
    set items(v) { items = [...v]; renderChips(); renderList(); },
    get text() { return items.join(separator); },
    // Строка → плашки: подряд идущие части, образующие известный вариант целиком, склеиваются (в вариантах бывают запятые)
    setText(text) {
      const parts = String(text || '').split(/,\s*/).map(s => s.trim()).filter(Boolean);
      const out = [];
      for (let i = 0; i < parts.length;) {
        let j = parts.length;
        while (j > i + 1 && !opts.includes(parts.slice(i, j).join(', '))) j--;
        out.push(parts.slice(i, j).join(', '));
        i = j;
      }
      items = out;
      renderChips(); renderList();
    },
    setOptions(o) { opts = o; renderList(); },
    setNeed(on) { host.classList.toggle('need', !!on); },
  };
}
