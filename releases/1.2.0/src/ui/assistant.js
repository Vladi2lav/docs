// Чат-помощник по этому сайту. Ответы даёт Microsoft Copilot через наш прокси на Vercel (адрес — src/chat-config.js),
// ключ хранится в секретах Vercel, пользователь ничего не вводит. Помощник видит краткую сводку о странице (не файлы)
// и ничего в ведомости не меняет. Что отвечать и на какие темы — задаёт прокси.

import { CHAT } from '../chat-config.js';
import { esc } from './preview.js';

// Значок Copilot — один раз, в шапке окна чата: видно, чья модель отвечает
const COPILOT_MARK = `<svg class="as-logo" viewBox="0 0 24 24" aria-hidden="true"><defs>
  <radialGradient id="cpl-rg1" cx="0" cy="0" r="1" gradientUnits="userSpaceOnUse" gradientTransform="matrix(-9.77 19.84 38.94 15.77 12.67 0.59)">
    <stop offset=".25" stop-color="#109df2"/><stop offset=".72" stop-color="#aad445"/><stop offset=".97" stop-color="#ffd400"/>
  </radialGradient>
  <radialGradient id="cpl-rg2" cx="0" cy="0" r="1" gradientUnits="userSpaceOnUse" gradientTransform="matrix(-9.94 20.80 24.65 12.53 21.95 4.76)">
    <stop offset=".07" stop-color="#8c48ff"/><stop offset=".5" stop-color="#f2598a"/><stop offset=".9" stop-color="#ffb152"/>
  </radialGradient>
  <linearGradient id="cpl-lg1" gradientUnits="userSpaceOnUse" x1="8.275" y1="0.534" x2="8.177" y2="17.776">
    <stop stop-color="#3dcbff"/><stop offset=".25" stop-color="#0588f7" stop-opacity="0"/>
  </linearGradient>
  <linearGradient id="cpl-lg2" gradientUnits="userSpaceOnUse" x1="15.839" y1="5.211" x2="15.829" y2="9.907">
    <stop offset=".06" stop-color="#f8adfa"/><stop offset=".71" stop-color="#a86edd" stop-opacity="0"/>
  </linearGradient></defs>
  <path d="M16.295 5.29l-0.242 0.986h-2.32c-1.721 0-3.22 1.197-3.638 2.902l-2.101 8.593H3.798c-2.44 0-4.229-2.342-3.637-4.761l2.106-8.612c0.557-2.274 2.556-3.87 4.85-3.87h5.54c2.44 0 4.23 2.343 3.638 4.761" fill="url(#cpl-rg1)"/>
  <path d="M16.295 5.29l-0.242 0.986h-2.32c-1.721 0-3.22 1.197-3.638 2.902l-2.101 8.593H3.798c-2.44 0-4.229-2.342-3.637-4.761l2.106-8.612c0.557-2.274 2.556-3.87 4.85-3.87h5.54c2.44 0 4.23 2.343 3.638 4.761" fill="url(#cpl-lg1)"/>
  <path d="M7.752 18.758l0.242-0.986h2.321c1.72 0 3.22-1.197 3.637-2.903l2.102-8.593h4.196c2.439 0 4.228 2.343 3.637 4.762l-2.107 8.611c-0.556 2.274-2.555 3.87-4.85 3.87h-5.54c-2.44 0-4.229-2.342-3.638-4.761" fill="url(#cpl-rg2)"/>
  <path d="M7.752 18.758l0.242-0.986h2.321c1.72 0 3.22-1.197 3.637-2.903l2.102-8.593h4.196c2.439 0 4.228 2.343 3.637 4.762l-2.107 8.611c-0.556 2.274-2.555 3.87-4.85 3.87h-5.54c-2.44 0-4.229-2.342-3.638-4.761" fill="url(#cpl-lg2)"/></svg>`;

export function initAssistant({ button, panel, getContext }) {
  const history = [];
  panel.innerHTML = `<div class="as-head">${COPILOT_MARK}<b>Copilot</b> <small>на базе Microsoft Copilot · только по этому сайту</small>
      <button type="button" data-a="close" title="Закрыть">✕</button></div>
    <div class="as-log"><div class="as-msg bot">${CHAT.proxy
      ? 'Здравствуйте! Спрашивайте о работе с ведомостью: «почему в октябре нет пар?», «что ещё не заполнено?», «как добавить практику?».'
      : 'Помощник пока не подключён (не задан адрес прокси в src/chat-config.js).'}</div></div>
    <form class="as-form"><input placeholder="Сообщение…" autocomplete="off"${CHAT.proxy ? '' : ' disabled'}><button class="primary"${CHAT.proxy ? '' : ' disabled'}>➤</button></form>`;
  const log = panel.querySelector('.as-log');
  const input = panel.querySelector('.as-form input');

  const add = (role, text) => {
    log.insertAdjacentHTML('beforeend', `<div class="as-msg ${role}">${esc(text).replace(/\n/g, '<br>')}</div>`);
    log.scrollTop = log.scrollHeight;
    return log.lastElementChild;
  };
  button.onclick = () => { panel.hidden = !panel.hidden; if (!panel.hidden) input.focus(); };
  panel.querySelector('[data-a="close"]').onclick = () => { panel.hidden = true; };

  panel.querySelector('.as-form').onsubmit = async e => {
    e.preventDefault();
    const text = input.value.trim();
    if (!text || !CHAT.proxy) return;
    input.value = '';
    add('user', text);
    history.push({ role: 'user', text });
    const bubble = add('bot', '…');
    try {
      const res = await fetch(CHAT.proxy, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ messages: history.slice(-12), context: getContext() }),
      });
      if (!res.ok) throw new Error(res.status === 429 ? 'Лимит бесплатных запросов исчерпан. Попробуйте позже.' : `Помощник недоступен (ошибка ${res.status}).`);
      const answer = (await res.json()).text || 'Не удалось получить ответ, переформулируйте вопрос.';
      history.push({ role: 'model', text: answer });
      bubble.innerHTML = esc(answer).replace(/\n/g, '<br>');
    } catch (err) {
      history.pop();
      bubble.className = 'as-msg err';
      bubble.textContent = err instanceof TypeError ? 'Нет связи с помощником. Проверьте интернет.' : err.message;
    }
    log.scrollTop = log.scrollHeight;
  };
}
