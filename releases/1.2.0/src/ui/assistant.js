// Чат-помощник по этому сайту. Ответы даёт Google Gemini через наш прокси на Vercel (адрес — src/chat-config.js),
// ключ хранится в секретах Vercel, пользователь ничего не вводит. Помощник видит краткую сводку о странице (не файлы)
// и ничего в ведомости не меняет. Что отвечать и на какие темы — задаёт прокси.

import { CHAT } from '../chat-config.js';
import { esc } from './preview.js';

// Значок Gemini — один раз, в шапке окна чата: видно, чья модель отвечает
const GEMINI_MARK = `<svg class="as-logo" viewBox="0 0 24 24" aria-hidden="true"><defs><linearGradient id="gmg" x1="0" y1="0" x2="1" y2="1">
  <stop offset="0" stop-color="#4285f4"/><stop offset=".5" stop-color="#9b72cb"/><stop offset="1" stop-color="#d96570"/></linearGradient></defs>
  <path fill="url(#gmg)" d="M12 0C12 6.6 17.4 12 24 12C17.4 12 12 17.4 12 24C12 17.4 6.6 12 0 12C6.6 12 12 6.6 12 0Z"/></svg>`;

export function initAssistant({ button, panel, getContext }) {
  const history = [];
  panel.innerHTML = `<div class="as-head">${GEMINI_MARK}<b>Помощник</b> <small>на базе Google Gemini · только по этому сайту</small>
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
