// Чат-помощник на GitHub Models (OpenAI-совместимый API). Токен GitHub (fine-grained, право «Models: read»)
// вводит сам пользователь; он хранится только в этом браузере (localStorage) и уходит только на адрес API.
// Адрес можно заменить (например, на свой прокси). Помощник видит краткую сводку о состоянии страницы,
// но не загруженные файлы целиком. Он только объясняет — ничего в ведомости сам не меняет.

import { esc } from './preview.js';

const KEY = 'assistant-settings';
const DEFAULTS = { token: '', endpoint: 'https://models.github.ai/inference/chat/completions', model: 'openai/gpt-4.1-mini' };

const SYSTEM = `Ты помощник в веб-приложении «Учёт времени педагога»: оно строит ведомость учёта учебного времени (Форма Приказ №130, Казахстан) из графика учебного процесса, расписаний занятий (по курсам), нагрузки и документов практики.
Отвечай по-русски, коротко и по делу. Объясняй, почему в ведомости получились такие часы или чего не хватает, и что нажать. Опирайся на «Состояние страницы» ниже; если данных нет — так и скажи, ничего не выдумывай (ни часов, ни фамилий, ни номеров модулей).
Правила приложения: пустая неделя группы в графике — теория, любой код — практика/каникулы; час-пара = 2 часа; праздники берутся из вкладки «Календарь»; план часов вводится вручную или берётся из нагрузки только при однозначном совпадении; практика задаётся периодами «ч/день» или «Расширенным видом» по парам.`;

const load = () => { try { return { ...DEFAULTS, ...JSON.parse(localStorage.getItem(KEY) || '{}') }; } catch { return { ...DEFAULTS }; } };
const save = s => { try { localStorage.setItem(KEY, JSON.stringify(s)); } catch { /* приватный режим: настройки живут до закрытия вкладки */ } };

// getContext() → строка со сводкой состояния страницы (без файлов)
export function initAssistant({ button, panel, getContext }) {
  let settings = load();
  const history = [];

  panel.innerHTML = `<div class="as-head"><b>Помощник</b> <small>GitHub Models</small>
      <button type="button" data-a="cfg" title="Настройки">⚙</button><button type="button" data-a="close" title="Закрыть">✕</button></div>
    <div class="as-cfg" hidden>
      <label>Токен GitHub (право Models: read)<input type="password" data-f="token" placeholder="github_pat_…" autocomplete="off"></label>
      <label>Модель<input data-f="model"></label>
      <label>Адрес API<input data-f="endpoint"></label>
      <small>Токен хранится только в этом браузере и отправляется только на указанный адрес. Бесплатный лимит GitHub Models небольшой.</small>
    </div>
    <div class="as-log"><p class="hint">Спросите, например: «почему в октябре нет пар?», «что ещё не заполнено?».</p></div>
    <form class="as-form"><input placeholder="Сообщение…" autocomplete="off"><button class="primary">➤</button></form>`;
  const $ = s => panel.querySelector(s);
  const log = $('.as-log'), cfg = $('.as-cfg'), input = $('.as-form input');
  panel.querySelectorAll('[data-f]').forEach(el => {
    el.value = settings[el.dataset.f];
    el.oninput = () => { settings = { ...settings, [el.dataset.f]: el.value.trim() }; save(settings); };
  });

  const add = (role, text) => {
    log.insertAdjacentHTML('beforeend', `<div class="as-msg ${role}">${esc(text).replace(/\n/g, '<br>')}</div>`);
    log.scrollTop = log.scrollHeight;
    return log.lastElementChild;
  };

  button.onclick = () => { panel.hidden = !panel.hidden; if (!panel.hidden) { if (!settings.token) cfg.hidden = false; input.focus(); } };
  $('[data-a="close"]').onclick = () => { panel.hidden = true; };
  $('[data-a="cfg"]').onclick = () => { cfg.hidden = !cfg.hidden; };

  $('.as-form').onsubmit = async e => {
    e.preventDefault();
    const text = input.value.trim();
    if (!text) return;
    if (!settings.token) { cfg.hidden = false; add('err', 'Введите токен GitHub в настройках (⚙).'); return; }
    input.value = '';
    add('user', text);
    history.push({ role: 'user', content: text });
    const bubble = add('bot', '…');
    try {
      const res = await fetch(settings.endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${settings.token}`, Accept: 'application/json' },
        body: JSON.stringify({
          model: settings.model,
          messages: [{ role: 'system', content: `${SYSTEM}\n\nСостояние страницы:\n${getContext()}` }, ...history.slice(-12)],
          temperature: 0.2, max_tokens: 700,
        }),
      });
      if (!res.ok) {
        const detail = (await res.text()).slice(0, 200);
        throw new Error(res.status === 401 ? 'Токен не принят (401). Проверьте токен и право «Models: read».'
          : res.status === 429 ? 'Лимит бесплатных запросов GitHub Models исчерпан (429). Попробуйте позже.' : `Ошибка ${res.status}: ${detail}`);
      }
      const answer = (await res.json()).choices?.[0]?.message?.content?.trim() || '(пустой ответ)';
      history.push({ role: 'assistant', content: answer });
      bubble.innerHTML = esc(answer).replace(/\n/g, '<br>');
    } catch (err) {
      history.pop();
      bubble.className = 'as-msg err';
      bubble.textContent = err instanceof TypeError ? 'Не удалось связаться с API (сеть или запрет CORS). Укажите адрес своего прокси в настройках (⚙).' : err.message;
    }
    log.scrollTop = log.scrollHeight;
  };
}
