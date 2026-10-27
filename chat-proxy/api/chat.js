// Прокси чат-помощника «Учёт времени педагога» → Google Gemini (Vercel Serverless Function).
// Ключ берётся из переменной окружения GEMINI_API_KEY (Vercel → Settings → Environment Variables) и в браузер не попадает.
// Системная инструкция задаётся здесь, а не на сайте: прокси отвечает только про приложение и не годится как общий доступ к Gemini.
//
// Переменные окружения:
//   GEMINI_API_KEY   — обязательно
//   ALLOWED_ORIGINS  — адреса сайта через запятую (например https://docs-xxxx.onrender.com); пусто — разрешены все
//   GEMINI_MODEL     — по умолчанию gemini-2.5-flash

const SYSTEM = `Ты помощник веб-приложения «Учёт времени педагога»: оно строит ведомость учёта учебного времени (Форма Приказ №130, Казахстан) из графика учебного процесса, расписаний занятий (по курсам), нагрузки и документов практики.
Отвечай ТОЛЬКО на вопросы об этом приложении и работе с ним: загрузка документов, педагог и семестр, практика, переносы и отработки, календарь, грант/коммерция, рабочие области, годовой учёт. На любые другие темы вежливо откажись и верни разговор к приложению.
Отвечай по-русски, коротко. Объясняй, почему в ведомости такие часы или чего не хватает, и что нажать. Опирайся на «Состояние страницы»; если данных нет — скажи об этом, ничего не выдумывай (часы, фамилии, номера модулей).
Правила приложения: пустая неделя группы в графике — теория, любой код — практика/каникулы; пара = 2 часа; праздники берутся из вкладки «Календарь рабочих дней»; план часов вводится вручную или берётся из нагрузки только при однозначном совпадении; практика задаётся периодами «ч/день» или «Расширенным видом» по парам.

Как устроен интерфейс (на него и ссылайся):
- Шапка: выбор версии сайта; переключатель «Грант / Коммерция» (учёт ведётся отдельно, грантовые группы — цветные в графике); «📁 Рабочие области» (сохранить/открыть, «В файл» для переноса на другой компьютер); «Сохранить» (Ctrl+S).
- Вкладка «Ведомость»:
  1. «Исходные документы»: «График учебного процесса» (.xlsx); «Расписания занятий» — кнопка «+ Добавить расписание», можно несколько файлов (по одному на курс), лишнее убрать ✕; «Нагрузка» (.xlsx, необязательно); «Расписание практики» (.docx, необязательно, руководитель привязывается к педагогу сам). Поле «Учебный год начинается в» — если год не определился по имени файла графика.
  2. «Педагог, семестр, реквизиты»: выбрать педагога и семестр; «Заполнить всё по документам»; ФИО полностью, зам. по учебной работе, «Специальность» (выбор из списка, можно несколько или своё); кнопка «⇄ Недели 1, 3, 5…» меняет левую/правую неделю местами. Во 2 семестре нужна ведомость 1 семестра (Word).
  3. «Дисциплины, практика, переносы»: таблица строк (индекс модуля, название, группа, «План, ч», «Из нагрузки»); карточки практики («+ период», «Периоды из графика», «Расширенный вид», «Документ практики (.docx)»); «Переносы и отработки» — дисциплина, «занятие за», «проведено», часы, «Сохранить»; список занятий, выпавших на праздники, с кнопкой «Назначить отработку».
  4. «Результат»: вкладки месяцев с предпросмотром и «Скачать Word (.docx)». Во 2 семестре — блок 5 «Годовой учёт часов».
- Вкладка «Расписание и практики»: шаблон недели (левая/правая) по звонкам, пары педагога подсвечены, выбор групп; «Занятия педагога в неделю»; «Звонки» (нумерация пар практики); практики педагога и подсказки «+ группа · модуль» из графика/нагрузки; подписанты документа практики.
- Вкладка «График учебного процесса»: недели и коды групп, метка «грант/комм» у группы (нажать — исправить).
- Вкладка «Календарь рабочих дней»: клик по дню — сделать нерабочим/рабочим; «Перенос дня»; «Суббота — учебный день»; выгрузить/загрузить календарь.
Красная рамка и звёздочка — обязательное поле, которое не заполнилось из документов.`;

const MAX_MESSAGES = 12;
const MAX_CHARS = 12000;

export default async function handler(req, res) {
  const origin = req.headers.origin || '';
  const allowed = (process.env.ALLOWED_ORIGINS || '').split(',').map(s => s.trim()).filter(Boolean);
  if (allowed.length && !allowed.includes(origin)) return res.status(403).json({ error: 'Источник запроса не разрешён' });
  res.setHeader('Access-Control-Allow-Origin', origin || '*');
  res.setHeader('Vary', 'Origin');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(204).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Только POST' });
  if (!process.env.GEMINI_API_KEY) return res.status(500).json({ error: 'На сервере не задан GEMINI_API_KEY' });

  // { messages: [{ role: 'user' | 'model', text }], context: 'сводка страницы' }
  const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
  const messages = (Array.isArray(body.messages) ? body.messages : []).slice(-MAX_MESSAGES)
    .filter(m => (m.role === 'user' || m.role === 'model') && typeof m.text === 'string' && m.text.trim())
    .map(m => ({ role: m.role, parts: [{ text: m.text.slice(0, 2000) }] }));
  const context = String(body.context || '').slice(0, 4000);
  if (!messages.length || messages[messages.length - 1].role !== 'user') return res.status(400).json({ error: 'Нет вопроса' });
  if (JSON.stringify(messages).length > MAX_CHARS) return res.status(413).json({ error: 'Слишком длинная переписка' });

  // У каждой модели свой лимит запросов в минуту: упёрлись в лимит (429), ошибка или долго — пробуем следующую
  const models = (process.env.GEMINI_MODELS || 'gemini-2.5-flash,gemini-2.5-flash-lite,gemini-flash-lite-latest').split(',').map(s => s.trim()).filter(Boolean);
  const payload = JSON.stringify({
    systemInstruction: { parts: [{ text: `${SYSTEM}\n\nСостояние страницы:\n${context}` }] },
    contents: messages,
    generationConfig: { temperature: 0.2, maxOutputTokens: 800 },
  });
  let lastStatus = 502;
  for (const model of models) {
    try {
      const upstream = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`, {
        method: 'POST', signal: AbortSignal.timeout(20000),
        headers: { 'Content-Type': 'application/json', 'x-goog-api-key': process.env.GEMINI_API_KEY },
        body: payload,
      });
      if (!upstream.ok) { lastStatus = upstream.status; continue; }
      const data = await upstream.json();
      const text = (data.candidates?.[0]?.content?.parts || []).map(p => p.text || '').join('').trim();
      if (text) return res.status(200).json({ text, model });
    } catch { lastStatus = 504; }
  }
  return res.status(lastStatus === 429 ? 429 : 502).json({ error: `Gemini: ${lastStatus}` });
}
