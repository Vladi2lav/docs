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
Правила приложения: пустая неделя группы в графике — теория, любой код — практика/каникулы; пара = 2 часа; праздники берутся из вкладки «Календарь рабочих дней»; план часов вводится вручную или берётся из нагрузки только при однозначном совпадении; практика задаётся периодами «ч/день» или «Расширенным видом» по парам.`;

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

  const model = process.env.GEMINI_MODEL || 'gemini-2.5-flash';
  const upstream = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-goog-api-key': process.env.GEMINI_API_KEY },
    body: JSON.stringify({
      systemInstruction: { parts: [{ text: `${SYSTEM}\n\nСостояние страницы:\n${context}` }] },
      contents: messages,
      generationConfig: { temperature: 0.2, maxOutputTokens: 800 },
    }),
  });
  if (!upstream.ok) return res.status(upstream.status === 429 ? 429 : 502).json({ error: `Gemini: ${upstream.status}` });
  const data = await upstream.json();
  const text = (data.candidates?.[0]?.content?.parts || []).map(p => p.text || '').join('').trim();
  return res.status(200).json({ text });
}
