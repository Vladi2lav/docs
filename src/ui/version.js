// Версия в шапке и проверка обновления: если на сервере уже новая версия, а браузер работает со старыми файлами
// из кэша, показывается полоска с кнопкой, которая перезапрашивает файлы в обход кэша и перезагружает страницу.

import { APP_VERSION } from '../version.js';

export async function initVersion({ label, banner, button }) {
  label.textContent = `версия ${APP_VERSION}`;
  try {
    const res = await fetch('version.json', { cache: 'no-store' });
    if (!res.ok) return;
    const { version } = await res.json();
    if (version && version !== APP_VERSION) {
      banner.querySelector('[data-version]').textContent = version;
      banner.hidden = false;
      button.onclick = refreshFiles;
    }
  } catch { /* локальный запуск без version.json — проверять нечего */ }
}

// Перезапрашивает в обход кэша все свои файлы страницы (скрипты, стили, данные), затем перезагружает её
async function refreshFiles() {
  const own = performance.getEntriesByType('resource')
    .map(e => e.name)
    .filter(url => url.startsWith(location.origin) && /\.(js|css|json)(\?|$)/.test(url));
  await Promise.all([location.href.split('#')[0], ...own].map(url => fetch(url, { cache: 'reload' }).catch(() => null)));
  location.reload();
}
