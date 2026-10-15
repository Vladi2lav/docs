// Обновляет номер версии перед коммитом: node tools/bump-version.mjs
// Пишет одно и то же значение в src/version.js (код, который работает в браузере) и version.json (что лежит на сервере).
// Если они расходятся, страница показывает «Сайт обновлён» и предлагает подтянуть свежие файлы.

import { writeFileSync } from 'node:fs';

const now = new Date();
const p2 = n => String(n).padStart(2, '0');
const version = `${now.getFullYear()}-${p2(now.getMonth() + 1)}-${p2(now.getDate())} ${p2(now.getHours())}:${p2(now.getMinutes())}`;

writeFileSync(new URL('../src/version.js', import.meta.url),
  `// Генерируется tools/bump-version.mjs — не редактировать вручную\nexport const APP_VERSION = '${version}';\n`);
writeFileSync(new URL('../version.json', import.meta.url), `${JSON.stringify({ version })}\n`);
console.log('version', version);
