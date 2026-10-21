// Обновляет номер сборки версии сайта перед коммитом:  node tools/bump-version.mjs 1.1.0
// Пишет одно и то же значение в releases/<версия>/src/version.js (код в браузере) и releases/<версия>/version.json
// (что лежит на сервере). Если они расходятся, страница показывает «Сайт обновлён» и предлагает подтянуть файлы.

import { existsSync, writeFileSync } from 'node:fs';

const release = process.argv[2];
const dir = new URL(`../releases/${release}/`, import.meta.url);
if (!release || !existsSync(dir)) {
  console.error('Укажите существующую версию: node tools/bump-version.mjs 1.1.0');
  process.exit(1);
}

const now = new Date();
const p2 = n => String(n).padStart(2, '0');
const version = `${now.getFullYear()}-${p2(now.getMonth() + 1)}-${p2(now.getDate())} ${p2(now.getHours())}:${p2(now.getMinutes())}`;

writeFileSync(new URL('src/version.js', dir),
  `// Генерируется tools/bump-version.mjs — не редактировать вручную\nexport const APP_VERSION = '${version}';\n`);
writeFileSync(new URL('version.json', dir), `${JSON.stringify({ version })}\n`);
console.log(release, 'сборка', version);
