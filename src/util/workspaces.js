// Рабочие области: сохранение хода работы (загруженные файлы + все правки) в браузере (IndexedDB)
// и перенос между устройствами через файл .json.
//
// Запись: { id, name, updatedAt, data } — data целиком определяется приложением (см. snapshot() в main.js).
// Двоичные файлы лежат в data как ArrayBuffer; в .json они кодируются в base64.

const DB = 'timesheet-workspaces';
const STORE = 'workspaces';

function openDb() {
  return new Promise((resolve, reject) => {
    if (!('indexedDB' in globalThis)) return reject(new Error('Хранилище браузера недоступно — используйте выгрузку в файл'));
    const req = indexedDB.open(DB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE, { keyPath: 'id' });
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function run(mode, action) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, mode);
    const request = action(tx.objectStore(STORE));
    tx.oncomplete = () => { db.close(); resolve(request.result); };
    tx.onerror = tx.onabort = () => { db.close(); reject(tx.error); };
  });
}

export const getWorkspace = id => run('readonly', s => s.get(id));
export const putWorkspace = ws => run('readwrite', s => s.put(ws));
export const deleteWorkspace = id => run('readwrite', s => s.delete(id));
export async function listWorkspaces() {
  const all = await run('readonly', s => s.getAll());
  return all.map(({ id, name, updatedAt }) => ({ id, name, updatedAt })).sort((a, b) => b.updatedAt - a.updatedAt);
}

// ───────── файл .json (для переноса на другое устройство) ─────────

function toBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
}

function fromBase64(text) {
  const binary = atob(text);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

export const workspaceToJson = ws => JSON.stringify(ws, (_, v) => (v instanceof ArrayBuffer ? { __base64: toBase64(v) } : v));

export function workspaceFromJson(text) {
  const ws = JSON.parse(text, (_, v) => (v && typeof v === 'object' && '__base64' in v ? fromBase64(v.__base64) : v));
  if (!ws || !ws.data || !ws.name) throw new Error('Это не файл рабочей области');
  return ws;
}
