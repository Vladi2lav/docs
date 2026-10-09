// Даты хранятся строками ISO «YYYY-MM-DD»; месяц везде 0-based (как в Date).

const p2 = n => String(n).padStart(2, '0');

export const iso = (y, m, d) => `${y}-${p2(m + 1)}-${p2(d)}`;
export const toDate = s => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
export const fromDate = dt => iso(dt.getFullYear(), dt.getMonth(), dt.getDate());
export const daysInMonth = (y, m) => new Date(y, m + 1, 0).getDate();
export const weekday = s => toDate(s).getDay();                // 0 = Вс … 6 = Сб
export const addDays = (s, n) => { const dt = toDate(s); dt.setDate(dt.getDate() + n); return fromDate(dt); };
export const ruDate = s => s.split('-').reverse().join('.');
export const ymOf = s => ({ y: Number(s.slice(0, 4)), m: Number(s.slice(5, 7)) - 1 });
export const ymIndex = ({ y, m }) => y * 12 + m;

export function monthsBetween(from, to) {
  const out = [];
  for (let i = ymIndex(from); i <= ymIndex(to); i++) out.push({ y: Math.floor(i / 12), m: i % 12 });
  return out;
}

// «23.12.2024» → «2024-12-23»
export function parseRuDate(s) {
  const m = /(\d{1,2})\.(\d{1,2})\.(\d{4})/.exec(s);
  return m ? iso(Number(m[3]), Number(m[2]) - 1, Number(m[1])) : null;
}
