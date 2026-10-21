// Обёртка над листом SheetJS: плоская матрица текста + размеры объединённых ячеек.
// Значение объединённой области лежит только в её левой верхней ячейке — как и в самом Excel.

function wrapSheet(XLSX, ws) {
  const cells = new Map();
  let rowCount = 0;
  let colCount = 0;
  for (const addr of Object.keys(ws)) {
    if (addr[0] === '!') continue;
    const { r, c } = XLSX.utils.decode_cell(addr);
    const cell = ws[addr];
    const raw = cell.t === 'n' ? cell.v : (cell.w ?? cell.v);
    const text = String(raw ?? '').replace(/\s+/g, ' ').trim();
    if (!text) continue;
    cells.set(`${r},${c}`, text);
    rowCount = Math.max(rowCount, r + 1);
    colCount = Math.max(colCount, c + 1);
  }

  const spans = new Map();
  for (const m of ws['!merges'] || []) {
    spans.set(`${m.s.r},${m.s.c}`, { rows: m.e.r - m.s.r + 1, cols: m.e.c - m.s.c + 1 });
  }

  return {
    rowCount,
    colCount,
    get: (r, c) => cells.get(`${r},${c}`) || '',
    span: (r, c) => spans.get(`${r},${c}`) || { rows: 1, cols: 1 },
    // цвет заливки ячейки (hex) или '' — нужен графику: цветом там отмечены грантовые группы
    fill: (r, c) => {
      const s = ws[XLSX.utils.encode_cell({ r, c })]?.s;
      return s && s.patternType && s.patternType !== 'none' ? (s.fgColor?.rgb || `theme${s.fgColor?.theme ?? ''}`) : '';
    },
  };
}

// Первый лист книги
export function readSheet(XLSX, arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) throw new Error('В файле нет листов');
  return wrapSheet(XLSX, ws);
}

// Все листы книги: [{ name, sheet }] — документы колледжа иногда лежат в одной книге (график, расписания, нагрузка)
export function readSheets(XLSX, arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true });
  if (!wb.SheetNames.length) throw new Error('В файле нет листов');
  return wb.SheetNames.filter(name => wb.Sheets[name]['!ref']).map(name => ({ name, sheet: wrapSheet(XLSX, wb.Sheets[name]) }));
}
