// «Заполнить по документам»: возвращает строку к значениям из документов (то, что приложение получило из файлов).
// Чего в документах нет — остаётся пустым, интерфейс подсвечивает такие поля красным.

export function resetRow(row) {
  Object.assign(row, { moduleIndex: row.defaults.moduleIndex, name: row.defaults.name, group: row.defaults.group, plan: row.defaults.plan });
  if (row.kind === 'practice') row.blocks = row.defaults.blocks.map(b => ({ ...b }));
}
