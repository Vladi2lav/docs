// «Заполнить по документам»: сбрасывает поля строки к значениям из документов
// и перечисляет, чего в документах нет и что нужно ввести вручную. Ничего не угадывается.

import { normGroup } from '../parsers/graph.js';

const clone = blocks => blocks.map(b => ({ ...b }));

export function fillRow(row, { graph }) {
  const filled = [];
  const missing = [];
  Object.assign(row, { moduleIndex: row.defaults.moduleIndex, name: row.defaults.name, group: row.defaults.group, plan: row.defaults.plan });
  if (row.kind === 'practice') row.blocks = clone(row.defaults.blocks);
  row.moduleRef = '';

  if (!graph.groups[normGroup(row.group)]) {
    missing.push(`группа ${row.group} не найдена в графике (файл графика, шаг 1): теория, каникулы и практика для неё не определены`);
  }

  if (row.kind === 'theory') {
    filled.push('группа, дисциплина и пары по дням — из расписания');
    missing.push('индекс и полное название модуля: в расписании не указаны — выберите модуль в списке «Модуль из графика» над строкой');
  } else {
    const fromDoc = row.blocks.filter(b => b.fromDoc).length;
    if (fromDoc) filled.push(`модуль ${row.moduleIndex} и часы по дням — из docx практики`);
    const noHours = row.blocks.filter(b => !(Number(b.hoursPerDay) > 0) && !b.days).length;
    if (!row.blocks.length) missing.push(`периоды практики: в графике у группы ${row.group} нет недель этого модуля — добавьте период в блоке «Практика»`);
    if (noHours) missing.push('часы практики в день: в графике их нет — укажите в блоке «Практика» или загрузите docx практики');
  }

  if (row.plan === null) {
    missing.push('план часов: нет ни в одном документе — введите в поле «План, ч» (пока колонки «Запланировано» и «Остаток» остаются пустыми)');
  }
  return { filled, missing };
}
