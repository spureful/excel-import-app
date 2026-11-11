export const SHEET_NAME = 'Шаблон'

// Сколько строк данных подготовить в шаблоне (с запасом к ~2000)
export const MAX_ROWS = 3000

// Служебный лист со значениями выпадающих списков
export const LISTS_SHEET_NAME = '_lists'

/**
 * Колонки шаблона.
 *
 * header   — заголовок в Excel
 * key      — ключ в объекте строки после разбора
 * required — ячейку нельзя оставить пустой
 * list     — выпадающий список. Строка или { label, value }:
 *            label попадает в Excel, value — в объект
 * width    — ширина колонки в символах
 */
export const COLUMNS = [
  { header: 'ид', key: 'id', required: true, width: 16 },
  { header: 'имя', key: 'name', width: 36 },
  {
    header: 'программа',
    key: 'program',
    width: 22,
    list: [
      { label: 'новая', value: 'new' },
      { label: 'существующая', value: 'existing' },
      { label: 'модернизация', value: 'modernization' },
      { label: 'демонтаж', value: 'demolition' }
    ]
  }
]

export function listOptions(column) {
  return (column.list ?? []).map((item) =>
    item !== null && typeof item === 'object'
      ? { label: String(item.label).trim(), value: item.value }
      : { label: String(item).trim(), value: item }
  )
}

export function displayValue(column, value) {
  if (value === null || value === undefined || value === '') return ''
  const option = listOptions(column).find((item) => item.value === value)
  return option ? option.label : String(value)
}
