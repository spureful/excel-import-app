export const SHEET_NAME = 'Шаблон'

// Сколько строк данных подготовить в шаблоне (с запасом к ~2000)
export const MAX_ROWS = 3000

export const HEADERS = ['ид', 'имя', 'программа']

export const PROGRAMS = [
  { label: 'новая', code: 'new' },
  { label: 'существующая', code: 'existing' },
  { label: 'модернизация', code: 'modernization' },
  { label: 'демонтаж', code: 'demolition' }
]

export const PROGRAM_BY_LABEL = Object.fromEntries(
  PROGRAMS.map((p) => [p.label, p.code])
)

export const PROGRAM_LABEL_BY_CODE = Object.fromEntries(
  PROGRAMS.map((p) => [p.code, p.label])
)
