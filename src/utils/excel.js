import {
  HEADERS,
  MAX_ROWS,
  PROGRAMS,
  PROGRAM_BY_LABEL,
  SHEET_NAME
} from '../constants'

const XLSX_MIME =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

// ExcelJS тяжёлый (~1 МБ), поэтому грузим его лениво, только при первом использовании
async function loadExcelJS() {
  const mod = await import('exceljs')
  return mod.default ?? mod
}

/* ---------------------------- Шаблон ---------------------------- */

export async function buildTemplateBlob() {
  const ExcelJS = await loadExcelJS()
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet(SHEET_NAME, {
    views: [{ state: 'frozen', ySplit: 1 }]
  })

  ws.columns = [
    { header: HEADERS[0], key: 'id', width: 16 },
    { header: HEADERS[1], key: 'name', width: 36 },
    { header: HEADERS[2], key: 'program', width: 22 }
  ]

  // Заголовки: «ид» — красный (обязательное), остальные — серые
  const headerRow = ws.getRow(1)
  headerRow.height = 22
  headerRow.eachCell((cell, col) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    cell.alignment = { vertical: 'middle', horizontal: 'left' }
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: col === 1 ? 'FFD32F2F' : 'FF607D8B' }
    }
  })
  ws.getCell('A1').note = 'Обязательное поле'
  ws.getCell('B1').note = 'Можно заполнить, только если заполнен ид'

  const programsFormula = `"${PROGRAMS.map((p) => p.label).join(',')}"`
  const lastRow = MAX_ROWS + 1

  for (let r = 2; r <= lastRow; r++) {
    // имя — вводится только при заполненном ид
    ws.getCell(`B${r}`).dataValidation = {
      type: 'custom',
      allowBlank: true,
      formulae: [`LEN(TRIM($A${r}))>0`],
      showErrorMessage: true,
      errorStyle: 'stop',
      errorTitle: 'Сначала заполните ид',
      error: 'Имя можно ввести только после заполнения поля «ид».'
    }
    // программа — выпадающий список
    ws.getCell(`C${r}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: [programsFormula],
      showErrorMessage: true,
      errorStyle: 'stop',
      errorTitle: 'Недопустимое значение',
      error: `Выберите значение из списка: ${PROGRAMS.map((p) => p.label).join(', ')}.`
    }
  }

  // Подсветка пустого ид, если в строке уже есть другие данные
  ws.addConditionalFormatting({
    ref: `A2:A${lastRow}`,
    rules: [
      {
        type: 'expression',
        formulae: ['AND(LEN(TRIM($A2))=0,OR(LEN($B2)>0,LEN($C2)>0))'],
        style: {
          fill: {
            type: 'pattern',
            pattern: 'solid',
            bgColor: { argb: 'FFFFCDD2' }
          }
        }
      }
    ]
  })

  const buffer = await wb.xlsx.writeBuffer()
  return new Blob([buffer], { type: XLSX_MIME })
}

/* ----------------------------- Парсинг ----------------------------- */

// Приводит значение ячейки ExcelJS (формула, rich text, ссылка…) к примитиву
function cellToPrimitive(v) {
  if (v === null || v === undefined) return null
  if (v instanceof Date) return v.toISOString()
  if (typeof v === 'object') {
    if ('result' in v) return cellToPrimitive(v.result)
    if ('richText' in v) return v.richText.map((t) => t.text).join('')
    if ('text' in v) return cellToPrimitive(v.text)
    return null // ошибки (#N/A и т.п.)
  }
  return v
}

// Строки обрезаем по краям, пустые -> null, числа оставляем числами
function normalize(v) {
  const value = cellToPrimitive(v)
  if (typeof value === 'string') {
    const trimmed = value.trim()
    return trimmed === '' ? null : trimmed
  }
  return value
}

export async function parseExcelFile(file) {
  const ExcelJS = await loadExcelJS()
  const wb = new ExcelJS.Workbook()
  try {
    await wb.xlsx.load(await file.arrayBuffer())
  } catch {
    throw new Error('Не удалось прочитать файл. Убедитесь, что это .xlsx.')
  }

  const ws = wb.getWorksheet(SHEET_NAME) ?? wb.worksheets[0]
  if (!ws) throw new Error('В файле нет ни одного листа.')

  // Проверяем, что это наш шаблон
  const headerRow = ws.getRow(1)
  const headersOk = HEADERS.every((h, i) => {
    const text = normalize(headerRow.getCell(i + 1).value)
    return String(text ?? '').toLowerCase() === h
  })
  if (!headersOk) {
    throw new Error(
      `Заголовки не совпадают с шаблоном. Ожидается: ${HEADERS.join(', ')}.`
    )
  }

  const items = []
  const errors = []

  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r)
    const id = normalize(row.getCell(1).value)
    const nameRaw = normalize(row.getCell(2).value)
    const programRaw = normalize(row.getCell(3).value)

    // полностью пустая строка — пропускаем без ошибки
    if (id === null && nameRaw === null && programRaw === null) continue

    const rowErrors = []

    if (id === null) {
      rowErrors.push(
        nameRaw !== null
          ? 'Указано имя, но не заполнен ид'
          : 'Не заполнен обязательный ид'
      )
    }

    let program = null
    if (programRaw !== null) {
      program = PROGRAM_BY_LABEL[String(programRaw).toLowerCase()] ?? null
      if (program === null) {
        rowErrors.push(`Программа «${programRaw}» не входит в список допустимых`)
      }
    }

    if (rowErrors.length) {
      rowErrors.forEach((message) => errors.push({ row: r, message }))
      continue
    }

    items.push({
      id,
      name: nameRaw === null ? null : String(nameRaw),
      program
    })
  }

  return { items, errors }
}
