import {
  HEADERS,
  MAX_ROWS,
  PROGRAMS,
  PROGRAM_BY_LABEL,
  SHEET_NAME
} from '../constants'

const XLSX_MIME =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

// Библиотеки грузим лениво — только при первом скачивании шаблона / загрузке файла
async function loadXLSX() {
  const mod = await import('xlsx-js-style')
  return mod.default ?? mod
}
async function loadJSZip() {
  const mod = await import('jszip')
  return mod.default ?? mod
}

const escapeXml = (s) =>
  String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')

/* ---------------------------- Шаблон ---------------------------- */

// xlsx-js-style (как и SheetJS CE) не записывает валидацию данных,
// поэтому XML-блок <dataValidations> добавляем в готовый файл вручную.
function buildDataValidationsXml() {
  const last = MAX_ROWS + 1
  const programs = PROGRAMS.map((p) => p.label).join(',')

  // Формула относительная: для диапазона B2:B{last} она считается от ячейки B2
  const nameRule =
    `<dataValidation type="custom" allowBlank="1" showErrorMessage="1" errorStyle="stop"` +
    ` errorTitle="${escapeXml('Сначала заполните ид')}"` +
    ` error="${escapeXml('Имя можно ввести только после заполнения поля «ид».')}"` +
    ` sqref="B2:B${last}"><formula1>${escapeXml('LEN(TRIM($A2))>0')}</formula1></dataValidation>`

  const programRule =
    `<dataValidation type="list" allowBlank="1" showErrorMessage="1" errorStyle="stop"` +
    ` errorTitle="${escapeXml('Недопустимое значение')}"` +
    ` error="${escapeXml(`Выберите значение из списка: ${programs.replace(/,/g, ', ')}.`)}"` +
    ` sqref="C2:C${last}"><formula1>${escapeXml(`"${programs}"`)}</formula1></dataValidation>`

  return `<dataValidations count="2">${nameRule}${programRule}</dataValidations>`
}

async function patchSheetXml(buffer) {
  const JSZip = await loadJSZip()
  const zip = await JSZip.loadAsync(buffer)
  const path = 'xl/worksheets/sheet1.xml'
  let xml = await zip.file(path).async('string')

  // Закрепляем строку заголовка
  xml = xml.replace(
    /<sheetView\b([^>]*?)\/>/,
    (_, attrs) =>
      `<sheetView${attrs}><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>` +
      `<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView>`
  )

  // По схеме OOXML <dataValidations> стоит после <sheetData> и перед
  // hyperlinks / pageMargins / ignoredErrors / drawing и т.д.
  const dv = buildDataValidationsXml()
  const next = xml.match(
    /<(hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|ignoredErrors|drawing|legacyDrawing|tableParts|extLst)\b/
  )
  xml = next
    ? xml.replace(next[0], () => `${dv}${next[0]}`)
    : xml.replace('</worksheet>', () => `${dv}</worksheet>`)

  zip.file(path, xml)
  return zip.generateAsync({ type: 'blob', mimeType: XLSX_MIME })
}

export async function buildTemplateBlob() {
  const XLSX = await loadXLSX()

  const headerStyle = (rgb) => ({
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: { patternType: 'solid', fgColor: { rgb } },
    alignment: { vertical: 'center', horizontal: 'left' }
  })

  const ws = XLSX.utils.aoa_to_sheet([HEADERS])
  // «ид» — красный (обязательное поле), остальные — серые
  ws.A1.s = headerStyle('D32F2F')
  ws.B1.s = headerStyle('607D8B')
  ws.C1.s = headerStyle('607D8B')
  ws['!cols'] = [{ wch: 16 }, { wch: 36 }, { wch: 22 }]
  ws['!rows'] = [{ hpt: 22 }]

  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, SHEET_NAME)

  const buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' })
  return patchSheetXml(buffer)
}

/* ----------------------------- Парсинг ----------------------------- */

// Значение ячейки -> примитив. Строки trim, пустые -> null, числа остаются числами
function cellValue(ws, address) {
  const cell = ws[address]
  if (!cell || cell.t === 'e') return null // нет ячейки или ошибка (#N/A)
  const v = cell.v
  if (v === undefined || v === null) return null
  if (typeof v === 'string') {
    const trimmed = v.trim()
    return trimmed === '' ? null : trimmed
  }
  return v
}

export async function parseExcelFile(file) {
  const XLSX = await loadXLSX()

  let wb
  try {
    wb = XLSX.read(await file.arrayBuffer(), { type: 'array' })
  } catch {
    throw new Error('Не удалось прочитать файл. Убедитесь, что это .xlsx.')
  }

  const ws = wb.Sheets[SHEET_NAME] ?? wb.Sheets[wb.SheetNames[0]]
  if (!ws || !ws['!ref']) throw new Error('Лист пуст или не найден.')

  // Проверяем, что это наш шаблон
  const headersOk = HEADERS.every(
    (h, i) =>
      String(cellValue(ws, XLSX.utils.encode_cell({ r: 0, c: i })) ?? '')
        .toLowerCase() === h
  )
  if (!headersOk) {
    throw new Error(
      `Заголовки не совпадают с шаблоном. Ожидается: ${HEADERS.join(', ')}.`
    )
  }

  const lastRow = XLSX.utils.decode_range(ws['!ref']).e.r + 1 // номер строки Excel
  const items = []
  const errors = []

  for (let r = 2; r <= lastRow; r++) {
    const id = cellValue(ws, `A${r}`)
    const nameRaw = cellValue(ws, `B${r}`)
    const programRaw = cellValue(ws, `C${r}`)

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
