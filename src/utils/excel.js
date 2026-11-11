import {
  COLUMNS,
  LISTS_SHEET_NAME,
  MAX_ROWS,
  SHEET_NAME,
  listOptions
} from '../constants'

const XLSX_MIME =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

const HEADER_REQUIRED = 'D32F2F'
const HEADER_OPTIONAL = '607D8B'

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

function columnLetter(index) {
  let n = index + 1
  let letters = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    letters = String.fromCharCode(65 + rem) + letters
    n = Math.floor((n - 1) / 26)
  }
  return letters
}

function normalizeColumns(columns) {
  if (!Array.isArray(columns) || columns.length === 0) {
    throw new Error('Нужно задать хотя бы одну колонку.')
  }

  const headers = new Set()
  const keys = new Set()

  return columns.map((column) => {
    const header = typeof column?.header === 'string' ? column.header.trim() : ''
    const key = typeof column?.key === 'string' ? column.key.trim() : ''
    if (!header) throw new Error('У каждой колонки должен быть заголовок.')
    if (!key) throw new Error(`У колонки «${header}» должен быть ключ объекта.`)

    const headerId = header.toLowerCase()
    if (headers.has(headerId)) throw new Error(`Повторяется заголовок «${header}».`)
    if (keys.has(key)) throw new Error(`Повторяется ключ «${key}».`)
    headers.add(headerId)
    keys.add(key)

    let list = null
    if (column.list) {
      if (!Array.isArray(column.list) || !column.list.length) {
        throw new Error(`Список колонки «${header}» пуст.`)
      }
      const labels = new Set()
      for (const item of column.list) {
        const isOption = item !== null && typeof item === 'object'
        const label = isOption ? String(item.label ?? '').trim() : String(item ?? '').trim()
        if (!label) throw new Error(`В списке «${header}» есть пустое значение.`)
        if (isOption && !Object.prototype.hasOwnProperty.call(item, 'value')) {
          throw new Error(`У пункта «${label}» списка «${header}» нет value.`)
        }
        const labelId = label.toLowerCase()
        if (labels.has(labelId)) {
          throw new Error(`В списке «${header}» повторяется «${label}».`)
        }
        labels.add(labelId)
      }
      list = listOptions(column)
    }

    return {
      header,
      key,
      required: Boolean(column.required),
      width: column.width,
      list
    }
  })
}

/* ---------------------------- Шаблон ---------------------------- */

// xlsx-js-style (как и SheetJS CE) не записывает валидацию данных,
// поэтому XML-блок <dataValidations> добавляем в готовый файл вручную.
function buildDataValidationsXml(rules) {
  if (!rules.length) return ''

  const body = rules
    .map(
      (rule) =>
        `<dataValidation type="${rule.type}" allowBlank="${rule.allowBlank ? '1' : '0'}"` +
        ` showErrorMessage="1" errorStyle="stop"` +
        ` errorTitle="${escapeXml(rule.errorTitle)}"` +
        ` error="${escapeXml(rule.error)}"` +
        ` sqref="${rule.sqref}"><formula1>${escapeXml(rule.formula)}</formula1></dataValidation>`
    )
    .join('')

  return `<dataValidations count="${rules.length}">${body}</dataValidations>`
}

function validationRules(columns, maxRows) {
  const last = maxRows + 1
  const rules = []
  let listColumn = 0

  columns.forEach((column, index) => {
    const letter = columnLetter(index)
    const sqref = `${letter}2:${letter}${last}`

    if (column.list) {
      const labels = column.list.map((option) => option.label)
      const shown =
        labels.join(', ').length > 180
          ? 'Выберите значение из списка.'
          : `Выберите значение из списка: ${labels.join(', ')}.`
      rules.push({
        type: 'list',
        allowBlank: !column.required,
        errorTitle: 'Недопустимое значение',
        error: column.required ? `Поле «${column.header}» обязательно. ${shown}` : shown,
        sqref,
        formula: `'${LISTS_SHEET_NAME}'!$${columnLetter(listColumn)}$1:$${columnLetter(listColumn)}$${labels.length}`
      })
      listColumn += 1
      return
    }

    if (column.required) {
      rules.push({
        type: 'custom',
        allowBlank: false,
        errorTitle: 'Обязательное поле',
        error: `Заполните поле «${column.header}».`,
        sqref,
        formula: `LEN(TRIM(${letter}2))>0`
      })
    }
  })

  return rules
}

function appendListSheet(XLSX, wb, columns) {
  const lists = columns.filter((column) => column.list)
  if (!lists.length) return

  if (wb.SheetNames.includes(LISTS_SHEET_NAME)) {
    throw new Error(`Имя листа «${LISTS_SHEET_NAME}» занято.`)
  }

  const height = Math.max(...lists.map((column) => column.list.length))
  const rows = Array.from({ length: height }, (_, row) =>
    lists.map((column) => column.list[row]?.label ?? null)
  )

  const ws = XLSX.utils.aoa_to_sheet(rows)
  XLSX.utils.book_append_sheet(wb, ws, LISTS_SHEET_NAME)

  wb.Workbook = wb.Workbook || {}
  wb.Workbook.Sheets = wb.SheetNames.map((_, index) =>
    index === wb.SheetNames.length - 1 ? { Hidden: 1 } : { Hidden: 0 }
  )
}

async function patchSheetXml(buffer, rules) {
  const dv = buildDataValidationsXml(rules)
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

  if (dv) {
    // По схеме OOXML <dataValidations> стоит после <sheetData> и перед
    // hyperlinks / pageMargins / ignoredErrors / drawing и т.д.
    const next = xml.match(
      /<(hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|ignoredErrors|drawing|legacyDrawing|tableParts|extLst)\b/
    )
    xml = next
      ? xml.replace(next[0], () => `${dv}${next[0]}`)
      : xml.replace('</worksheet>', () => `${dv}</worksheet>`)
  }

  zip.file(path, xml)
  return zip.generateAsync({ type: 'blob', mimeType: XLSX_MIME })
}

export async function buildTemplateBlob(columns = COLUMNS, options = {}) {
  const spec = normalizeColumns(columns)
  const sheetName = options.sheetName ?? SHEET_NAME
  const maxRows = options.maxRows ?? MAX_ROWS
  const XLSX = await loadXLSX()

  const headerStyle = (rgb) => ({
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: { patternType: 'solid', fgColor: { rgb } },
    alignment: { vertical: 'center', horizontal: 'left' }
  })

  const ws = XLSX.utils.aoa_to_sheet([spec.map((column) => column.header)])
  spec.forEach((column, index) => {
    const address = `${columnLetter(index)}1`
    ws[address].s = headerStyle(column.required ? HEADER_REQUIRED : HEADER_OPTIONAL)
  })
  ws['!cols'] = spec.map((column) => ({ wch: column.width ?? 18 }))
  ws['!rows'] = [{ hpt: 22 }]

  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, sheetName)
  appendListSheet(XLSX, wb, spec)

  const buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' })
  return patchSheetXml(buffer, validationRules(spec, maxRows))
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

export async function parseExcelFile(file, columns = COLUMNS, options = {}) {
  const spec = normalizeColumns(columns)
  const sheetName = options.sheetName ?? SHEET_NAME
  const XLSX = await loadXLSX()

  let wb
  try {
    wb = XLSX.read(await file.arrayBuffer(), { type: 'array' })
  } catch {
    throw new Error('Не удалось прочитать файл. Убедитесь, что это .xlsx.')
  }

  const sheet =
    wb.Sheets[sheetName] ??
    wb.SheetNames.map((name) => wb.Sheets[name]).find(
      (candidate) => candidate && candidate !== wb.Sheets[LISTS_SHEET_NAME]
    )
  const ws = sheet
  if (!ws || !ws['!ref']) throw new Error('Лист пуст или не найден.')

  const headersOk = spec.every(
    (column, index) =>
      String(cellValue(ws, XLSX.utils.encode_cell({ r: 0, c: index })) ?? '')
        .toLowerCase() === column.header.toLowerCase()
  )
  if (!headersOk) {
    throw new Error(
      `Заголовки не совпадают с шаблоном. Ожидается: ${spec.map((column) => column.header).join(', ')}.`
    )
  }

  const lastRow = XLSX.utils.decode_range(ws['!ref']).e.r + 1
  const items = []
  const errors = []

  for (let r = 2; r <= lastRow; r++) {
    const raw = spec.map((_, index) =>
      cellValue(ws, XLSX.utils.encode_cell({ r: r - 1, c: index }))
    )
    if (raw.every((value) => value === null)) continue

    const rowErrors = []
    const item = {}

    spec.forEach((column, index) => {
      const value = raw[index]
      if (value === null) {
        if (column.required) {
          rowErrors.push({
            header: column.header,
            message: `Не заполнено обязательное поле «${column.header}»`
          })
        }
        item[column.key] = null
        return
      }

      if (column.list) {
        const option = column.list.find(
          (entry) => entry.label.toLowerCase() === String(value).toLowerCase()
        )
        if (!option) {
          rowErrors.push({
            header: column.header,
            message: `Значение «${value}» не входит в список колонки «${column.header}»`
          })
          return
        }
        item[column.key] = option.value
        return
      }

      item[column.key] = value
    })

    if (rowErrors.length) {
      rowErrors.forEach((error) => errors.push({ row: r, ...error }))
      continue
    }

    items.push(item)
  }

  return { items, errors }
}
