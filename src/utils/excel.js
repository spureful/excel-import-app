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

// xlsx-js-style грузим только при скачивании шаблона или разборе файла.
async function loadXLSX() {
  const mod = await import('xlsx-js-style')
  return mod.default ?? mod
}

// В пакете нет записи dataValidations, поэтому после XLSX.write
// в sheet XML добавляется только этот блок (и закрепление строки).
async function loadJSZip() {
  const mod = await import('jszip')
  return mod.default ?? mod
}

const escapeXml = (value) =>
  String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')

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

function headerStyle(rgb) {
  return {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: { patternType: 'solid', fgColor: { rgb } },
    alignment: { vertical: 'center', horizontal: 'left' }
  }
}

function listErrorText(column) {
  const labels = column.list.map((option) => option.label)
  const shown =
    labels.join(', ').length > 180
      ? 'Выберите значение из списка.'
      : `Выберите значение из списка: ${labels.join(', ')}.`
  return column.required ? `Поле «${column.header}» обязательно. ${shown}` : shown
}

// Правила лежат на листе так, как их описывает SheetJS (!dataValidations).
// Сериализацию делает patchSheetXml: штатный writer их выбрасывает.
function assignValidations(XLSX, ws, columns, maxRows) {
  const rules = []
  let listColumn = 0

  columns.forEach((column, index) => {
    const sqref = XLSX.utils.encode_range({
      s: { r: 1, c: index },
      e: { r: maxRows, c: index }
    })

    if (column.list) {
      const letter = XLSX.utils.encode_col(listColumn)
      const last = column.list.length
      rules.push({
        type: 'list',
        allowBlank: !column.required,
        errorStyle: 'stop',
        showErrorMessage: true,
        errorTitle: 'Недопустимое значение',
        error: listErrorText(column),
        sqref,
        formula1: `'${LISTS_SHEET_NAME.replace(/'/g, "''")}'!$${letter}$1:$${letter}$${last}`
      })
      listColumn += 1
      return
    }

    if (column.required) {
      const letter = XLSX.utils.encode_col(index)
      rules.push({
        type: 'custom',
        allowBlank: false,
        errorStyle: 'stop',
        showErrorMessage: true,
        errorTitle: 'Обязательное поле',
        error: `Заполните поле «${column.header}».`,
        sqref,
        formula1: `LEN(TRIM(${letter}2))>0`
      })
    }
  })

  if (rules.length) ws['!dataValidations'] = rules
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

function dataValidationsXml(rules) {
  const body = rules
    .map(
      (rule) =>
        `<dataValidation type="${rule.type}" allowBlank="${rule.allowBlank ? '1' : '0'}"` +
        ` showErrorMessage="1" errorStyle="${rule.errorStyle || 'stop'}"` +
        ` errorTitle="${escapeXml(rule.errorTitle || '')}"` +
        ` error="${escapeXml(rule.error || '')}"` +
        ` sqref="${rule.sqref}"><formula1>${escapeXml(rule.formula1)}</formula1></dataValidation>`
    )
    .join('')

  return `<dataValidations count="${rules.length}">${body}</dataValidations>`
}

async function patchSheetXml(buffer, ws) {
  const rules = ws['!dataValidations'] || []
  const JSZip = await loadJSZip()
  const zip = await JSZip.loadAsync(buffer)
  const path = 'xl/worksheets/sheet1.xml'
  const sheetFile = zip.file(path)
  if (!sheetFile) throw new Error('В книге нет первого листа.')

  let xml = await sheetFile.async('string')

  xml = xml.replace(
    /<sheetView\b([^>]*?)\/>/,
    (_, attrs) =>
      `<sheetView${attrs}><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>` +
      `<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView>`
  )

  if (rules.length) {
    const dv = dataValidationsXml(rules)
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

  const ws = XLSX.utils.aoa_to_sheet([spec.map((column) => column.header)])
  spec.forEach((column, index) => {
    const address = XLSX.utils.encode_cell({ r: 0, c: index })
    ws[address].s = headerStyle(column.required ? HEADER_REQUIRED : HEADER_OPTIONAL)
  })
  ws['!cols'] = spec.map((column) => ({ wch: column.width ?? 18 }))
  ws['!rows'] = [{ hpt: 22 }]
  assignValidations(XLSX, ws, spec, maxRows)

  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, sheetName)
  appendListSheet(XLSX, wb, spec)

  const buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' })
  return patchSheetXml(buffer, ws)
}

function normalizeCell(value) {
  if (value === undefined || value === null) return null
  if (typeof value === 'string') {
    const trimmed = value.trim()
    return trimmed === '' ? null : trimmed
  }
  if (typeof value === 'number' && !Number.isFinite(value)) return null
  return value
}

function readTable(XLSX, ws, spec) {
  const used = XLSX.utils.decode_range(ws['!ref'])
  if (used.e.r < 1) return []

  used.s = { r: 1, c: 0 }
  used.e.c = Math.max(used.e.c, spec.length - 1)

  return XLSX.utils.sheet_to_json(ws, {
    header: spec.map((column) => column.key),
    range: XLSX.utils.encode_range(used),
    defval: null,
    raw: true,
    blankrows: true
  })
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

  const ws =
    wb.Sheets[sheetName] ??
    wb.SheetNames
      .filter((name) => name !== LISTS_SHEET_NAME)
      .map((name) => wb.Sheets[name])
      .find(Boolean)
  if (!ws || !ws['!ref']) throw new Error('Лист пуст или не найден.')

  const headersOk = spec.every((column, index) => {
    const address = XLSX.utils.encode_cell({ r: 0, c: index })
    return String(normalizeCell(ws[address]?.v) ?? '').toLowerCase() === column.header.toLowerCase()
  })
  if (!headersOk) {
    throw new Error(
      `Заголовки не совпадают с шаблоном. Ожидается: ${spec.map((column) => column.header).join(', ')}.`
    )
  }

  const items = []
  const errors = []

  readTable(XLSX, ws, spec).forEach((row, index) => {
    const excelRow = index + 2
    const values = spec.map((column) => normalizeCell(row[column.key]))
    if (values.every((value) => value === null)) return

    const rowErrors = []
    const item = {}

    spec.forEach((column, columnIndex) => {
      const value = values[columnIndex]
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
      rowErrors.forEach((error) => errors.push({ row: excelRow, ...error }))
      return
    }

    items.push(item)
  })

  return { items, errors }
}
