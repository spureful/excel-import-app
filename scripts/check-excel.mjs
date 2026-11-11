import JSZip from 'jszip'
import { buildTemplateBlob, parseExcelFile } from '../src/utils/excel.js'

const blob = await buildTemplateBlob()
const zip = await JSZip.loadAsync(await blob.arrayBuffer())
const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string')
const workbook = await zip.file('xl/workbook.xml').async('string')
const lists = await zip.file('xl/worksheets/sheet2.xml').async('string')

const checks = [
  ['dataValidations', sheet.includes('<dataValidations count="2">')],
  ['required id', sheet.includes('LEN(TRIM(A2))&gt;0') && sheet.includes('sqref="A2:A3001"')],
  ['program list', sheet.includes("'_lists'!$A$1:$A$4") && sheet.includes('sqref="C2:C3001"')],
  ['frozen header', sheet.includes('ySplit="1"')],
  ['hidden lists', /name="_lists"[^>]*state="hidden"/.test(workbook) || /state="hidden"[^>]*name="_lists"/.test(workbook)],
  ['list values', lists.includes('новая') && lists.includes('демонтаж')]
]

const XLSX = (await import('xlsx-js-style')).default
const wb = XLSX.utils.book_new()
const ws = XLSX.utils.aoa_to_sheet([
  ['ид', 'имя', 'программа'],
  [101, 'Объект А', 'новая'],
  [null, null, null],
  [null, 'Без ид', 'новая'],
  [7, 'Пустая программа', 'нет такой'],
  [8, 'Регистр', 'Новая'],
  [null, 'Только имя', null]
])
delete ws.A4
delete ws.A7
XLSX.utils.book_append_sheet(wb, ws, 'Шаблон')
const filled = new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })])
const parsed = await parseExcelFile(filled)

const okItems =
  parsed.items.length === 2 &&
  parsed.items[0].id === 101 &&
  parsed.items[0].name === 'Объект А' &&
  parsed.items[0].program === 'new' &&
  parsed.items[1].id === 8 &&
  parsed.items[1].program === 'new'
const okErrors =
  parsed.errors.some((error) => error.row === 4 && error.header === 'ид') &&
  parsed.errors.some((error) => error.row === 5 && error.header === 'программа') &&
  parsed.errors.some((error) => error.row === 7 && error.header === 'ид')

console.log(JSON.stringify({ checks, okItems, okErrors, items: parsed.items, errors: parsed.errors }, null, 2))
if (checks.some(([, ok]) => !ok) || !okItems || !okErrors) process.exit(1)
