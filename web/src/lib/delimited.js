import { MAX_FILE_BYTES, MAX_ROWS_PER_SHEET, guessHeaderRow, openWorkbook } from './workbook.js'

const MAX_COLUMNS = 100
const DELIMITERS = [';', ',', '\t']

function decodeTable(bytes) {
  const input = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  if (input.byteLength > MAX_FILE_BYTES) throw new Error('Файл больше 20 МБ.')
  let text
  let encoding = 'UTF-8'
  try {
    text = new TextDecoder('utf-8', { fatal: true }).decode(input)
  } catch {
    encoding = 'Windows-1251'
    text = new TextDecoder('windows-1251', { fatal: true }).decode(input)
  }
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1)
  if (/[\u0000-\u0008\u000e-\u001f]/u.test(text)) {
    throw new Error('Текстовый файл содержит управляющие символы. Проверьте кодировку или сохраните таблицу как CSV/TSV.')
  }
  return { text, encoding }
}

// The row number is a CSV record number. startLine also records the physical
// source line when a quoted field contains line breaks.
export function parseDelimited(text, delimiter, maxRecords = MAX_ROWS_PER_SHEET, sample = false) {
  const rows = []
  let fields = []
  let field = ''
  let quoted = false
  let closed = false
  let line = 1
  let startLine = 1

  function finishField() {
    fields.push(field)
    field = ''
    closed = false
    if (fields.length > MAX_COLUMNS) throw new Error('В таблице больше 100 колонок; этот файл пока не обрабатывается.')
  }
  function finishRow() {
    finishField()
    rows.push({ fields, startLine })
    if (rows.length > maxRecords) throw new Error(`В таблице больше ${maxRecords} строк; этот файл пока не обрабатывается.`)
    fields = []
    startLine = line + 1
  }

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i]
    if (quoted) {
      if (char === '"') {
        if (text[i + 1] === '"') { field += '"'; i += 1 }
        else { quoted = false; closed = true }
      } else {
        field += char
        if (char === '\r') { if (text[i + 1] === '\n') { field += '\n'; i += 1 }; line += 1 }
        else if (char === '\n') line += 1
      }
      continue
    }
    if (char === delimiter) { finishField(); continue }
    if (char === '\r' || char === '\n') {
      finishRow()
      if (sample && rows.length === maxRecords) return rows
      if (char === '\r' && text[i + 1] === '\n') i += 1
      line += 1
      startLine = line
      continue
    }
    if (char === '"' && field === '' && !closed) { quoted = true; continue }
    if (char === '"' || (closed && !/\s/u.test(char))) {
      throw new Error(`Некорректные кавычки в строке ${line}. Проверьте разделитель и сохраните CSV повторно.`)
    }
    if (closed && /\s/u.test(char)) continue
    field += char
  }
  if (quoted) throw new Error(`Не закрыта кавычка в строке ${startLine}.`)
  if (field || fields.length || closed) finishRow()
  return rows
}

function detectDelimiter(text, extension) {
  if (extension === 'tsv') return '\t'
  const candidates = DELIMITERS.flatMap(delimiter => {
    try {
      const rows = parseDelimited(text, delimiter, 40, true).filter(row => row.fields.some(value => value.trim())).slice(0, 20)
      const widths = rows.map(row => row.fields.length)
      const counts = new Map()
      for (const width of widths) counts.set(width, (counts.get(width) || 0) + 1)
      const [width, count] = [...counts].sort((a, b) => b[1] - a[1] || b[0] - a[0])[0] || [0, 0]
      return width > 1 ? [{ delimiter, score: count * 100 + width * 10 - (rows.length - count) * 50 }] : []
    } catch { return [] }
  }).sort((a, b) => b.score - a.score)
  if (!candidates.length) throw new Error('Не удалось определить колонки CSV. Нужна таблица с разделителем «;», «,» или табуляцией.')
  if (candidates[1]?.score === candidates[0].score) {
    throw new Error('Разделитель CSV неоднозначен. Сохраните файл с разделителем «;» или как TSV.')
  }
  return candidates[0].delimiter
}

export async function openDelimitedTable(bytes, fileName) {
  const extension = fileName.split('.').pop().toLowerCase()
  if (!['csv', 'tsv'].includes(extension)) throw new Error('Для текстовой таблицы нужен файл .csv или .tsv.')
  const { text, encoding } = decodeTable(bytes)
  const delimiter = detectDelimiter(text, extension)
  const rows = parseDelimited(text, delimiter)
  if (!rows.some(row => row.fields.some(value => value.trim()))) throw new Error('В текстовой таблице нет данных.')
  const { default: ExcelJS } = await import('exceljs')
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('Таблица')
  worksheet.sourceLines = new Map()
  for (const [index, row] of rows.entries()) {
    worksheet.addRow(row.fields)
    worksheet.sourceLines.set(index + 1, row.startLine)
  }
  return {
    workbook,
    format: extension,
    importInfo: { encoding, delimiter: delimiter === '\t' ? 'табуляция' : delimiter, rowCoordinates: 'номер записи', sourceLines: true },
    sheets: [{ name: worksheet.name, rowCount: worksheet.rowCount, columnCount: worksheet.columnCount,
      headerRow: guessHeaderRow(worksheet), overLimit: false }],
  }
}

export async function openTableFile(bytes, fileName) {
  if (/\.xlsx$/i.test(fileName)) {
    return { ...await openWorkbook(bytes), format: 'xlsx' }
  }
  return openDelimitedTable(bytes, fileName)
}
