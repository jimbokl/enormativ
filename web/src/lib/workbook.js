export const MAX_FILE_BYTES = 20 * 1024 * 1024
export const MAX_ROWS_PER_SHEET = 10000

export function cellValue(cell) {
  const value = cell?.value
  if (value === null || value === undefined) return { text: '', kind: 'empty', raw: null }
  if (typeof value === 'number') return { text: String(value), kind: 'number', raw: value }
  if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE', kind: 'boolean', raw: value }
  if (typeof value === 'string') return { text: value.trim(), kind: 'text', raw: value }
  if (value instanceof Date) return { text: value.toISOString(), kind: 'date', raw: value.toISOString() }
  if (value.formula || value.sharedFormula) {
    const result = value.result
    return {
      text: result === null || result === undefined ? '' : String(result),
      kind: 'formula',
      raw: result ?? null,
      formula: value.formula || value.sharedFormula,
      missingCachedResult: result === null || result === undefined,
    }
  }
  if (value.richText) return { text: value.richText.map(part => part.text).join('').trim(), kind: 'richText', raw: value.richText.map(part => part.text).join('') }
  if (typeof value.text === 'string') return { text: value.text.trim(), kind: 'hyperlink', raw: value.text }
  if (value.error) return { text: String(value.error), kind: 'error', raw: value.error }
  return { text: String(cell.text ?? ''), kind: 'unknown', raw: null }
}

export function sourceCell(worksheet, rowNumber, columnNumber) {
  if (!columnNumber) return null
  const cell = worksheet.getRow(rowNumber).getCell(columnNumber)
  const parsed = cellValue(cell)
  const master = cell.isMerged && cell.master?.address !== cell.address ? cell.master.address : null
  return {
    sheet: worksheet.name,
    address: cell.address,
    master,
    ...parsed,
  }
}

export function worksheetColumns(worksheet, headerRowNumber) {
  const row = worksheet.getRow(headerRowNumber)
  const columns = []
  for (let column = 1; column <= Math.min(worksheet.columnCount, 100); column += 1) {
    const label = cellValue(row.getCell(column)).text
    if (label) columns.push({ column, label, address: row.getCell(column).address })
  }
  return columns
}

export function guessHeaderRow(worksheet) {
  let bestRow = 1
  let bestScore = -1
  for (let rowNumber = 1; rowNumber <= Math.min(worksheet.rowCount, 20); rowNumber += 1) {
    const columns = worksheetColumns(worksheet, rowNumber)
    const score = columns.length + columns.filter(column => /источник|веществ|координат|выброс|номер|код|наименован/i.test(column.label)).length
    if (score > bestScore) {
      bestScore = score
      bestRow = rowNumber
    }
  }
  return bestRow
}

export function previewWorksheet(worksheet, headerRowNumber, limit = 5) {
  const columns = worksheetColumns(worksheet, headerRowNumber)
  const rows = []
  for (let rowNumber = headerRowNumber + 1; rowNumber <= worksheet.rowCount && rows.length < limit; rowNumber += 1) {
    const cells = columns.map(column => sourceCell(worksheet, rowNumber, column.column))
    if (cells.some(cell => cell.text)) rows.push({ rowNumber, cells })
  }
  return { columns, rows }
}

export async function openWorkbook(bytes) {
  if (bytes.byteLength > MAX_FILE_BYTES) throw new Error('Файл больше 20 МБ. Для первой версии выберите меньшую книгу.')
  const { default: ExcelJS } = await import('exceljs')
  const workbook = new ExcelJS.Workbook()
  try {
    await workbook.xlsx.load(bytes)
  } catch {
    throw new Error('Не удалось прочитать XLSX. Проверьте, что файл не повреждён и не защищён паролем.')
  }
  if (!workbook.worksheets.length) throw new Error('В книге не найдено листов.')
  const sheets = workbook.worksheets.map(worksheet => ({
    name: worksheet.name,
    rowCount: worksheet.rowCount,
    columnCount: worksheet.columnCount,
    headerRow: guessHeaderRow(worksheet),
    overLimit: worksheet.rowCount > MAX_ROWS_PER_SHEET,
  }))
  return { workbook, sheets }
}

export async function sha256(bytes) {
  const subtle = globalThis.crypto?.subtle
  if (!subtle) throw new Error('Для обработки файла требуется HTTPS. На этом адресе защищённое соединение пока недоступно; повторите попытку позже.')
  const digest = await subtle.digest('SHA-256', bytes)
  return Array.from(new Uint8Array(digest), value => value.toString(16).padStart(2, '0')).join('')
}
