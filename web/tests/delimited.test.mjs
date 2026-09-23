import test from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { openDelimitedTable, parseDelimited } from '../src/lib/delimited.js'
import { openWorkbook, sourceCell, worksheetColumns } from '../src/lib/workbook.js'
import { SOURCE_FIELDS, suggestMapping, validateWorkbook } from '../src/modules/ecology/validation.js'
import { compareSourceInventories } from '../src/modules/ecology/comparison.js'

const utf8 = value => new TextEncoder().encode(value)

test('CSV quotes, escaped quotes, CRLF and decimal commas preserve record and source line', async () => {
  const text = '\uFEFFНомер источника;Наименование;Широта;Долгота;Система координат\r\n' +
    '001;"Труба; северная";55,7;37,6;WGS84\r\n' +
    '002;"Площадка ""Юг""\r\nвторая строка";92;37,7;WGS84\r\n'
  const opened = await openDelimitedTable(utf8(text), 'источники.csv')
  const sheet = opened.workbook.getWorksheet('Таблица')
  assert.deepEqual(opened.importInfo, { encoding: 'UTF-8', delimiter: ';', rowCoordinates: 'номер записи', sourceLines: true })
  assert.equal(sourceCell(sheet, 2, 1).text, '001')
  assert.equal(sourceCell(sheet, 2, 2).raw, 'Труба; северная')
  assert.equal(sourceCell(sheet, 3, 2).raw, 'Площадка "Юг"\r\nвторая строка')
  assert.equal(sourceCell(sheet, 3, 3).sourceLine, 3)
  const config = { sources: { sheet: 'Таблица', headerRow: 1, mapping: suggestMapping(worksheetColumns(sheet, 1), SOURCE_FIELDS) } }
  const result = validateWorkbook(opened.workbook, config, 'csv-hash')
  assert.equal(result.summary.sources, 2)
  assert.equal(result.issues.find(issue => issue.code === 'WGS84_RANGE_INVALID').location.sourceLine, 3)
})

test('TSV and Windows-1251 retain strings, leading zeros and physical lines', async () => {
  const tsv = await openDelimitedTable(utf8('Номер источника\tНаименование\n001\tТруба\n'), 'input.tsv')
  assert.equal(tsv.importInfo.delimiter, 'табуляция')
  assert.equal(sourceCell(tsv.workbook.getWorksheet('Таблица'), 2, 1).text, '001')
  const cp1251 = Uint8Array.from([0xc8, 0xf1, 0xf2, 0xee, 0xf7, 0xed, 0xe8, 0xea, 0x3b, 0xc8, 0xec, 0xff, 0x0a, 0x30, 0x30, 0x31, 0x3b, 0xd2, 0xf0, 0xf3, 0xe1, 0xe0])
  const opened = await openDelimitedTable(cp1251, 'legacy.csv')
  assert.equal(opened.importInfo.encoding, 'Windows-1251')
  assert.equal(sourceCell(opened.workbook.getWorksheet('Таблица'), 2, 2).text, 'Труба')
})

test('equivalent CSV and XLSX produce the same structural findings and addresses', async () => {
  const rows = [
    ['Номер источника', 'Наименование', 'Широта', 'Долгота', 'Система координат'],
    ['001', 'Труба', '55,7', '37,6', 'WGS84'],
    ['001', 'Дубликат', '92', '37,6', 'WGS84'],
  ]
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Таблица')
  rows.forEach(row => sheet.addRow(row))
  const xlsx = await openWorkbook(await book.xlsx.writeBuffer())
  const csv = await openDelimitedTable(utf8(rows.map(row => row.join(';')).join('\n')), 'table.csv')
  const config = { sources: { sheet: 'Таблица', headerRow: 1, mapping: { id: 1, name: 2, x: 3, y: 4, coordinateSystem: 5 } } }
  const issues = opened => validateWorkbook(opened.workbook, config, 'same-hash').issues.map(issue => [issue.code, issue.location.cell, issue.observed])
  assert.deepEqual(issues(csv), issues(xlsx))
})

test('CSV and XLSX compare with traceable lines on the text side', async () => {
  const first = await openDelimitedTable(utf8('Номер источника;Широта;Долгота;Система координат\n001;55,7;37,6;WGS84\n'), 'first.csv')
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Источники')
  sheet.addRow(['Номер источника', 'Широта', 'Долгота', 'Система координат'])
  sheet.addRow(['001', 55.9, 37.6, 'WGS84'])
  const config = (name) => ({ sources: { sheet: name, headerRow: 1, mapping: { id: 1, x: 2, y: 3, coordinateSystem: 4 } } })
  const left = validateWorkbook(first.workbook, config('Таблица'), 'first')
  const right = validateWorkbook(book, config('Источники'), 'second')
  const result = compareSourceInventories(left, right, {
    primaryFile: { name: 'first.csv', hash: 'first' },
    secondaryFile: { name: 'second.xlsx', hash: 'second' },
    scopeConfirmed: true,
  })
  assert.equal(result.findings[0].code, 'SOURCE_COORDINATES_DIFFER')
  assert.equal(result.findings[0].location.sourceLine, 2)
  assert.equal(result.findings[0].relatedLocation.cell, 'B2')
})

test('malformed quotes, control bytes and excessive rows fail without a misleading validation result', async () => {
  assert.throws(() => parseDelimited('a;b\n"unclosed', ';'), /Не закрыта кавычка/)
  await assert.rejects(openDelimitedTable(Uint8Array.from([0, 1, 2]), 'bad.csv'), /управляющие символы/)
  const tooMany = `id;name\n${Array.from({ length: 10000 }, (_, i) => `${i};x`).join('\n')}`
  await assert.rejects(openDelimitedTable(utf8(tooMany), 'large.csv'), /больше 10000 строк/)
})
