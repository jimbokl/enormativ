import test from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { openWorkbook, worksheetColumns, sourceCell } from '../src/lib/workbook.js'
import { parseDecimal, suggestMapping, SOURCE_FIELDS, EMISSION_FIELDS, validateWorkbook } from '../src/modules/ecology/validation.js'

async function fixture() {
  const book = new ExcelJS.Workbook()
  const source = book.addWorksheet('Источники')
  source.addRow(['Номер источника', 'Наименование', 'Широта', 'Долгота', 'Система координат'])
  source.addRow(['001', 'Труба', 55.7, 37.6, 'WGS84'])
  source.addRow(['001', 'Дубликат', 92, 37.6, 'WGS84'])
  source.addRow(['003', 'Площадка', 55.8, null, 'WGS84'])
  const emission = book.addWorksheet('Выбросы')
  emission.addRow(['Номер источника', 'Вещество', 'Выброс', 'Единица измерения'])
  emission.addRow(['001', 'Пыль', '0,22', 'г/с'])
  emission.addRow(['002', 'NO₂', -1, 'г/с'])
  emission.addRow(['003', '', { formula: '1+1' }, 'неизвестно'])
  const bytes = await book.xlsx.writeBuffer()
  const opened = await openWorkbook(bytes)
  const sources = opened.workbook.getWorksheet('Источники')
  const emissions = opened.workbook.getWorksheet('Выбросы')
  const config = {
    sources: { sheet: 'Источники', headerRow: 1, mapping: suggestMapping(worksheetColumns(sources, 1), SOURCE_FIELDS) },
    emissions: { sheet: 'Выбросы', headerRow: 1, mapping: suggestMapping(worksheetColumns(emissions, 1), EMISSION_FIELDS) },
  }
  return { ...opened, config }
}

test('opens XLSX and keeps original cell addresses after mapping', async () => {
  const { workbook, sheets, config } = await fixture()
  assert.equal(sheets.length, 2)
  assert.equal(config.sources.mapping.id, 1)
  assert.equal(config.sources.mapping.x, 3)
  assert.equal(config.emissions.mapping.amount, 3)
  assert.equal(sourceCell(workbook.getWorksheet('Источники'), 3, 1).address, 'A3')
})

test('findings are deterministic and point to specific source cells', async () => {
  const { workbook, config } = await fixture()
  const first = validateWorkbook(workbook, config, 'fixture-hash')
  const second = validateWorkbook(workbook, config, 'fixture-hash')
  assert.deepEqual(first, second)
  assert.equal(first.summary.sources, 3)
  assert.equal(first.summary.emissions, 3)
  const byCode = Object.fromEntries(first.issues.map(item => [item.code, item]))
  assert.equal(byCode.SOURCE_ID_DUPLICATE.location.cell, 'A3')
  assert.equal(byCode.SOURCE_ID_DUPLICATE.observed, '001')
  assert.equal(byCode.WGS84_RANGE_INVALID.location.cell, 'C3')
  assert.equal(byCode.COORDINATE_PAIR_INCOMPLETE.location.cell, 'D4')
  assert.equal(byCode.EMISSION_SOURCE_MISSING.location.cell, 'A3')
  assert.equal(byCode.AMOUNT_NEGATIVE.location.cell, 'C3')
  assert.equal(byCode.FORMULA_RESULT_MISSING.location.cell, 'C4')
  assert.equal(byCode.SUBSTANCE_EMPTY.location.cell, 'B4')
  assert.equal(byCode.UNIT_UNRECOGNIZED.location.cell, 'D4')
  assert.ok(first.issues.every(item => item.ruleVersion && item.basis))
})

test('merged source identifier is flagged for review, not as a duplicate', async () => {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Источники')
  sheet.addRow(['Номер источника', 'Наименование'])
  sheet.addRow(['001', 'Основная строка'])
  sheet.addRow([null, 'Продолжение'])
  sheet.mergeCells('A2:A3')
  const opened = await openWorkbook(await book.xlsx.writeBuffer())
  const result = validateWorkbook(opened.workbook, {
    sources: { sheet: 'Источники', headerRow: 1, mapping: { id: 1, name: 2 } },
  }, 'merged-fixture')
  assert.equal(result.issues.some(item => item.code === 'SOURCE_ID_DUPLICATE'), false)
  const merged = result.issues.find(item => item.code === 'SOURCE_ID_MERGED')
  assert.equal(merged.location.cell, 'A3')
  assert.equal(merged.mergedFrom, 'A2')
})

test('numeric parser distinguishes commas, missing data, and invalid text', () => {
  assert.deepEqual(parseDecimal({ text: ' 1 234,56 ', kind: 'text' }), { state: 'valid', value: 1234.56 })
  assert.equal(parseDecimal({ text: 'н/д', kind: 'text' }).state, 'not-applicable')
  assert.equal(parseDecimal({ text: '0,1 г/с', kind: 'text' }).state, 'invalid')
})

test('source sheet and required mappings must contain usable data', async () => {
  const { workbook, config } = await fixture()
  assert.throws(() => validateWorkbook(workbook, { ...config, sources: { ...config.sources, mapping: {} } }, 'h'), /номер/)
  assert.throws(() => validateWorkbook(workbook, { ...config, sources: { ...config.sources, headerRow: 100 } }, 'h'), /не найдено строк/)
})
