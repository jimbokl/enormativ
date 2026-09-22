import test from 'node:test'
import assert from 'node:assert/strict'
import ExcelJS from 'exceljs'
import { compareSourceInventories } from '../src/modules/ecology/comparison.js'
import { validateWorkbook } from '../src/modules/ecology/validation.js'

function inventory(rows, hash, withCoordinates = true) {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Источники')
  sheet.addRow(['Номер источника', 'Координата X', 'Координата Y', 'Система координат'])
  rows.forEach(row => sheet.addRow(row))
  return validateWorkbook(book, {
    sources: { sheet: 'Источники', headerRow: 1, mapping: withCoordinates ? { id: 1, x: 2, y: 3, coordinateSystem: 4 } : { id: 1 } },
  }, hash)
}

const files = { primaryFile: { name: 'инвентаризация.xlsx', hash: 'first' }, secondaryFile: { name: 'проект.xlsx', hash: 'second' }, scopeConfirmed: true }

test('compares source IDs and coordinates with both file addresses and stable outcomes', () => {
  const first = inventory([['001', 55.7, 37.6, 'WGS84'], ['002', 55.8, 37.7, 'WGS84']], 'first')
  const second = inventory([['001', 55.9, 37.6, 'WGS84'], ['003', 55.8, 37.7, 'WGS84']], 'second')
  const result = compareSourceInventories(first, second, files)
  assert.deepEqual(result, compareSourceInventories(first, second, files))
  assert.deepEqual(result.findings.map(item => item.code), ['SOURCE_COORDINATES_DIFFER', 'SOURCE_ONLY_PRIMARY', 'SOURCE_ONLY_SECONDARY'])
  assert.deepEqual(result.findings[0].location.cells, ['B2', 'C2'])
  assert.deepEqual(result.findings[0].relatedLocation.cells, ['B2', 'C2'])
  assert.equal(result.findings[0].location.fileHash, 'first')
  assert.equal(result.findings[0].relatedLocation.fileHash, 'second')
  assert.equal(result.summary.sharedIds, 1)
  assert.equal(result.summary.issues, 3)
  assert.equal(result.rulePackage.normative, false)
})

test('ambiguous IDs and differing coordinate systems stay unknown', () => {
  const first = inventory([['001', 55.7, 37.6, 'WGS84'], ['001', 55.7, 37.6, 'WGS84'], ['002', 55.8, 37.7, 'WGS84']], 'first')
  const second = inventory([['001', 55.7, 37.6, 'WGS84'], ['002', 55.8, 37.7, 'МСК']], 'second')
  const result = compareSourceInventories(first, second, files)
  assert.equal(result.summary.unknown, 2)
  assert.equal(result.findings.length, 0)
  assert.equal(result.evaluations.find(item => item.entity === '001').status, 'unknown')
  assert.equal(result.evaluations.find(item => item.entity === '002' && item.code === 'SOURCE_COORDINATES').status, 'unknown')
})

test('unmapped coordinates are not applicable and comparison needs confirmed context', () => {
  const first = inventory([['001']], 'first', false)
  const second = inventory([['001']], 'second', false)
  const result = compareSourceInventories(first, second, files)
  assert.equal(result.summary.notApplicable, 1)
  assert.equal(result.summary.issues, 0)
  assert.throws(() => compareSourceInventories(first, second, { ...files, scopeConfirmed: false }), /Подтвердите/)
  assert.throws(() => compareSourceInventories(first, second, { ...files, secondaryFile: files.primaryFile }), /одинаковые файлы/)
})
