import test from 'node:test'
import assert from 'node:assert/strict'
import { csvCell } from '../src/lib/csv.js'

test('CSV export keeps workbook formulas and formula-like text inert', () => {
  for (const value of ['=1+1', '+SUM(1,2)', '-HYPERLINK("https://example.com")', '@SUM(1)', '  =1+1', '\t=1+1']) {
    const cell = csvCell(value)
    assert.equal(cell, `"'${value.replaceAll('"', '""')}"`)
  }
  assert.equal(csvCell('обычный; "текст"'), '"обычный; ""текст"""')
  assert.equal(csvCell(null), '""')
})
