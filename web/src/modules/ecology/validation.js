import { MAX_ROWS_PER_SHEET, sourceCell } from '../../lib/workbook.js'

export const SCHEMA_VERSION = '0.1.0'
export const RULESET_VERSION = 'structural-0.1.0'

export const SOURCE_FIELDS = [
  { key: 'id', label: 'Номер источника', required: true },
  { key: 'name', label: 'Наименование', required: false },
  { key: 'x', label: 'Координата X / широта', required: false },
  { key: 'y', label: 'Координата Y / долгота', required: false },
  { key: 'coordinateSystem', label: 'Система координат', required: false },
]

export const EMISSION_FIELDS = [
  { key: 'sourceId', label: 'Номер источника', required: true },
  { key: 'substance', label: 'Вещество / код', required: true },
  { key: 'amount', label: 'Величина выброса', required: false },
  { key: 'unit', label: 'Единица измерения', required: false },
]

const HEADER_PATTERNS = {
  id: [/^(?:номер|№|id|код)\s*(?:и?з[а-я]*|источника)?$/i, /номер.*источник/i, /^источник\s*№/i],
  name: [/^наименован/i, /^назван/i],
  x: [/^x$/i, /^координат[аы]?\s*x/i, /^широт/i],
  y: [/^y$/i, /^координат[аы]?\s*y/i, /^долгот/i],
  coordinateSystem: [/систем.*координат/i, /^(?:ск|crs)$/i],
  sourceId: [/номер.*источник/i, /^источник\s*№/i, /^источник$/i],
  substance: [/веществ/i, /загрязняющ/i, /^код\s*зв/i],
  amount: [/выброс.*(?:г\/с|т\/год|величин|масса)/i, /^(?:выброс|масса|значение)$/i],
  unit: [/единиц.*измер/i, /^ед\.?\s*изм/i, /^единица$/i],
}

export function suggestMapping(columns, fields) {
  const mapping = {}
  const used = new Set()
  for (const field of fields) {
    const matches = columns.filter(column => !used.has(column.column) && HEADER_PATTERNS[field.key]?.some(pattern => pattern.test(column.label.trim())))
    if (matches.length === 1) {
      mapping[field.key] = matches[0].column
      used.add(matches[0].column)
    }
  }
  return mapping
}

export function normalizeId(cell) {
  if (!cell) return ''
  return cell.text.replace(/\s+/g, ' ').trim()
}

export function parseDecimal(cell) {
  if (!cell || cell.kind === 'empty' || !cell.text.trim()) return { state: 'empty', value: null }
  if (cell.kind === 'number') return { state: 'valid', value: cell.raw }
  const text = cell.text.trim().replace(/[\u00a0\u202f\s]/g, '')
  if (/^(?:н\/п|неприменимо|нетданных|н\/д|—|-)$/i.test(text)) return { state: 'not-applicable', value: null }
  if (!/^[+-]?(?:\d+[.,]?\d*|[.,]\d+)$/.test(text)) return { state: 'invalid', value: null }
  const normalized = text.replace(',', '.')
  const number = Number(normalized)
  return Number.isFinite(number) ? { state: 'valid', value: number } : { state: 'invalid', value: null }
}

function issue(code, severity, cell, message, action, entity = '') {
  return {
    code,
    ruleVersion: RULESET_VERSION,
    severity,
    entity,
    location: cell ? { sheet: cell.sheet, cell: cell.address, row: Number(cell.address.match(/\d+$/)?.[0] ?? 0) } : null,
    observed: cell?.text ?? '',
    mergedFrom: cell?.master ?? null,
    message,
    action,
    basis: 'Структурная проверка данных; не является нормативным заключением',
  }
}

function mappedRow(worksheet, rowNumber, mapping) {
  return Object.fromEntries(Object.entries(mapping).map(([field, column]) => [field, sourceCell(worksheet, rowNumber, Number(column))]))
}

function hasAnyValue(fields) {
  return Object.values(fields).some(cell => cell?.text)
}

export function validateWorkbook(workbook, config, fileHash) {
  const sourceWorksheet = workbook.getWorksheet(config.sources.sheet)
  if (!sourceWorksheet) throw new Error('Выберите лист источников.')
  if (!config.sources.mapping.id) throw new Error('Укажите колонку с номером источника.')
  const emissionsConfig = config.emissions?.sheet ? config.emissions : null
  const emissionWorksheet = emissionsConfig ? workbook.getWorksheet(emissionsConfig.sheet) : null
  if (emissionsConfig && (!emissionWorksheet || !emissionsConfig.mapping.sourceId || !emissionsConfig.mapping.substance)) {
    throw new Error('Для листа выбросов укажите колонки источника и вещества.')
  }
  if ([sourceWorksheet, emissionWorksheet].some(sheet => sheet?.rowCount > MAX_ROWS_PER_SHEET)) {
    throw new Error('Один из выбранных листов превышает 10 000 строк.')
  }

  const issues = []
  const sources = []
  const emissions = []
  const sourceIds = new Map()

  for (let rowNumber = Number(config.sources.headerRow) + 1; rowNumber <= sourceWorksheet.rowCount; rowNumber += 1) {
    const fields = mappedRow(sourceWorksheet, rowNumber, config.sources.mapping)
    if (!hasAnyValue(fields)) continue
    const id = normalizeId(fields.id)
    const record = { id, row: rowNumber, sheet: sourceWorksheet.name, fields }
    sources.push(record)
    if (!id) {
      issues.push(issue('SOURCE_ID_EMPTY', 'error', fields.id, 'У источника не указан номер.', 'Заполните номер или исключите строку из таблицы источников.'))
    } else if (fields.id.master) {
      issues.push(issue('SOURCE_ID_MERGED', 'info', fields.id, `Номер источника «${id}» получен из объединённой ячейки ${fields.id.master}.`, 'Проверьте, является ли эта строка отдельным источником или продолжением предыдущей записи.', id))
    } else if (sourceIds.has(id)) {
      issues.push(issue('SOURCE_ID_DUPLICATE', 'error', fields.id, `Номер источника «${id}» уже встречался в ${sourceIds.get(id)}.`, 'Проверьте, является ли это отдельным источником или продолжением записи.', id))
    } else {
      sourceIds.set(id, fields.id.address)
    }
    const x = parseDecimal(fields.x)
    const y = parseDecimal(fields.y)
    if ((x.state === 'valid') !== (y.state === 'valid') && (fields.x || fields.y)) {
      issues.push(issue('COORDINATE_PAIR_INCOMPLETE', 'warning', (x.state === 'valid' ? fields.y : fields.x) || fields.x || fields.y, 'Указана только одна координата пары или вторая не распознана.', 'Проверьте обе координаты и формат числа.', id))
    }
    if (x.state === 'invalid' || y.state === 'invalid') {
      issues.push(issue('COORDINATE_NUMBER_INVALID', 'warning', x.state === 'invalid' ? fields.x : fields.y, 'Координата не распознана как число.', 'Уточните формат координаты и систему координат.', id))
    }
    if (x.state === 'valid' && y.state === 'valid') {
      const system = fields.coordinateSystem?.text.trim().toUpperCase() ?? ''
      if (!system) {
        issues.push(issue('COORDINATE_SYSTEM_UNKNOWN', 'info', fields.coordinateSystem || fields.x, 'Система координат не указана; диапазон и положение точки не проверялись.', 'Укажите СК или подтвердите её вручную.', id))
      } else if (/WGS\s*84|EPSG:?4326|ШИРОТ/i.test(system) && (x.value < -90 || x.value > 90 || y.value < -180 || y.value > 180)) {
        issues.push(issue('WGS84_RANGE_INVALID', 'warning', fields.x, 'Координаты выходят за диапазон широты и долготы WGS84.', 'Проверьте порядок осей, СК и исходные значения.', id))
      }
    }
  }
  if (!sources.length) throw new Error('На листе источников не найдено строк после заголовка. Проверьте лист и строку заголовков.')

  if (emissionWorksheet) {
    for (let rowNumber = Number(emissionsConfig.headerRow) + 1; rowNumber <= emissionWorksheet.rowCount; rowNumber += 1) {
      const fields = mappedRow(emissionWorksheet, rowNumber, emissionsConfig.mapping)
      if (!hasAnyValue(fields)) continue
      const sourceId = normalizeId(fields.sourceId)
      const substance = fields.substance?.text.trim() ?? ''
      const amount = parseDecimal(fields.amount)
      emissions.push({ sourceId, substance, amount, row: rowNumber, sheet: emissionWorksheet.name, fields })
      if (!sourceId) issues.push(issue('EMISSION_SOURCE_EMPTY', 'error', fields.sourceId, 'В строке выброса нет номера источника.', 'Укажите номер источника из листа источников.'))
      else if (!sourceIds.has(sourceId)) issues.push(issue('EMISSION_SOURCE_MISSING', 'error', fields.sourceId, `Источник «${sourceId}» не найден на листе источников.`, 'Проверьте номер и выбранный лист источников.', sourceId))
      if (!substance) issues.push(issue('SUBSTANCE_EMPTY', 'warning', fields.substance, 'В строке выброса не указано вещество или код.', 'Заполните вещество или уточните структуру листа.', sourceId))
      if (fields.amount && amount.state === 'invalid') issues.push(issue('AMOUNT_INVALID', 'warning', fields.amount, 'Величина выброса не распознана как число.', 'Проверьте десятичный разделитель и ячейку формулы.', sourceId))
      if (fields.amount?.kind === 'formula' && fields.amount.missingCachedResult) issues.push(issue('FORMULA_RESULT_MISSING', 'warning', fields.amount, 'У формулы нет сохранённого результата; величина выброса не проверена.', 'Откройте книгу в табличном редакторе, пересчитайте и сохраните её.', sourceId))
      if (amount.state === 'valid' && amount.value < 0) issues.push(issue('AMOUNT_NEGATIVE', 'warning', fields.amount, 'Величина выброса отрицательная.', 'Проверьте знак и единицу измерения.', sourceId))
      const unit = fields.unit?.text.trim().toLowerCase().replace(/\s+/g, '') ?? ''
      if (fields.unit && unit && !['г/с', 'г/сек', 'кг/ч', 'т/год', 'т/г', 'мг/с'].includes(unit)) {
        issues.push(issue('UNIT_UNRECOGNIZED', 'info', fields.unit, `Единица «${fields.unit.text}» не входит в короткий список распознаваемых.`, 'Уточните единицу и маппинг; значение не пересчитывалось.', sourceId))
      }
    }
  }

  return {
    schemaVersion: SCHEMA_VERSION,
    rulesetVersion: RULESET_VERSION,
    rulePackage: { id: 'ecology-structural', version: RULESET_VERSION, kind: 'structural', normative: false, expertReviewed: false },
    fileHash,
    mapping: config,
    summary: { sources: sources.length, emissions: emissions.length, issues: issues.length },
    sources,
    emissions,
    issues,
  }
}
