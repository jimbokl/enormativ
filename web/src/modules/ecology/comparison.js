import { parseDecimal } from './validation.js'

export const COMPARISON_VERSION = 'ecology-cross-file-0.1.0'

function evidence(record, file) {
  const cell = record?.fields.id
  return record && cell ? {
    fileName: file.name,
    fileHash: file.hash,
    sheet: record.sheet,
    cell: cell.address,
    row: record.row,
    sourceLine: cell.sourceLine ?? null,
    observed: cell.text,
  } : null
}

function coordinateEvidence(record, file) {
  const cell = record.fields.x || record.fields.y || record.fields.id
  return {
    fileName: file.name,
    fileHash: file.hash,
    sheet: record.sheet,
    cell: cell.address,
    cells: [record.fields.x?.address, record.fields.y?.address].filter(Boolean),
    row: record.row,
    sourceLine: cell.sourceLine ?? null,
    observed: [record.fields.x?.text, record.fields.y?.text].filter(Boolean).join(' / '),
  }
}

function indexSources(sources) {
  const index = new Map()
  for (const record of sources) {
    if (!record.id) continue
    const matches = index.get(record.id) || []
    matches.push(record)
    index.set(record.id, matches)
  }
  return index
}

function finding(code, entity, location, relatedLocation, message, action) {
  return {
    code,
    ruleVersion: COMPARISON_VERSION,
    severity: 'warning',
    entity,
    location,
    relatedLocation,
    observed: location?.observed ?? '',
    relatedObserved: relatedLocation?.observed ?? '',
    message,
    action,
    basis: 'Сопоставление двух файлов по подтверждённому пользователем контексту; не является нормативным заключением',
  }
}

function coordinateState(record) {
  const fields = record.fields
  if (!fields.x && !fields.y) return { status: 'not-applicable', reason: 'Координаты не сопоставлены с колонками.' }
  const x = parseDecimal(fields.x)
  const y = parseDecimal(fields.y)
  const system = fields.coordinateSystem?.text.trim().toUpperCase().replace(/\s+/g, '') || ''
  if (x.state !== 'valid' || y.state !== 'valid' || !system) {
    return { status: 'unknown', reason: 'Нужны обе числовые координаты и явно указанная система координат.' }
  }
  return { status: 'ready', x: x.value, y: y.value, system }
}

export function compareSourceInventories(primary, secondary, { primaryFile, secondaryFile, scopeConfirmed }) {
  if (!scopeConfirmed) throw new Error('Подтвердите, что файлы относятся к одному объекту и сопоставимому состоянию.')
  if (!primaryFile?.hash || !secondaryFile?.hash) throw new Error('Для обоих файлов нужен SHA-256.')
  if (primaryFile.hash === secondaryFile.hash) throw new Error('Выбраны одинаковые файлы. Для сверки нужны два разных файла.')

  const left = indexSources(primary.sources)
  const right = indexSources(secondary.sources)
  const ids = [...new Set([...left.keys(), ...right.keys()])].sort((a, b) => a.localeCompare(b, 'ru', { numeric: true }))
  const evaluations = []
  const findings = []

  for (const id of ids) {
    const first = left.get(id) || []
    const second = right.get(id) || []
    const firstEvidence = evidence(first[0], primaryFile)
    const secondEvidence = evidence(second[0], secondaryFile)
    if (first.length > 1 || second.length > 1 || first[0]?.fields.id?.master || second[0]?.fields.id?.master) {
      evaluations.push({ code: 'SOURCE_PRESENCE', entity: id, status: 'unknown', reason: 'Номер повторяется или получен из объединённой ячейки; однозначная связь не установлена.', evidence: [firstEvidence, secondEvidence].filter(Boolean) })
      continue
    }
    if (!first.length || !second.length) {
      const inPrimary = first.length > 0
      const location = inPrimary ? firstEvidence : secondEvidence
      evaluations.push({ code: 'SOURCE_PRESENCE', entity: id, status: 'issue', reason: inPrimary ? 'Номер есть только в первом файле.' : 'Номер есть только во втором файле.', evidence: [location] })
      findings.push(finding(inPrimary ? 'SOURCE_ONLY_PRIMARY' : 'SOURCE_ONLY_SECONDARY', id, location, null,
        `Источник «${id}» найден только в ${inPrimary ? 'первом' : 'втором'} файле.`,
        'Проверьте полноту таблиц, период и границы объекта; изменение может быть обоснованным.'))
      continue
    }

    evaluations.push({ code: 'SOURCE_PRESENCE', entity: id, status: 'pass', reason: 'Номер найден в обоих файлах.', evidence: [firstEvidence, secondEvidence] })
    const firstCoordinates = coordinateState(first[0])
    const secondCoordinates = coordinateState(second[0])
    const coordinatePair = [coordinateEvidence(first[0], primaryFile), coordinateEvidence(second[0], secondaryFile)]
    if (firstCoordinates.status === 'not-applicable' && secondCoordinates.status === 'not-applicable') {
      evaluations.push({ code: 'SOURCE_COORDINATES', entity: id, status: 'not-applicable', reason: 'Координаты не сопоставлены ни в одном файле.', evidence: coordinatePair })
    } else if (firstCoordinates.status !== 'ready' || secondCoordinates.status !== 'ready' || firstCoordinates.system !== secondCoordinates.system) {
      evaluations.push({ code: 'SOURCE_COORDINATES', entity: id, status: 'unknown', reason: firstCoordinates.system && secondCoordinates.system && firstCoordinates.system !== secondCoordinates.system ? 'У файлов разные обозначения систем координат; преобразование не выполнялось.' : 'Не хватает сопоставимых числовых координат или системы координат.', evidence: coordinatePair })
    } else if (Math.abs(firstCoordinates.x - secondCoordinates.x) > 1e-6 || Math.abs(firstCoordinates.y - secondCoordinates.y) > 1e-6) {
      evaluations.push({ code: 'SOURCE_COORDINATES', entity: id, status: 'issue', reason: 'Числовые координаты отличаются более чем на 0,000001 в одинаково обозначенной системе координат.', evidence: coordinatePair })
      findings.push(finding('SOURCE_COORDINATES_DIFFER', id, coordinatePair[0], coordinatePair[1],
        `Координаты источника «${id}» отличаются между файлами.`,
        'Сверьте исходные значения, точность, систему координат и версию документов.'))
    } else {
      evaluations.push({ code: 'SOURCE_COORDINATES', entity: id, status: 'pass', reason: 'Числовые координаты совпадают с точностью 0,000001 в одинаково обозначенной системе координат.', evidence: coordinatePair })
    }
  }

  return {
    schemaVersion: '0.1.0',
    rulePackage: { id: 'ecology-cross-file', version: COMPARISON_VERSION, kind: 'data-consistency', normative: false, expertReviewed: false },
    files: [primaryFile, secondaryFile],
    context: { scopeConfirmed: true, confirmedBy: 'user' },
    summary: {
      firstSources: primary.sources.length,
      secondSources: secondary.sources.length,
      sharedIds: evaluations.filter(item => item.code === 'SOURCE_PRESENCE' && item.status === 'pass').length,
      issues: findings.length,
      unknown: evaluations.filter(item => item.status === 'unknown').length,
      notApplicable: evaluations.filter(item => item.status === 'not-applicable').length,
    },
    evaluations,
    findings,
  }
}
