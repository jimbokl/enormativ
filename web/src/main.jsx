import React, { useMemo, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import {
  ArrowDownRight, ArrowRight, Check, ChevronDown, CircleAlert, Download,
  FileSpreadsheet, FileText, Fingerprint, LockKeyhole, MapPin, RotateCcw,
  ScanSearch, ShieldCheck, Sparkles, UploadCloud, X,
} from 'lucide-react'
import { createExampleFile } from './lib/example.js'
import { csvCell } from './lib/csv.js'
import { MAX_FILE_BYTES, openWorkbook, previewWorksheet, sha256, worksheetColumns } from './lib/workbook.js'
import { compareSourceInventories } from './modules/ecology/comparison.js'
import { EMISSION_FIELDS, SOURCE_FIELDS, suggestMapping, validateWorkbook } from './modules/ecology/validation.js'
import './style.css'
import './portal.css'

const severities = {
  error: { label: 'Ошибка', className: 'severity-error' },
  warning: { label: 'Проверить', className: 'severity-warning' },
  info: { label: 'Уточнить', className: 'severity-info' },
}

function initialConfig(workbook, sheets) {
  const sourceSheet = sheets.find(sheet => /источник|иза/i.test(sheet.name)) || sheets[0]
  const emissionSheet = sheets.find(sheet => /выброс|веществ/i.test(sheet.name) && sheet.name !== sourceSheet.name)
  const make = (sheet, fields) => sheet ? {
    sheet: sheet.name,
    headerRow: sheet.headerRow,
    mapping: suggestMapping(worksheetColumns(workbook.getWorksheet(sheet.name), sheet.headerRow), fields),
  } : { sheet: '', headerRow: 1, mapping: {} }
  return { sources: make(sourceSheet, SOURCE_FIELDS), emissions: make(emissionSheet, EMISSION_FIELDS) }
}

function downloadBlob(filename, content, type) {
  const url = URL.createObjectURL(new Blob([content], { type }))
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = filename
  document.body.append(anchor)
  anchor.click()
  anchor.remove()
  window.setTimeout(() => URL.revokeObjectURL(url), 1000)
}

function issueCsv(result) {
  const rows = [['Уровень', 'Код', 'Файл', 'SHA-256', 'Лист', 'Ячейка', 'Исходное значение', 'Второй файл', 'SHA-256 второго файла', 'Лист второго файла', 'Ячейка второго файла', 'Значение второго файла', 'Объединена с', 'Сущность', 'Замечание', 'Действие', 'Основание', 'Версия правила']]
  for (const finding of result.issues) {
    rows.push([
      severities[finding.severity].label, finding.code, finding.location?.fileName, finding.location?.fileHash, finding.location?.sheet,
      finding.location?.cells?.join(' / ') || finding.location?.cell, finding.observed, finding.relatedLocation?.fileName, finding.relatedLocation?.fileHash,
      finding.relatedLocation?.sheet, finding.relatedLocation?.cells?.join(' / ') || finding.relatedLocation?.cell, finding.relatedObserved,
      finding.mergedFrom, finding.entity, finding.message, finding.action,
      finding.basis, finding.ruleVersion,
    ])
  }
  return `\uFEFF${rows.map(row => row.map(csvCell).join(';')).join('\r\n')}`
}

function withFile(finding, file) {
  return { ...finding, location: finding.location ? { ...finding.location, fileName: file.name, fileHash: file.hash } : null }
}

function Brand({ root = './' }) {
  return <a className="brand" href={root} aria-label="еНорматив — на главную">
    <img src={`${root}brand-mark.png`} alt="" />
    <span>e<span className="brand-accent">normativ</span><span className="brand-domain">.ru</span></span>
  </a>
}

function SiteHeader({ root = './', active = '' }) {
  return <header className="site-header" id="top"><div className="container header-inner"><Brand root={root} /><nav aria-label="Основная навигация"><a href={`${root}tools/check/`} aria-current={active === 'check' ? 'page' : undefined}>Проверить XLSX</a><a href={`${root}tools/compare/`} aria-current={active === 'compare' ? 'page' : undefined}>Сравнить файлы</a></nav><a className="header-action" href={active ? root : `${root}tools/check/`}>{active ? 'Все инструменты' : 'Начать проверку'} <ArrowRight size={16} /></a></div></header>
}

function SiteFooter({ root = './' }) {
  return <footer className="site-footer"><div className="container footer-inner"><Brand root={root} /><p>еНорматив · локальные проверки профессиональных данных</p><a href={root}>Все инструменты ↑</a></div></footer>
}

function FieldSelector({ field, columns, value, onChange }) {
  return <label className="field-selector">
    <span>{field.label}{field.required && <b aria-label="обязательно"> *</b>}</span>
    <span className="select-wrap">
      <select value={value || ''} onChange={event => onChange(event.target.value ? Number(event.target.value) : '')}>
        <option value="">Не выбрано</option>
        {columns.map(column => <option key={column.column} value={column.column}>{column.address.replace(/\d+$/, '')} · {column.label}</option>)}
      </select>
      <ChevronDown size={16} aria-hidden="true" />
    </span>
  </label>
}

function SheetMapping({ type, title, description, fields, sheets, workbook, config, onChange, optional = false }) {
  const worksheet = config.sheet ? workbook.getWorksheet(config.sheet) : null
  const columns = worksheet ? worksheetColumns(worksheet, Number(config.headerRow)) : []
  const preview = worksheet ? previewWorksheet(worksheet, Number(config.headerRow), 4) : null
  return <div className="mapping-card">
    <div className="mapping-heading">
      <div><span className="eyebrow">{type === 'sources' ? '01 / Основа' : '02 / Связи'}</span><h3>{title}</h3><p>{description}</p></div>
      {optional && <span className="optional">Необязательно</span>}
    </div>
    <div className="mapping-grid">
      <label className="field-selector"><span>Лист</span><span className="select-wrap">
        <select value={config.sheet} onChange={event => onChange({ ...config, sheet: event.target.value, headerRow: sheets.find(sheet => sheet.name === event.target.value)?.headerRow || 1, mapping: event.target.value ? suggestMapping(worksheetColumns(workbook.getWorksheet(event.target.value), sheets.find(sheet => sheet.name === event.target.value)?.headerRow || 1), fields) : {} })}>
          {optional && <option value="">Нет отдельного листа</option>}
          {!optional && <option value="">Выберите лист</option>}
          {sheets.map(sheet => <option key={sheet.name} value={sheet.name}>{sheet.name} · {sheet.rowCount} строк</option>)}
        </select><ChevronDown size={16} aria-hidden="true" /></span></label>
      {worksheet && <label className="field-selector"><span>Строка заголовков</span><span className="select-wrap">
        <select value={config.headerRow} onChange={event => { const headerRow = Number(event.target.value); onChange({ ...config, headerRow, mapping: suggestMapping(worksheetColumns(worksheet, headerRow), fields) }) }}>
          {Array.from({ length: Math.min(worksheet.rowCount, 30) }, (_, index) => index + 1).map(row => <option key={row} value={row}>Строка {row}</option>)}
        </select><ChevronDown size={16} aria-hidden="true" /></span></label>}
    </div>
    {worksheet && <>
      {worksheet.rowCount > 10000 && <p className="inline-alert"><CircleAlert size={17} /> Лист превышает лимит первой версии — 10 000 строк. Выберите меньший лист.</p>}
      <div className="field-grid">{fields.map(field => <FieldSelector key={field.key} field={field} columns={columns} value={config.mapping[field.key]} onChange={column => onChange({ ...config, mapping: { ...config.mapping, [field.key]: column } })} />)}</div>
      {preview?.rows.length > 0 && <div className="preview-wrap"><p className="preview-label">Предпросмотр исходных строк</p><div className="table-scroll"><table><thead><tr><th>№</th>{preview.columns.slice(0, 8).map(column => <th key={column.column}>{column.label}</th>)}</tr></thead><tbody>{preview.rows.map(row => <tr key={row.rowNumber}><th>{row.rowNumber}</th>{row.cells.slice(0, 8).map((cell, index) => <td key={index} title={`${cell.address}${cell.master ? ` · объединено с ${cell.master}` : ''}`}>{cell.text || <span className="empty-cell">—</span>}{cell.kind === 'formula' && <sup>ƒ</sup>}</td>)}</tr>)}</tbody></table></div></div>}
    </>}
  </div>
}

function ComparisonPanel({ inputRef, parsed, file, config, confirmed, busy, onLoad, onRemove, onMapping, onConfirm }) {
  return <section className="comparison-panel" aria-label="Сверка двух файлов">
    <div className="comparison-heading"><div><span className="eyebrow">Шаг 02 · второй файл</span><h3>Сопоставить два файла</h3><p>Сверим номера и координаты источников между двумя XLSX. Изменения в документах могут быть обоснованными; решение остаётся за экологом.</p></div></div>
    {!parsed ? <label className="comparison-picker"><FileSpreadsheet size={19} /><span>Добавить второй XLSX с источниками</span><input ref={inputRef} type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" disabled={busy} onChange={event => onLoad(event.target.files[0])} /></label> : <>
      <div className="comparison-file"><FileSpreadsheet size={19} /><strong>{file.name}</strong><span>SHA-256 {file.hash.slice(0, 12)}…</span><button type="button" className="icon-button" onClick={onRemove} aria-label="Удалить второй файл из памяти"><X size={18} /></button></div>
      <SheetMapping type="sources" title="Источники во втором файле" description="Подтвердите лист, строку заголовков и колонки. Номера будут сопоставлены с первым файлом." fields={SOURCE_FIELDS} sheets={parsed.sheets} workbook={parsed.workbook} config={config} onChange={onMapping} />
      <label className="comparison-confirm"><input type="checkbox" checked={confirmed} onChange={event => onConfirm(event.target.checked)} /><span>Подтверждаю, что оба файла относятся к одному объекту и сопоставимому состоянию данных.</span></label>
      <p className="comparison-note">Разные системы координат или неоднозначные номера дадут результат «неизвестно». Файлы не покидают браузер.</p>
    </>}
  </section>
}

function ResultRow({ finding, selected, onClick }) {
  const severity = severities[finding.severity]
  return <button type="button" className={`result-row ${selected ? 'selected' : ''}`} onClick={onClick}>
    <span className={`severity-dot ${severity.className}`}></span>
    <span className="result-main"><strong>{finding.message}</strong><small>{finding.code} · {finding.ruleVersion}</small></span>
    <span className="result-location">{finding.location ? `${finding.location.fileName ? `${finding.location.fileName} · ` : ''}${finding.location.sheet} · ${finding.location.cell}` : 'Без адреса'}</span>
    <ArrowRight size={17} aria-hidden="true" />
  </button>
}

function FindingDetail({ finding, rulesetVersion, hasIssues, hasUncertain, isComparison }) {
  if (!finding) return <aside className="issue-detail"><span className="detail-status severity-info">Готово</span><h3>{hasIssues ? 'В этом фильтре замечаний нет' : hasUncertain ? 'Есть неполные проверки' : 'Замечаний не найдено'}</h3><p>{hasIssues ? 'Выберите другой фильтр, чтобы увидеть найденные замечания.' : hasUncertain ? 'Откройте список источников выше: часть результатов требует уточнения или не проверялась.' : isComparison ? `Сверка двух файлов по правилам версии ${rulesetVersion} не выявила расхождений. Это не нормативное заключение.` : `Проверка применяла структурные правила версии ${rulesetVersion}.`}</p></aside>
  const place = location => location ? `${location.fileName ? `${location.fileName} · ` : ''}${location.sheet} · ${location.cell}${location.cells?.length > 1 ? ` / ${location.cells[1]}` : ''}` : '—'
  return <aside className="issue-detail">
    <span className={`detail-status ${severities[finding.severity].className}`}>{severities[finding.severity].label}</span>
    <h3>{finding.message}</h3>
    <div className="detail-location"><span>Исходное место</span><strong>{place(finding.location)}</strong></div>
    {finding.relatedLocation && <div className="detail-location"><span>Во втором файле</span><strong>{place(finding.relatedLocation)}</strong></div>}
    <div className="detail-block"><span>Исходное значение</span><p>{finding.observed || 'Пусто'}{finding.mergedFrom ? ` · объединена с ${finding.mergedFrom}` : ''}</p></div>
    {finding.relatedLocation && <div className="detail-block"><span>Значение во втором файле</span><p>{finding.relatedObserved || 'Пусто'}</p></div>}
    <div className="detail-block"><span>Что сделать</span><p>{finding.action}</p></div>
    <div className="detail-block"><span>Основание</span><p>{finding.basis}</p></div>
    <small>Правило {finding.code} · {finding.ruleVersion}</small>
  </aside>
}

function ToolPage({ mode }) {
  const canProcessFiles = Boolean(globalThis.crypto?.subtle)
  const fileInput = useRef(null)
  const comparisonInput = useRef(null)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const [dragging, setDragging] = useState(false)
  const [fileInfo, setFileInfo] = useState(null)
  const [parsed, setParsed] = useState(null)
  const [config, setConfig] = useState(null)
  const [comparisonParsed, setComparisonParsed] = useState(null)
  const [comparisonFile, setComparisonFile] = useState(null)
  const [comparisonConfig, setComparisonConfig] = useState(null)
  const [scopeConfirmed, setScopeConfirmed] = useState(false)
  const [result, setResult] = useState(null)
  const [filter, setFilter] = useState('all')
  const [selectedIssue, setSelectedIssue] = useState(null)

  const filteredIssues = useMemo(() => result?.issues.map((finding, index) => ({ finding, index })).filter(({ finding }) => filter === 'all' || finding.severity === filter) || [], [result, filter])
  const activeIssueIndex = filteredIssues.find(({ index }) => index === selectedIssue)?.index ?? filteredIssues[0]?.index ?? null
  const selectedFinding = activeIssueIndex === null ? null : result.issues[activeIssueIndex]

  async function loadFile(file) {
    if (!file) return
    setError(''); setResult(null); setSelectedIssue(null); setBusy(true)
    setComparisonParsed(null); setComparisonFile(null); setComparisonConfig(null); setScopeConfirmed(false)
    try {
      if (!/\.xlsx$/i.test(file.name)) throw new Error('Поддерживается только формат .xlsx. Файлы .xls и PDF пока не проверяются.')
      if (file.size > MAX_FILE_BYTES) throw new Error('Файл больше 20 МБ. Для первой версии выберите меньшую книгу.')
      const bytes = await file.arrayBuffer()
      const hash = await sha256(bytes)
      const opened = await openWorkbook(bytes)
      setParsed(opened)
      setFileInfo({ name: file.name, size: file.size, hash })
      setConfig(initialConfig(opened.workbook, opened.sheets))
      document.getElementById('checker')?.scrollIntoView({ behavior: 'smooth', block: 'start' })
    } catch (caught) {
      setParsed(null); setFileInfo(null); setConfig(null)
      setError(caught.message || 'Не удалось открыть файл.')
    } finally { setBusy(false) }
  }

  async function loadComparisonFile(file) {
    if (!file) return
    setError(''); setResult(null); setSelectedIssue(null); setBusy(true)
    try {
      if (!/\.xlsx$/i.test(file.name)) throw new Error('Второй файл должен быть в формате .xlsx.')
      if (file.size > MAX_FILE_BYTES) throw new Error('Второй файл больше 20 МБ.')
      const bytes = await file.arrayBuffer()
      const hash = await sha256(bytes)
      const opened = await openWorkbook(bytes)
      setComparisonParsed(opened)
      setComparisonFile({ name: file.name, size: file.size, hash })
      setComparisonConfig(initialConfig(opened.workbook, opened.sheets).sources)
      setScopeConfirmed(false)
    } catch (caught) {
      setComparisonParsed(null); setComparisonFile(null); setComparisonConfig(null); setScopeConfirmed(false)
      setError(caught.message || 'Не удалось открыть второй файл.')
    } finally { setBusy(false) }
  }

  function removeComparisonFile() {
    setComparisonParsed(null); setComparisonFile(null); setComparisonConfig(null); setScopeConfirmed(false); setResult(null)
    if (comparisonInput.current) comparisonInput.current.value = ''
  }

  async function openExample() {
    setBusy(true); setError('')
    try { await loadFile(await createExampleFile()) }
    catch { setError('Не удалось создать учебный пример.'); setBusy(false) }
  }

  function reset() {
    setParsed(null); setFileInfo(null); setConfig(null); setResult(null); setError(''); setSelectedIssue(null); setFilter('all')
    setComparisonParsed(null); setComparisonFile(null); setComparisonConfig(null); setScopeConfirmed(false)
    if (fileInput.current) fileInput.current.value = ''
    if (comparisonInput.current) comparisonInput.current.value = ''
  }

  function runValidation() {
    setError('')
    try {
      const chosen = [config.sources.sheet, config.emissions.sheet].filter(Boolean)
      if (chosen.some(name => parsed.sheets.find(sheet => sheet.name === name)?.overLimit)) throw new Error('Один из выбранных листов превышает 10 000 строк.')
      const validated = validateWorkbook(parsed.workbook, config, fileInfo.hash)
      let issues = validated.issues.map(finding => withFile(finding, fileInfo))
      let comparison = null
      let secondaryValidation = null
      if (mode === 'compare') {
        if (!comparisonParsed) throw new Error('Добавьте второй XLSX для сравнения.')
        if (!scopeConfirmed) throw new Error('Подтвердите общий объект и сопоставимое состояние двух файлов.')
        if (comparisonParsed.sheets.find(sheet => sheet.name === comparisonConfig.sheet)?.overLimit) throw new Error('Выбранный лист второго файла превышает 10 000 строк.')
        secondaryValidation = validateWorkbook(comparisonParsed.workbook, { sources: comparisonConfig }, comparisonFile.hash)
        comparison = compareSourceInventories(validated, secondaryValidation, {
          primaryFile: fileInfo, secondaryFile: comparisonFile, scopeConfirmed,
        })
        issues = [...issues, ...secondaryValidation.issues.map(finding => withFile(finding, comparisonFile)), ...comparison.findings]
      }
      const complete = {
        ...validated,
        files: comparisonFile ? [fileInfo, comparisonFile] : [fileInfo],
        issues,
        summary: { ...validated.summary, issues: issues.length, secondarySources: secondaryValidation?.sources.length ?? null },
        comparison,
        secondaryValidation,
      }
      setResult(complete); setSelectedIssue(complete.issues.length ? 0 : null); setFilter('all')
      window.setTimeout(() => document.getElementById('results')?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 80)
    } catch (caught) { setError(caught.message || 'Проверка не завершена.') }
  }


  const isCompare = mode === 'compare'
  return <>
    <SiteHeader root="../../" active={mode} />
    <main>
      <section className="tool-intro"><div className="container tool-intro-inner">
        <div><a className="breadcrumb" href="../../">← Все инструменты</a><span className="eyebrow">Экология / {isCompare ? 'Сопоставление' : 'Структурная проверка'}</span>
          <h1>{isCompare ? 'Сравните два файла источников.' : 'Проверьте таблицу до сдачи.'}</h1>
          <p>{isCompare ? 'Найдите расхождения номеров и координат между двумя XLSX. Сначала подтвердите, что данные относятся к одному объекту и сопоставимому состоянию.' : 'Откройте XLSX-инвентаризацию, укажите колонки и получите замечания с адресами ячеек. Проверка выполняется локально в браузере.'}</p>
        </div>
        <div className="tool-facts"><div><LockKeyhole size={19}/><span>Файлы остаются в браузере</span></div><div><MapPin size={19}/><span>Замечания с адресом ячейки</span></div><div><FileText size={19}/><span>Отчёт CSV и JSON</span></div></div>
      </div></section>
      <section className="checker-section section" id="checker"><div className="container">
        <div className="checker-title"><div><span className="eyebrow">{isCompare ? '01 / Первый файл' : 'Рабочая область'}</span><h2>{isCompare ? 'Подготовьте данные' : 'Загрузите книгу'}</h2><p>{isCompare ? 'Загрузите первую книгу, затем добавьте вторую и подтвердите сопоставимость. Лист выбросов в первой книге можно проверить дополнительно.' : 'Начните с листа источников. Лист выбросов можно добавить для проверки связей.'}</p></div><span className="beta-label"><span></span> Ранняя версия</span></div>
        <div className="checker-shell">
        <div className="checker-topline"><div><span className="live-dot"></span> Локальный режим</div><span>Поддерживается .xlsx · до 20 МБ</span></div>
        {!canProcessFiles && <div className="error-banner" role="alert"><CircleAlert size={18} /><span>Обработка файлов временно недоступна: для неё требуется HTTPS. Защищённое соединение для enormativ.ru ещё настраивается.</span></div>}
        {!parsed && canProcessFiles && <div className={`dropzone ${dragging ? 'dragging' : ''}`} onDragOver={event => { event.preventDefault(); setDragging(true) }} onDragLeave={event => { event.preventDefault(); setDragging(false) }} onDrop={event => { event.preventDefault(); setDragging(false); loadFile(event.dataTransfer.files[0]) }}><div className="drop-icon"><UploadCloud size={29} /></div><h3>Перетащите книгу сюда</h3><p>или выберите XLSX с устройства. Данные обрабатываются только в памяти этой вкладки.</p><input ref={fileInput} type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onChange={event => loadFile(event.target.files[0])} aria-label="Выбрать XLSX файл" /><div className="drop-actions"><button className="button button-primary" onClick={() => fileInput.current?.click()} disabled={busy}>{busy ? 'Открываем файл…' : 'Выбрать XLSX'} <ArrowRight size={17} /></button><button className="button button-light" onClick={openExample} disabled={busy}>Открыть учебный пример</button></div><small>Учебный пример синтетический; он не подтверждает качество на реальных книгах.</small></div>}
        {parsed && config && <div className="file-workspace"><div className="file-summary"><div className="file-icon"><FileSpreadsheet size={24} /></div><div><strong>{fileInfo.name}</strong><span>{parsed.sheets.length} {parsed.sheets.length % 10 === 1 && parsed.sheets.length % 100 !== 11 ? 'лист' : parsed.sheets.length % 10 >= 2 && parsed.sheets.length % 10 <= 4 && (parsed.sheets.length % 100 < 12 || parsed.sheets.length % 100 > 14) ? 'листа' : 'листов'} · {(fileInfo.size / 1024 / 1024).toFixed(2)} МБ · SHA-256 {fileInfo.hash.slice(0, 12)}…</span></div><button className="icon-button" onClick={reset} aria-label="Удалить файл из памяти" title="Удалить файл из памяти"><X size={18} /></button></div><div className="mapping-layout"><SheetMapping type="sources" title="Таблица источников" description="Одна строка — один источник. Повтор номера будет отмечен как замечание." fields={SOURCE_FIELDS} sheets={parsed.sheets} workbook={parsed.workbook} config={config.sources} onChange={sources => { setConfig({ ...config, sources }); setResult(null) }} /><SheetMapping type="emissions" title="Таблица выбросов" description="Свяжем вещества с источниками по номеру, без пересчёта величин." fields={EMISSION_FIELDS} sheets={parsed.sheets} workbook={parsed.workbook} config={config.emissions} onChange={emissions => { setConfig({ ...config, emissions }); setResult(null) }} optional /></div>{mode === 'compare' && <ComparisonPanel inputRef={comparisonInput} parsed={comparisonParsed} file={comparisonFile} config={comparisonConfig} confirmed={scopeConfirmed} busy={busy} onLoad={loadComparisonFile} onRemove={removeComparisonFile} onMapping={value => { setComparisonConfig(value); setResult(null) }} onConfirm={value => { setScopeConfirmed(value); setResult(null) }} />}<div className="run-row"><p><LockKeyhole size={16} /> Маппинг и файл остаются только в памяти вкладки. Ничего не сохраняется автоматически.</p><button className="button button-primary" onClick={runValidation}>{mode === 'compare' ? 'Сравнить два файла' : 'Проверить данные'} <ArrowRight size={18} /></button></div></div>}
        {error && <div className="error-banner" role="alert"><CircleAlert size={18} /><span>{error}</span><button onClick={() => setError('')} aria-label="Закрыть сообщение"><X size={16} /></button></div>}

        </div>
      </div></section>
      {result && <section className="results-section section" id="results"><div className="container"><div className="results-heading"><div><span className="eyebrow">Результат проверки</span><h2>{result.issues.length ? 'Есть что проверить' : (result.comparison?.summary.unknown || result.comparison?.summary.notApplicable) ? 'Есть что уточнить' : 'Структурных замечаний нет'}</h2><p>{result.issues.length ? 'Откройте замечание, чтобы увидеть исходную ячейку, основание и действие.' : 'Это не подтверждает нормативную корректность документа. Проверьте маппинг и исходные данные.'}</p></div><div className="export-actions"><button className="button button-outline" onClick={() => downloadBlob('enormativ-issues.csv', issueCsv(result), 'text/csv;charset=utf-8')}><Download size={17} /> Отчёт CSV</button><button className="button button-outline" onClick={() => downloadBlob('enormativ-result.json', JSON.stringify(result, null, 2), 'application/json;charset=utf-8')}><Download size={17} /> Данные JSON</button></div></div>{result.comparison && <div className="comparison-summary" role="status"><strong>Сверка двух файлов:</strong> {result.comparison.summary.sharedIds} общих номеров, {result.comparison.summary.issues} расхождений, {result.comparison.summary.unknown} результатов «неизвестно», {result.comparison.summary.notApplicable} неприменимых проверок. Неизвестное не считается успешной проверкой.</div>}{result.comparison?.evaluations.some(item => item.status === 'unknown' || item.status === 'not-applicable') && <details className="comparison-uncertain"><summary>Показать источники, которые требуют уточнения или не проверялись</summary><ul>{result.comparison.evaluations.filter(item => item.status === 'unknown' || item.status === 'not-applicable').map((item, index) => <li key={`${item.code}-${item.entity}-${index}`}><strong>Источник {item.entity} · {item.code === 'SOURCE_PRESENCE' ? 'сопоставление номеров' : 'координаты'} · {item.status === 'unknown' ? 'нужно уточнить' : 'не проверялось'}</strong><span>{item.reason}</span><span>{item.evidence.map(place => `${place.fileName} · ${place.sheet} · ${place.cell}`).join(' / ')}</span></li>)}</ul></details>}<div className="results-summary"><div><strong>{result.summary.sources}</strong><span>источников</span></div><div><strong>{result.summary.emissions}</strong><span>строк выбросов</span></div><div><strong>{result.issues.filter(item => item.severity === 'error').length}</strong><span>ошибок</span></div><div><strong>{result.summary.issues}</strong><span>замечаний всего</span></div></div><div className="results-grid"><div className="issue-panel"><div className="filter-row">{[['all', 'Все'], ['error', 'Ошибки'], ['warning', 'Проверить'], ['info', 'Уточнить']].map(([value, label]) => <button key={value} className={filter === value ? 'active' : ''} onClick={() => setFilter(value)}>{label}{value !== 'all' && <span>{result.issues.filter(item => item.severity === value).length}</span>}</button>)}</div><div className="issues-list">{filteredIssues.length ? filteredIssues.map(({ finding, index }) => <ResultRow key={`${finding.code}-${index}`} finding={finding} selected={activeIssueIndex === index} onClick={() => setSelectedIssue(index)} />) : <div className="empty-results"><Check size={28} /><strong>Здесь замечаний нет</strong><p>Выберите другой фильтр или проверьте маппинг.</p></div>}</div></div><FindingDetail finding={selectedFinding} rulesetVersion={result.rulesetVersion} hasIssues={result.issues.length > 0} hasUncertain={Boolean(result.comparison?.summary.unknown || result.comparison?.summary.notApplicable)} isComparison={Boolean(result.comparison)} /></div><button className="reset-link" onClick={reset}><RotateCcw size={16} /> Удалить файлы из памяти и начать заново</button></div></section>}

      <section className="boundary-section"><div className="container boundary-inner"><div><span className="eyebrow">Границы версии 0.1</span><h2>Заключение остаётся за экологом</h2></div><p>Сервис проверяет структуру, номера, связи и часть значений. Он не оценивает полное соответствие приказу № 871, не пересчитывает выбросы, не переводит местные системы координат и не меняет исходный XLSX. {mode === 'compare' && ' Межфайловая сверка относится только к источникам, сопоставимость которых подтверждена пользователем.'}</p></div></section>
    </main>
    <SiteFooter root="../../" />
  </>
}

function PortalPage() {
  return <>
    <SiteHeader />
    <main>
      <section className="portal-hero"><div className="container portal-hero-grid">
        <div className="portal-lead"><span className="eyebrow"><span className="eyebrow-dot"></span> Цифровые инструменты для экологов</span><h1>От файла — <em>к проверяемому замечанию.</em></h1><p>еНорматив помогает заметить ошибки в рабочих таблицах и увидеть, где именно они возникли. Первый модуль работает с данными инвентаризации источников выбросов.</p><div className="portal-actions"><a className="button button-primary" href="./tools/check/">Проверить XLSX <ArrowRight size={18}/></a><a className="portal-sub-link" href="#tools">Выбрать инструмент <ArrowDownRight size={17}/></a></div><div className="portal-proof"><span><LockKeyhole size={15}/> Локально в браузере</span><span><Fingerprint size={15}/> След до исходной ячейки</span></div></div>
        <div className="portal-visual"><img src="./hero-ecology.png" alt="Иллюстрация промышленной площадки с источниками выбросов" /><div className="visual-caption"><span>01 / ЭКОЛОГИЯ</span><strong>Рабочие данные<br/>под контролем</strong><small>XLSX → правило → замечание</small></div></div>
      </div></section>
      <section className="portal-tools" id="tools"><div className="container"><div className="portal-section-heading"><div><span className="eyebrow">Инструменты / 01—02</span><h2>Начните с вашей задачи.</h2></div><p>Оба инструмента работают с файлами на вашем устройстве. Они показывают основание проверки и сохраняют результат для дальнейшей работы.</p></div>
        <div className="portal-card-grid">
          <a className="portal-card portal-card-primary" href="./tools/check/"><div className="portal-card-top"><span>01 / Один файл</span><ScanSearch size={30}/></div><div><h3>Проверить<br/>XLSX</h3><p>Найдите пропуски, дубли номеров, ошибки связей и числовых полей в книге инвентаризации.</p></div><div className="portal-card-bottom"><span>.xlsx · до 20 МБ</span><strong>Открыть инструмент <ArrowRight size={18}/></strong></div></a>
          <a className="portal-card portal-card-secondary" href="./tools/compare/"><div className="portal-card-top"><span>02 / Два файла</span><FileSpreadsheet size={30}/></div><div><h3>Сравнить<br/>источники</h3><p>Сопоставьте номера и координаты двух XLSX, чтобы увидеть расхождения и требующие проверки случаи.</p></div><div className="portal-card-bottom"><span>Общий объект · подтверждённая область</span><strong>Открыть инструмент <ArrowRight size={18}/></strong></div></a>
        </div>
      </div></section>
      <section className="portal-method"><div className="container portal-method-grid"><div><span className="eyebrow">Подход</span><h2>Понятно, что проверено.<br/>Видно, что требует решения.</h2></div><div><p>Каждое замечание связано с файлом, листом, ячейкой и версией структурного правила. Отчёт можно выгрузить в CSV или JSON.</p><p>Сервис не выдаёт нормативное заключение: профессиональная оценка и работа с актуальной редакцией требований остаются за специалистом.</p></div></div></section>
    </main>
    <SiteFooter />
  </>
}

function App() {
  const page = document.body.dataset.page
  return page === 'check' || page === 'compare' ? <ToolPage mode={page} /> : <PortalPage />
}

createRoot(document.getElementById('root')).render(<App />)
