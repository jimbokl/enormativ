// Spreadsheet apps can interpret imported CSV text as a formula. Keep
// workbook-derived values as text even when they start with a formula marker.
export function csvCell(value) {
  const raw = String(value ?? '')
  const safe = /^[\s\u0000-\u001f]*[=+\-@]/u.test(raw) ? `'${raw}` : raw
  return `"${safe.replaceAll('"', '""')}"`
}
