export async function createExampleFile() {
  const { default: ExcelJS } = await import('exceljs')
  const workbook = new ExcelJS.Workbook()
  const sources = workbook.addWorksheet('Источники')
  sources.addRow(['Номер источника', 'Наименование', 'Широта', 'Долгота', 'Система координат'])
  sources.addRow(['ИЗА-001', 'Вентиляционная труба', 55.75, 37.62, 'WGS84'])
  sources.addRow(['ИЗА-002', 'Площадка хранения', 55.76, 37.64, 'WGS84'])
  sources.addRow(['ИЗА-002', 'Дублированный номер', 95.2, 37.64, 'WGS84'])
  sources.addRow(['ИЗА-004', 'Тепловой узел', 55.77, null, 'WGS84'])

  const emissions = workbook.addWorksheet('Выбросы')
  emissions.addRow(['Номер источника', 'Вещество', 'Выброс', 'Единица измерения'])
  emissions.addRow(['ИЗА-001', 'Оксид азота', '0,42', 'г/с'])
  emissions.addRow(['ИЗА-003', 'Диоксид серы', 0.11, 'г/с'])
  emissions.addRow(['ИЗА-002', 'Пыль', -0.2, 'не указано'])
  emissions.addRow(['ИЗА-004', '', 'н/д', 'т/год'])

  const bytes = await workbook.xlsx.writeBuffer()
  return new File([bytes], 'пример-инвентаризации-синтетический.xlsx', {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
}
