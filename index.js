const Excel = require('exceljs')

const file = new Excel.Workbook()
file.xlsx.readFile('./file.xlsx').then((file) => {
  const worksheet = file.getWorksheet(1)
  for (let i = 1; i < worksheet.actualRowCount; i++) {
    if (
      worksheet.getRow(i).getCell(23).value ===
      worksheet.getRow(i + 1).getCell(23).value
    ) {
      if (
        worksheet.getRow(i).getCell(22).value ===
        worksheet.getRow(i + 1).getCell(22).value
      ) {
        if (!worksheet.getRow(i).getCell(26).value) worksheet.spliceRows(i, 1)
        else if (!worksheet.getRow(i + 1).getCell(26).value)
          worksheet.spliceRows(i + 1, 1)
        else if (
          worksheet.getRow(i).getCell(26).value >
          worksheet.getRow(i + 1).getCell(26).value
        ) {
          worksheet.spliceRows(i + 1, 1)
        } else {
          worksheet.spliceRows(i, 1)
        }
      } else {
        if (
          worksheet.getRow(i).getCell(22).value >
          worksheet.getRow(i + 1).getCell(22).value
        ) {
          worksheet.spliceRows(i + 1, 1)
        } else {
          worksheet.spliceRows(i, 1)
        }
      }
      i--
    }
  }
  file.xlsx.writeFile('./new.xlsx')
})
