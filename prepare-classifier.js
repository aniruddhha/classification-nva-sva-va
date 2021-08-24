const XLSX = require('xlsx')
const wb = XLSX.readFile(`classify.xlsx`)

wb.SheetNames.forEach((nm, index) => {
    console.log(nm, index)
    const tab = wb.Sheets[nm]
    const jsonTab = XLSX.utils.sheet_to_json(tab)
    console.log(jsonTab)
})