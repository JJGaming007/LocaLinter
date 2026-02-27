const fs = require('fs');
const XLSX = require('xlsx');
const workbook = XLSX.readFile('I2Loc Indus BR Localization Final - Production.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
fs.writeFileSync('out.txt', JSON.stringify({
    sheetName,
    headers: data[0],
    row1: data[1],
    row2: data[2]
}, null, 2));
