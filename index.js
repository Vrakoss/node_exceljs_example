
const express = require('express');
const app = express();
const port = 5000;
const hostname = '127.0.0.1';
// Exceljs
const Exceljs = require('exceljs');
var workbook = null;

app.get('/', (req, res) => {
    res.sendFile('index.html', {root: __dirname});
});

app.listen(port, hostname, () => {
  console.log(`Server running at http://${hostname}:${port}/`)
});

const setupWorkbook = () => {
  workbook = new Exceljs.Workbook();
  workbook.creator = 'Vrakoss';
  workbook.lastModifiedBy = 'Vrakoss';
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.lastPrinted = new Date();
};

const setupWorksheet = () => {
  const worksheet = workbook.addWorksheet('Test numeric values');
  worksheet.columns = [
    { header: 'Numeric values', key: 'values1', width: 15 },
    { header: 'Numeric values', key: 'values2', width: 15 },
    { header: 'Calculation result', key: 'result', width: 15 },
  ];
};

const fillWorkSheet = () => {
  const worksheet = workbook.getWorksheet('Test numeric values');
  const column1 = worksheet.getColumn('values1');
  const column2 = worksheet.getColumn('values2');
  const column3 = worksheet.getColumn('result');
  const column1Values = [1, 2, -3, 4, -5, 6];
  const column2Values = [1, -2, 3, -4, 5, 6];
  column1.values = column1.values.concat(column1Values);
  column2.values = column2.values.concat(column2Values);
  // fill column3 values with the formula
  column3.values = column3.values.concat(column1Values.map((e, i) => getFormulaCellValue(`A${i+2}`, `B${i+2}`, column1Values[i] + column2Values[i])));
};

const getFormulaCellValue = (cellA, cellB, result) => {
  return {
    formula: `=SUM(${cellA},${cellB})`,
    result: result,
    date1904: false,
  };
};

app.get('/download', (req, res) => {
  setupWorkbook();
  setupWorksheet();
  fillWorkSheet();
  // set Header
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=' + 'test.xlsx');
  // send workbook
  workbook.xlsx.write(res).then(() => {
    res.status(200).end();
  });
});