const express = require('express');
const path = require('path');
const xlsx = require('xlsx');
const app = express();
const port = 3000;

app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Load the Excel file
const excelFileName = 'data.xlsx';
const excelFilePath = path.join(__dirname, excelFileName);

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.post('/submit', (req, res) => {
  const { name, age } = req.body;

  // Read existing data from Excel file
  let workbook;
  try {
    workbook = xlsx.readFile(excelFilePath);
  } catch (error) {
    // If the file doesn't exist, create a new workbook
    workbook = xlsx.utils.book_new();
  }

  const sheetName = 'Sheet1';

  // Check if the sheet already exists
  if (workbook.SheetNames.includes(sheetName)) {
    // If it exists, use it
    const sheet = workbook.Sheets[sheetName];
    // Append new data to the sheet
    xlsx.utils.sheet_add_json(sheet, [{ Name: name, Age: age }], { skipHeader: true, origin: -1 });
  } else {
    // If it doesn't exist, create a new sheet
    const sheet = xlsx.utils.json_to_sheet([{ Name: name, Age: age }]);
    xlsx.utils.book_append_sheet(workbook, sheet, sheetName);
  }

  // Write the updated workbook to the file
  xlsx.writeFile(workbook, excelFilePath);

  res.redirect('/');
});

app.get('/data', (req, res) => {
  // Read data from Excel file
  const workbook = xlsx.readFile(excelFilePath);
  const sheetName = 'Sheet1';
  const sheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(sheet);

  res.json(data);
});

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
