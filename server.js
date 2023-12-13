// server.js
const express = require('express');
const cors = require('cors');
const exceljs = require('exceljs');
const fs = require('fs');

const app = express();
const port = 3001;

app.use(cors());
app.use(express.json());

const excelFilePath = 'data.xlsx';

// Create Excel file if not exists
if (!fs.existsSync(excelFilePath)) {
  const workbook = new exceljs.Workbook();
  const worksheet = workbook.addWorksheet('data');
  worksheet.columns = [
    { header: 'Mood', key: 'mood' },
    { header: 'Food', key: 'food' },
    { header: 'Timestamp', key: 'timestamp' },
  ];
  workbook.xlsx.writeFile(excelFilePath);
}

app.post('/saveData', async (req, res) => {
  const { mood, food } = req.body;
  const timestamp = new Date().toLocaleString();

  try {
    const workbook = new exceljs.Workbook();

    // Reading the file
    await workbook.xlsx.readFile(excelFilePath);

    // Modifying the data
    const worksheet = workbook.getWorksheet('data');

    // Get the last row
    const lastRow = worksheet.lastRow || worksheet.getRow(1);
    const lastRowNumber = lastRow ? lastRow.number : 1;

    // Add a row with mood, food, and timestamp
    const newRow = [mood || '', food || '', timestamp];

    // Insert a row at the end
    worksheet.spliceRows(lastRowNumber + 1, 0, newRow);

    // Writing the modified data back to the file
    await workbook.xlsx.writeFile(excelFilePath);

    res.json({ success: true });
  } catch (error) {
    console.error('Error saving data:', error);
    res.status(500).json({ success: false, error: 'Error saving data' });
  }
});

app.get('/getData', async (req, res) => {
  try {
    const workbook = new exceljs.Workbook();

    // Reading the file
    await workbook.xlsx.readFile(excelFilePath);

    // Getting the data
    const worksheet = workbook.getWorksheet('data');
    const data = [];

    // Iterate through rows and push them into the data array
    worksheet.eachRow((row) => {
      data.push(row.values.slice(1)); // Exclude the first empty cell
    });

    // Extracting column names from the first row
    const columns = data[0];

    // Extracted data without the first row (column names)
    const extractedData = data.slice(1).map((row) =>
      Object.fromEntries(
        columns.map((col, index) => [col, row[index] === undefined ? null : row[index]])
      )
    );

    res.json({ success: true, data: extractedData });
  } catch (error) {
    console.error('Error fetching data:', error);
    res.status(500).json({ success: false, error: 'Error fetching data' });
  }
});


app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
});
