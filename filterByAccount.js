const XLSX = require("xlsx");

// 1. Read the Excel file
const workbook = XLSX.readFile("gl.xlsx");
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// 2. Convert sheet to JSON (all columns)
const data = XLSX.utils.sheet_to_json(worksheet);

// 3. Filter rows where Account is '491000040'
const filteredRows = data.filter(
  (row) => String(row["Account"]) === "491000040"
);

// 4. Create new worksheet from filtered rows
const newWorksheet = XLSX.utils.json_to_sheet(filteredRows);

// 5. Create new workbook and append the sheet
const newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Filtered");

// 6. Write to a new Excel file
XLSX.writeFile(newWorkbook, "filtered_491000040.xlsx");
