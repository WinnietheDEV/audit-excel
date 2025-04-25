const XLSX = require("xlsx");

// 1. Read the Excel file
const workbook = XLSX.readFile("gl.xlsx");
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// 2. Convert sheet to JSON
const data = XLSX.utils.sheet_to_json(worksheet);

// 3. Summarize the data
const summary = {};
data.forEach((row) => {
  const account = String(row["Account"]).trim();
  // if (account === "125210020") {
  const amount = parseFloat(row["Amount in Local Currency"]) || 0;
  console.log(amount);
  if (summary[account]) {
    summary[account] += amount;
  } else {
    summary[account] = amount;
  }
  // }
});

// 4. Convert summary object to array of objects (for writing to Excel)
const summaryArray = Object.entries(summary).map(([account, amount]) => ({
  Account: account,
  TotalAmount: amount,
}));

// 5. Create new worksheet and workbook
const newWorksheet = XLSX.utils.json_to_sheet(summaryArray);
const newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Summary");

// 6. Write to new Excel file
XLSX.writeFile(newWorkbook, "summary.xlsx");
