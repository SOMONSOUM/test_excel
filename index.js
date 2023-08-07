const XLSX = require("xlsx");

// Read the Excel file
const workbook = XLSX.readFile("hscode.xlsx");
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

// Convert worksheet to JSON
const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

// Get the column index of 'hscode'
const columnIndex = jsonData[0].indexOf("hscode");
console.log(columnIndex);

// Modify the data in the 'hscode' column
for (let i = 1; i < jsonData.length; i++) {
  const values = jsonData[i][columnIndex].split(",").map((value) => {
    return `"${value.trim()}"`;
  });
  jsonData[i][columnIndex] = values.join(", ");
}

// Convert JSON back to worksheet
const updatedWorksheet = XLSX.utils.json_to_sheet(jsonData);

// Update the original workbook with the updated worksheet
workbook.Sheets[workbook.SheetNames[0]] = updatedWorksheet;

// Save the modified workbook to a new file
XLSX.writeFile(workbook, "./output/test.xlsx");