const XLSX = require("xlsx");

const workbook = XLSX.readFile("export_sheet.xlsx");
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

//Column's name
const columnIndex = jsonData[0].indexOf("HS CODE 2022");

for (let i = 1; i < jsonData.length; i++) {
  const value = jsonData[i][columnIndex];
  if (typeof value === "string" && value.includes(",")) {
    const values = value.split(",").map((value) => {
      return `"${value.trim()}"`;
    });
    jsonData[i][columnIndex] = values.join(", ");
  }
}

const updatedWorksheet = XLSX.utils.json_to_sheet(jsonData);
workbook.Sheets[workbook.SheetNames[0]] = updatedWorksheet;

XLSX.writeFile(workbook, "./output/export_sheet_output.xlsx");
