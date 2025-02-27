const XLSX = require("xlsx");
const fs = require("fs");
const { inputFile, outputPrefix, parts } = require("./environment");

function splitExcelFile(inputFile, outputPrefix, parts) {
  const workbook = XLSX.readFile(inputFile);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(worksheet);

  const totalRows = data.length;
  const rowsPerFile = Math.ceil(totalRows / parts);

  for (let i = 0; i < parts; i++) {
    const start = i * rowsPerFile;
    const end = start + rowsPerFile;
    const chunk = data.slice(start, end);

    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.json_to_sheet(chunk);
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Sheet1");

    const outputFileName = `${outputPrefix}${i}.xlsx`;
    XLSX.writeFile(newWorkbook, outputFileName);
    console.log(`Создан файл: ${outputFileName}`);
  }
}

if (fs.existsSync(inputFile)) {
  splitExcelFile(inputFile, outputPrefix, parts);
} else {
  console.error(`Файл ${inputFile} не найден`);
}
