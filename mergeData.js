const XLSX = require("xlsx");
const fs = require("fs");
const { outputPrefix, parts, mergedFile } = require("./environment");

function getCurrentTimestamp() {
  const now = new Date();
  const date = now.toISOString().split("T")[0];
  const time = now.toTimeString().split(" ")[0].replace(/:/g, "-");
  return `${date}_${time}`;
}

function mergeExcelFiles(inputPrefix, numFiles, outputFile) {
  let mergedData = [];
  for (let i = 0; i < numFiles; i++) {
    const fileName = `../assets/${inputPrefix}${i}.xlsx`;
    if (fs.existsSync(fileName)) {
      const workbook = XLSX.readFile(fileName);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(worksheet);
      mergedData = mergedData.concat(data);
    } else {
      console.error(`Файл ${fileName} не найден`);
    }
  }

  const newWorkbook = XLSX.utils.book_new();
  const newWorksheet = XLSX.utils.json_to_sheet(mergedData);
  XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Sheet1");
  const timestamp = getCurrentTimestamp();
  const finalOutputFile = `../assets/${outputFile}_${timestamp}.xlsx`;
  XLSX.writeFile(newWorkbook, finalOutputFile);
  console.log(`Создан объединенный файл: ${finalOutputFile}`);
}

if (outputPrefix && parts && mergedFile) {
  mergeExcelFiles(outputPrefix, parts, mergedFile);
} else {
  console.error(`Переменные не найдены`);
}
