import xlsx from "xlsx";
import fs from "fs";

const INPUT_FILE_PATH = "./data/input.xlsx";
const OUTPUT_FILE_PATH = "./data/output.xlsx";
const OUTPUT_FILE_SHEET_NAME = "Sheet1";

// Основная функция, содержащая всю функциональность.
async function main() {
  // В этой части мы читаем файл и получаем данные.
  const buffer = readFile();
  const data = await parseExcelData(buffer); // Получить массив объектов
  let obj_id = {};
  let obj_idarr = {};
  // Просматриваем все строки файла Excel
  // Рядок назвал "row"
  for (let i = 0; i < data.length; i++) {
    let sku = data[i]["Код_товара"];
    let color = data[i]["Значение_Характеристики_1"];
    let id = `${sku}+${color}`;

    let tovari = obj_id[id];
    if (!tovari) {
      let arr_id = [data[i]];
      obj_id[id] = arr_id;
    } else {
      tovari.push(data[i]);
      obj_id[id] = tovari;
    }

    let model_ids = obj_idarr[id];
    if (!model_ids) {
      let arr_id = [data[i]["ID_группы_разновидностей"]];
      obj_idarr[id] = arr_id;
    } else {
      model_ids.push(data[i]["ID_группы_разновидностей"]);
      const uniqueIds = Array.from(new Set(model_ids));
      obj_idarr[id] = uniqueIds;
    }
  }

  for (let [k, v] of Object.entries(obj_idarr)) {
    if (v.length > 1) {
      console.log(`k: ${k}, v: ${v}`);
    }
  }

  // Сохранить данные в файле «output.xlsx».
  writeExcelFile(data);
}

// Вызов основной функции
main();

/// --- Утилиты ---

// Эта функция считывает двоичные данные файла и возвращает буфер.
function readFile() {
  const buffer = fs.readFileSync(INPUT_FILE_PATH);

  return buffer;
}

// Получение данных Excel с помощью библиотеки «xlsx».
async function parseExcelData(buffer) {
  const workbook = xlsx.read(buffer, { type: "buffer" }); // Получить данные рабочей тетради Excel
  const sheetName = workbook.SheetNames[0]; // Получить название первого листа в рабочей книге
  const sheet = workbook.Sheets[sheetName]; // Получить первый лист в рабочей тетради
  const jsonData = xlsx.utils.sheet_to_json(sheet); // преобразовать данные в json

  return jsonData;
}

// Записываем наш массив объектов в файл с помощью библиотеки «xlsx».
function writeExcelFile(data) {
  const worksheet = xlsx.utils.json_to_sheet(data);
  const workbook = xlsx.utils.book_new();

  xlsx.utils.book_append_sheet(workbook, worksheet, OUTPUT_FILE_SHEET_NAME);
  xlsx.writeFile(workbook, OUTPUT_FILE_PATH);
}
