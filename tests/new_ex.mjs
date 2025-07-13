import { getSheetNames, parseXlsx } from "extract-xlsx";
import * as fs from "fs";

async function run() {
  const filePath = "./SheetJSTest.xlsx";
  const buffer = fs.readFileSync(filePath);

  const sheetNames = await getSheetNames(buffer);
  console.log("✅ Sheet Names:", sheetNames);

  const data = await parseXlsx(buffer, sheetNames[0]);
  console.log("📊 Data from first sheet:", data);
}

run().catch(console.error);


// run this script with the command:
// node tests/new_ex.mjs
// ensure you have the 'extract-xlsx' package installed