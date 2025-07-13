import * as fs from "fs";
import { parseXlsx } from "./shared.js";

export async function readXlsxFromFile(filePath: string, sheetName?: string) {
  const buffer = fs.readFileSync(filePath);
  return parseXlsx(buffer, sheetName);
}

export { parseXlsx } from "./shared.js";
export { getSheetNames } from "./shared.js";
