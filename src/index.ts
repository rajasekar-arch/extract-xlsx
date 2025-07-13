// src/lib/parseXlsx.ts
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import * as fs from "fs";

export type SheetData = Record<string, any>;

/**
 * Parses an XLSX buffer and returns sheet rows.
 */
export async function parseXlsx(
  buffer: Buffer,
  sheetName?: string
): Promise<any[][]> {
  const zip = await JSZip.loadAsync(buffer);
  const parser = new XMLParser({ ignoreAttributes: false });

  const workbookXml = await zip.file("xl/workbook.xml")?.async("text");
  if (!workbookXml) throw new Error("Missing workbook.xml in XLSX file");

  const workbook = parser.parse(workbookXml);
  const sheets = workbook?.workbook?.sheets?.sheet;

  if (!sheets) throw new Error("No sheets found in workbook.xml");

  const sheetList = Array.isArray(sheets) ? sheets : [sheets];

  const targetSheet = sheetList.find((s: any) =>
    sheetName ? s["@_name"] === sheetName : true
  );

  if (!targetSheet) throw new Error(`Sheet "${sheetName}" not found`);

  const rId = targetSheet["@_r:id"];
  if (!rId) throw new Error(`No r:id found for sheet "${sheetName}"`);

  const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("text");
  if (!relsXml) throw new Error("Missing workbook.xml.rels in XLSX");

  const rels = parser.parse(relsXml);
  const relationships = rels.Relationships?.Relationship;
  const relList = Array.isArray(relationships)
    ? relationships
    : [relationships];

  const rel = relList.find((r: any) => r["@_Id"] === rId);
  if (!rel) throw new Error(`Relation not found for r:id "${rId}"`);

  const sheetPath = `xl/${rel["@_Target"]}`;
  const sheetXml = await zip.file(sheetPath)?.async("text");
  if (!sheetXml) throw new Error(`Sheet file not found at "${sheetPath}"`);

  const sheetData = parser.parse(sheetXml);
  const rows = sheetData?.worksheet?.sheetData?.row;

  if (!rows) return [];

  const rowList = Array.isArray(rows) ? rows : [rows];

  const result: any[][] = rowList.map((row: any) => {
    const cells = Array.isArray(row.c) ? row.c : [row.c];
    return cells.map((cell: any) => cell?.v ?? ""); // return value or empty string
  });

  return result;
}

/**
 * Returns sheet names from XLSX buffer.
 */
export async function getSheetNames(buffer: Buffer): Promise<string[]> {
  const zip = await JSZip.loadAsync(buffer);
  const workbookXml = await zip.file("xl/workbook.xml")?.async("text");

  if (!workbookXml) throw new Error("Invalid XLSX: Missing workbook.xml");

  const parser = new XMLParser({ ignoreAttributes: false });
  const parsed = parser.parse(workbookXml);

  const sheets = parsed?.workbook?.sheets?.sheet;
  const sheetArray = Array.isArray(sheets) ? sheets : [sheets];

  return sheetArray.map((sheet: any) => sheet["@_name"]).filter(Boolean);
}

/**
 * Reads an XLSX file from disk and returns parsed data from the specified sheet.
 */
export async function readXlsxFromFile(
  filePath: string,
  sheetName?: string
): Promise<SheetData[]> {
  const buffer = fs.readFileSync(filePath);
  return parseXlsx(buffer, sheetName);
}
