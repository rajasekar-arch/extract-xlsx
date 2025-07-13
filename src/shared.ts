import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";


export async function parseXlsx(buffer: Buffer, sheetName?: string): Promise<any[][]> {
  const zip = await JSZip.loadAsync(buffer);
  const parser = new XMLParser({ ignoreAttributes: false });

  const workbookXml = await zip.file("xl/workbook.xml")?.async("text");
  if (!workbookXml) throw new Error("Missing workbook.xml");

  const workbook = parser.parse(workbookXml);
  const sheets = workbook?.workbook?.sheets?.sheet;
  const sheetList = Array.isArray(sheets) ? sheets : [sheets];

  const targetSheet = sheetList.find((s: any) =>
    sheetName ? s["@_name"] === sheetName : true
  );
  if (!targetSheet) throw new Error(`Sheet "${sheetName}" not found`);

  const rId = targetSheet["@_r:id"];
  const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("text");
  if (!relsXml) throw new Error("Missing workbook.xml.rels");

  const rels = parser.parse(relsXml);
  const relationships = Array.isArray(rels.Relationships?.Relationship)
    ? rels.Relationships.Relationship
    : [rels.Relationships?.Relationship];

  const rel = relationships.find((r: any) => r["@_Id"] === rId);
  const sheetPath = `xl/${rel["@_Target"]}`;
  const sheetXml = await zip.file(sheetPath)?.async("text");
  if (!sheetXml) throw new Error(`Missing sheet XML at ${sheetPath}`);

  const sheetData = parser.parse(sheetXml);
  const rows = sheetData.worksheet?.sheetData?.row || [];

  const rowList = Array.isArray(rows) ? rows : [rows];
  return rowList.map((row: any) => {
    const cells = Array.isArray(row.c) ? row.c : [row.c];
    return cells.map((cell: any) => cell?.v ?? "");
  });
}

export async function getSheetNames(buffer: Buffer): Promise<string[]> {
  const zip = await JSZip.loadAsync(buffer);
  const workbookXml = await zip.file("xl/workbook.xml")?.async("text");
  if (!workbookXml) throw new Error("Missing workbook.xml");

  const parser = new XMLParser({ ignoreAttributes: false });
  const parsed = parser.parse(workbookXml);
  const sheets = parsed?.workbook?.sheets?.sheet;
  const sheetArray = Array.isArray(sheets) ? sheets : [sheets];

  return sheetArray.map((s: any) => s["@_name"]).filter(Boolean);
}
