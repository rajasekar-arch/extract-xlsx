import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import * as fs from "fs";

const parser = new XMLParser();

type SheetData = Record<string, any>;

export async function parseXlsx(
    buffer: Buffer,
    sheetName?: string
): Promise<SheetData[]> {
    const zip = await JSZip.loadAsync(buffer);

    // Load and parse workbook metadata
    const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
    if (!workbookXml) throw new Error("Missing workbook.xml");

    const workbook = parser.parse(workbookXml);
    const sheetNodes = Array.isArray(workbook.workbook.sheets.sheet)
        ? workbook.workbook.sheets.sheet
        : [workbook.workbook.sheets.sheet];

    // Find sheet info (name, relId)
    const selectedSheet = sheetName
        ? sheetNodes.find((s: any) => s["@_name"] === sheetName)
        : sheetNodes[0];

    if (!selectedSheet) throw new Error(`Sheet "${sheetName}" not found`);

    const relId = selectedSheet["@_r:id"];

    // Resolve sheet path from workbook relationships
    const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
    if (!relsXml) throw new Error("Missing workbook.xml.rels");

    const rels = parser.parse(relsXml);
    const relList = Array.isArray(rels.Relationships.Relationship)
        ? rels.Relationships.Relationship
        : [rels.Relationships.Relationship];

    const rel = relList.find((r: any) => r["@_Id"] === relId);
    if (!rel) throw new Error(`Relationship "${relId}" not found`);

    const sheetPath = `xl/${rel["@_Target"]}`;
    const sheetXml = await zip.file(sheetPath)?.async("string");
    if (!sheetXml) throw new Error(`Sheet XML not found at ${sheetPath}`);

    // Parse shared strings
    const sharedStrings: string[] = [];
    const sharedXml = await zip.file("xl/sharedStrings.xml")?.async("string");

    if (sharedXml) {
        const sharedObj = parser.parse(sharedXml);
        const siNodes = Array.isArray(sharedObj.sst?.si)
            ? sharedObj.sst.si
            : [sharedObj.sst?.si];
        for (const si of siNodes) {
            const value = typeof si?.t === "string" ? si.t : si?.t?.["#text"] || "";
            sharedStrings.push(value);
        }
    }

    // Parse rows and columns
    const sheetObj = parser.parse(sheetXml);
    const rows = sheetObj.worksheet.sheetData.row;
    if (!rows) return [];

    const result: SheetData[] = [];
    let headers: string[] = [];

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const cells = Array.isArray(row.c) ? row.c : [row.c];
        const rowData: SheetData = {};

        for (let j = 0; j < cells.length; j++) {
            const cell = cells[j];
            const rawValue = cell.v ?? "";
            const value = cell.t === "s" ? sharedStrings[Number(rawValue)] : rawValue;

            if (i === 0) {
                headers.push(value);
            } else {
                const key = headers[j] || `Column${j + 1}`;
                rowData[key] = value;
            }
        }

        if (i !== 0) result.push(rowData);
    }

    return result;
}

export async function readXlsxFromFile(
    filePath: string,
    sheetName?: string
): Promise<SheetData[]> {
    const buffer = fs.readFileSync(filePath);
    return parseXlsx(buffer, sheetName);
}
