"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
    return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.parseXlsx = parseXlsx;
exports.readXlsxFromFile = readXlsxFromFile;
var jszip_1 = require("jszip");
var fast_xml_parser_1 = require("fast-xml-parser");
var fs = require("fs");
var parser = new fast_xml_parser_1.XMLParser();
function parseXlsx(buffer, sheetName) {
    return __awaiter(this, void 0, void 0, function () {
        var zip, workbookXml, workbook, sheetNodes, selectedSheet, relId, relsXml, rels, relList, rel, sheetPath, sheetXml, sharedStrings, sharedXml, sharedObj, siNodes, _i, siNodes_1, si, value, sheetObj, rows, result, headers, i, row, cells, rowData, j, cell, rawValue, value, key;
        var _a, _b, _c, _d, _e, _f, _g, _h;
        return __generator(this, function (_j) {
            switch (_j.label) {
                case 0: return [4 /*yield*/, jszip_1.default.loadAsync(buffer)];
                case 1:
                    zip = _j.sent();
                    return [4 /*yield*/, ((_a = zip.file("xl/workbook.xml")) === null || _a === void 0 ? void 0 : _a.async("string"))];
                case 2:
                    workbookXml = _j.sent();
                    if (!workbookXml)
                        throw new Error("Missing workbook.xml");
                    workbook = parser.parse(workbookXml);
                    sheetNodes = Array.isArray(workbook.workbook.sheets.sheet)
                        ? workbook.workbook.sheets.sheet
                        : [workbook.workbook.sheets.sheet];
                    selectedSheet = sheetName
                        ? sheetNodes.find(function (s) { return s["@_name"] === sheetName; })
                        : sheetNodes[0];
                    if (!selectedSheet)
                        throw new Error("Sheet \"".concat(sheetName, "\" not found"));
                    relId = selectedSheet["@_r:id"];
                    return [4 /*yield*/, ((_b = zip.file("xl/_rels/workbook.xml.rels")) === null || _b === void 0 ? void 0 : _b.async("string"))];
                case 3:
                    relsXml = _j.sent();
                    if (!relsXml)
                        throw new Error("Missing workbook.xml.rels");
                    rels = parser.parse(relsXml);
                    relList = Array.isArray(rels.Relationships.Relationship)
                        ? rels.Relationships.Relationship
                        : [rels.Relationships.Relationship];
                    rel = relList.find(function (r) { return r["@_Id"] === relId; });
                    if (!rel)
                        throw new Error("Relationship \"".concat(relId, "\" not found"));
                    sheetPath = "xl/".concat(rel["@_Target"]);
                    return [4 /*yield*/, ((_c = zip.file(sheetPath)) === null || _c === void 0 ? void 0 : _c.async("string"))];
                case 4:
                    sheetXml = _j.sent();
                    if (!sheetXml)
                        throw new Error("Sheet XML not found at ".concat(sheetPath));
                    sharedStrings = [];
                    return [4 /*yield*/, ((_d = zip.file("xl/sharedStrings.xml")) === null || _d === void 0 ? void 0 : _d.async("string"))];
                case 5:
                    sharedXml = _j.sent();
                    if (sharedXml) {
                        sharedObj = parser.parse(sharedXml);
                        siNodes = Array.isArray((_e = sharedObj.sst) === null || _e === void 0 ? void 0 : _e.si)
                            ? sharedObj.sst.si
                            : [(_f = sharedObj.sst) === null || _f === void 0 ? void 0 : _f.si];
                        for (_i = 0, siNodes_1 = siNodes; _i < siNodes_1.length; _i++) {
                            si = siNodes_1[_i];
                            value = typeof (si === null || si === void 0 ? void 0 : si.t) === "string" ? si.t : ((_g = si === null || si === void 0 ? void 0 : si.t) === null || _g === void 0 ? void 0 : _g["#text"]) || "";
                            sharedStrings.push(value);
                        }
                    }
                    sheetObj = parser.parse(sheetXml);
                    rows = sheetObj.worksheet.sheetData.row;
                    if (!rows)
                        return [2 /*return*/, []];
                    result = [];
                    headers = [];
                    for (i = 0; i < rows.length; i++) {
                        row = rows[i];
                        cells = Array.isArray(row.c) ? row.c : [row.c];
                        rowData = {};
                        for (j = 0; j < cells.length; j++) {
                            cell = cells[j];
                            rawValue = (_h = cell.v) !== null && _h !== void 0 ? _h : "";
                            value = cell.t === "s" ? sharedStrings[Number(rawValue)] : rawValue;
                            if (i === 0) {
                                headers.push(value);
                            }
                            else {
                                key = headers[j] || "Column".concat(j + 1);
                                rowData[key] = value;
                            }
                        }
                        if (i !== 0)
                            result.push(rowData);
                    }
                    return [2 /*return*/, result];
            }
        });
    });
}
function readXlsxFromFile(filePath, sheetName) {
    return __awaiter(this, void 0, void 0, function () {
        var buffer;
        return __generator(this, function (_a) {
            buffer = fs.readFileSync(filePath);
            return [2 /*return*/, parseXlsx(buffer, sheetName)];
        });
    });
}

