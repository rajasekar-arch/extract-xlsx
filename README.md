# 📦 extract-xlsx

A **lightweight, zero-dependency** `.xlsx` (Excel) file parser written in **TypeScript**, built as an alternative to heavier libraries like [`xlsx`](https://www.npmjs.com/package/xlsx).

It reads Excel `.xlsx` files directly by parsing XML and returns data in structured JSON format.

---

## ✨ Features

- ✅ Reads `.xlsx` files (Office Open XML format)
- ✅ Dynamically detects sheets via `workbook.xml`
- ✅ Reads any sheet by **name** or defaults to the first sheet
- ✅ Parses **shared strings**, **headers**, and rows
- ✅ Works in **Node.js** (browser support planned)
- ✅ Minimal and blazing fast

---

## 📦 Installation

```bash
npm install extract-xlsx

or

yarn add extract-xlsx
```

🚀 Usage
✅ Read from file

```javascript
import { readXlsxFromFile } from "extract-xlsx";

// Reads the first sheet by default
const data = await readXlsxFromFile("path/to/file.xlsx");

console.log(data);
// Output: [ { Name: "Alice", Age: "30" }, { Name: "Bob", Age: "25" } ]
```

# ✅ Read specific sheet

```javascript
const data = await readXlsxFromFile("path/to/file.xlsx", "Employees");
```

✅ Read from buffer (custom use)

```javascript
import { parseXlsx } from "extract-xlsx";
import fs from "fs";

const buffer = fs.readFileSync("sample.xlsx");
const result = await parseXlsx(buffer, "Sheet1");
```

# 🧪 Output Format

An array of JSON objects using the first row as headers:

```javascript
[
  { Name: "Alice", Age: "30" },
  { Name: "Bob", Age: "25" },
];
```

# 📄 API

```bash

readXlsxFromFile(filePath: string, sheetName?: string): Promise<any[]>
Reads from local file

sheetName (optional) — defaults to first sheet

parseXlsx(buffer: Buffer, sheetName?: string): Promise<any[]>
Accepts file Buffer

Parses .xlsx format into JSON

```

# 🔒 Limitations

Only supports .xlsx files (not .xls)

Only works in Node.js (not browser yet)

Formula and styling not supported

Returns plain data only

# 📚 Roadmap

Browser support via FileReader

Multi-sheet parser

.xls (legacy) format support

Write/Export .xlsx (optional)


# 🧑‍💻 Author
Created by [RAJASEKAR E C <rajasekar_e_c@outlook.com>]
GitHub: [rajasekar-arch](https://github.com/rajasekar-arch/extract-xlsx)

