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

# ✅ Sample Usage in Browser (React or Vanilla JS)
```typescript
import { parseXlsx, getSheetNames } from "extract-xlsx";
import { Buffer } from "buffer"; // Need to be installed // npm install buffer
(window as any).Buffer = Buffer; // declare after installed

document.querySelector("#upload")?.addEventListener("change", async (e: any) => {
  const file = e.target.files[0];
  const arrayBuffer = await file.arrayBuffer();
  const buffer = Buffer.from(arrayBuffer); // If you're using `buffer` polyfill in browser

  const sheetNames = await getSheetNames(buffer);
  console.log("Sheets:", sheetNames);

  const data = await parseXlsx(buffer, sheetNames[0]);
  console.log("Data:", data);
});


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

Formula and styling not supported

Returns plain data only

# 📚 Roadmap


Multi-sheet parser

.xls (legacy) format support

Write/Export .xlsx (optional)


# 🧑‍💻 Author
Created by [RAJASEKAR E C <rajasekar_e_c@outlook.com>]
GitHub: [rajasekar-arch](https://github.com/rajasekar-arch/extract-xlsx)

