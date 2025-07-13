// ✅ Sample Usage in Browser (React or Vanilla JS)

import { parseXlsx, getSheetNames } from "extract-xlsx";

document.querySelector("#upload")?.addEventListener("change", async (e: any) => {
  const file = e.target.files[0];
  const arrayBuffer = await file.arrayBuffer();
  const buffer = Buffer.from(arrayBuffer); // If you're using `buffer` polyfill in browser

  const sheetNames = await getSheetNames(buffer);
  console.log("Sheets:", sheetNames);

  const data = await parseXlsx(buffer, sheetNames[0]);
  console.log("Data:", data);
});


// ✅ Sample Usage in Node.js

import { readXlsxFromFile, getSheetNames, parseXlsx } from "extract-xlsx";

async function run() {
  try {
    const filePath = "./SheetJSTest.xlsx"; // Replace with your actual file path

    // Option 1: Read sheet names from the file
    const sheetNames = await getSheetNames(
      await Bun.file(filePath).arrayBuffer() // OR use `fs.readFileSync` if running pure Node
    );
    console.log("✅ Sheet names:", sheetNames);

    // Option 2: Read and parse data from a specific sheet using helper
    const data = await readXlsxFromFile(filePath, sheetNames[0]); // Parse first sheet
    console.log("📊 Parsed data from sheet:", data);
  } catch (err: any) {
    console.error("❌ Error:", err.message);
  }
}

run();
