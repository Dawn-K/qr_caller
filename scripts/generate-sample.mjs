import fs from "node:fs";
import path from "node:path";
import * as XLSX from "xlsx";

const root = process.cwd();
const sampleDir = path.join(root, "sample");
const seedPath = path.join(sampleDir, "sample.seed.csv");
const outPath = path.join(sampleDir, "sample.xlsx");

if (!fs.existsSync(seedPath)) {
  throw new Error(`Seed file not found: ${seedPath}`);
}

const csvBuffer = fs.readFileSync(seedPath);
const workbook = XLSX.read(csvBuffer, { type: "buffer", raw: false });
const xlsxBuffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
fs.writeFileSync(outPath, xlsxBuffer);

console.log(`Generated ${path.relative(root, outPath)} from ${path.relative(root, seedPath)}`);
