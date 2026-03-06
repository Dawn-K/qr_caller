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

function parseCsv(text) {
  const normalized = text.replace(/^\uFEFF/, "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = normalized
    .split("\n")
    .filter((line) => line.trim() !== "");
  return lines.map((line) => parseCsvLine(line));
}

function parseCsvLine(line) {
  const cells = [];
  let current = "";
  let inQuotes = false;

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    const next = line[index + 1] ?? "";

    if (char === "\"") {
      if (inQuotes && next === "\"") {
        current += "\"";
        index += 1;
        continue;
      }
      inQuotes = !inQuotes;
      continue;
    }

    if (char === "," && !inQuotes) {
      cells.push(current);
      current = "";
      continue;
    }

    current += char;
  }

  cells.push(current);
  return cells;
}

const csvText = fs.readFileSync(seedPath, "utf8");
const rows = parseCsv(csvText);

if (rows.length === 0) {
  throw new Error(`Seed file is empty: ${seedPath}`);
}

const worksheet = XLSX.utils.aoa_to_sheet(rows);
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, "客户名单");
const xlsxBuffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
fs.writeFileSync(outPath, xlsxBuffer);

console.log(`Generated ${path.relative(root, outPath)} from ${path.relative(root, seedPath)}`);
