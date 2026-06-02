// Offline smoke test: extract a sample PDF and build both outputs, asserting
// the bytes are well-formed Office Open XML (a ZIP starting with "PK").
//
//   node scripts/smoke.mjs [path-to.pdf]
//
// Exits 0 on success, 1 on any failure.
import { readFile, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { extractPdf } from "../lib/extract.js";
import { buildWord } from "../lib/to-word.js";
import { buildExcel } from "../lib/to-excel.js";
import { convertWordToPdf } from "../lib/word-to-pdf.js";
import { convertExcelToPdf } from "../lib/excel-to-pdf.js";

const path =
  process.argv[2] || fileURLToPath(new URL("../sample.pdf", import.meta.url));
const TMP = tmpdir();

function isZip(buf) {
  return buf.length > 4 && buf[0] === 0x50 && buf[1] === 0x4b; // "PK"
}

function isPdf(buf) {
  // "%PDF"
  return buf.length > 4 && buf[0] === 0x25 && buf[1] === 0x50 && buf[2] === 0x44 && buf[3] === 0x46;
}

try {
  const bytes = await readFile(path);
  console.log(`→ read ${path} (${bytes.length} bytes)`);

  const extracted = await extractPdf(bytes);
  console.log(`→ pages: ${extracted.pages.length}, chars: ${extracted.charCount}`);
  const preview = extracted.pages[0]?.items.slice(0, 6).map((i) => i.str).join(" | ");
  console.log(`→ first items: ${preview}`);

  const docx = await buildWord(extracted);
  const xlsx = await buildExcel(extracted);

  const problems = [];
  if (extracted.charCount < 1) problems.push("no text extracted");
  if (!isZip(docx)) problems.push("docx is not a valid OOXML zip");
  if (!isZip(xlsx)) problems.push("xlsx is not a valid OOXML zip");
  if (docx.length < 200) problems.push("docx suspiciously small");
  if (xlsx.length < 200) problems.push("xlsx suspiciously small");

  await writeFile(join(TMP, "smoke.docx"), docx).catch(() => {});
  await writeFile(join(TMP, "smoke.xlsx"), xlsx).catch(() => {});

  // ── Round-trip: Office → PDF ───────────────────────────────────────────
  const pdfFromWord = await convertWordToPdf(docx);
  const pdfFromExcel = await convertExcelToPdf(xlsx);
  await writeFile(join(TMP, "smoke-from-word.pdf"), pdfFromWord).catch(() => {});
  await writeFile(join(TMP, "smoke-from-excel.pdf"), pdfFromExcel).catch(() => {});

  if (!isPdf(pdfFromWord)) problems.push("word-to-pdf output is not a PDF");
  if (!isPdf(pdfFromExcel)) problems.push("excel-to-pdf output is not a PDF");
  if (pdfFromWord.length < 200) problems.push("word-to-pdf output suspiciously small");
  if (pdfFromExcel.length < 200) problems.push("excel-to-pdf output suspiciously small");

  console.log(`→ word-to-pdf ${pdfFromWord.length}B, excel-to-pdf ${pdfFromExcel.length}B`);

  if (problems.length) {
    console.error("✗ SMOKE FAILED:");
    for (const p of problems) console.error(`  - ${p}`);
    process.exit(1);
  }
  console.log(`✓ SMOKE OK — docx ${docx.length}B, xlsx ${xlsx.length}B`);
  process.exit(0);
} catch (err) {
  console.error("✗ SMOKE ERROR:", err);
  process.exit(1);
}
