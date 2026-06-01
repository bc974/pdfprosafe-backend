// Reconstruct a real cell-grid .xlsx from extracted PDF geometry.
//
// Strategy: one worksheet per page. Rows come from line grouping (shared with
// the Word builder). Columns come from clustering item x-starts across the
// whole page, so values that line up vertically land in the same column —
// which is exactly what a tabular PDF needs. Free-flowing prose simply yields
// many narrow columns; the data is still all present and editable.

import ExcelJS from "exceljs";

import { groupLines, lineToText, deriveColumns, columnIndexFor } from "./group.js";

const MAX_COLS = 64; // guard against pathological pages exploding the grid

/**
 * @param {{pages: Array}} extracted  Output of extractPdf().
 * @returns {Promise<Buffer>}  .xlsx bytes.
 */
export async function buildExcel(extracted) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "pdfprosafe.com";
  wb.created = new Date();

  extracted.pages.forEach((page, pageIdx) => {
    const ws = wb.addWorksheet(`Page ${pageIdx + 1}`);
    const lines = groupLines(page.items);
    if (lines.length === 0) return;

    let { bounds } = deriveColumns(page.items);
    // If a page is prose (hundreds of x-starts), collapse to a single column of
    // full lines rather than a meaningless wide grid.
    const tabular = bounds.length > 1 && bounds.length <= MAX_COLS;

    lines.forEach((line, r) => {
      if (!tabular) {
        ws.getCell(r + 1, 1).value = lineToText(line);
        return;
      }
      // Place each item's text into its nearest column band. Multiple items in
      // the same cell (rare) are space-joined in reading order.
      const rowCells = new Array(bounds.length).fill(null);
      for (const it of line.items) {
        const c = columnIndexFor(it.x, bounds);
        const txt = it.str.replace(/\s+/g, " ").trim();
        if (!txt) continue;
        rowCells[c] = rowCells[c] ? `${rowCells[c]} ${txt}` : txt;
      }
      rowCells.forEach((val, c) => {
        if (val != null) ws.getCell(r + 1, c + 1).value = coerce(val);
      });
    });

    // Reasonable default widths so the sheet is readable on open.
    ws.columns.forEach((col) => {
      col.width = Math.min(40, Math.max(10, col.width || 12));
    });
  });

  if (wb.worksheets.length === 0) wb.addWorksheet("Page 1");

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}

/** Turn a numeric-looking string into a real number so Excel can compute on it. */
function coerce(text) {
  // Accept 1234, 1,234.56, -12.5, 1 234,56 (FR). Reject things with letters.
  const cleaned = text.replace(/\s/g, "");
  if (/^-?\d{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?$/.test(cleaned) || /^-?\d+(?:[.,]\d+)?$/.test(cleaned)) {
    // Normalize separators: drop thousands, use '.' as decimal.
    let normalized = cleaned;
    const lastComma = normalized.lastIndexOf(",");
    const lastDot = normalized.lastIndexOf(".");
    const decimalSep = lastComma > lastDot ? "," : ".";
    normalized = normalized
      .replace(new RegExp(`\\${decimalSep === "," ? "." : ","}`, "g"), "")
      .replace(decimalSep, ".");
    const n = Number(normalized);
    if (Number.isFinite(n)) return n;
  }
  return text;
}
