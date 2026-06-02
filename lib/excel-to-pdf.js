// Reconstruct a PDF from an .xlsx file using exceljs and pdfkit.
// Each worksheet becomes its own page; rows are rendered as a bordered table
// grid (when ≤ MAX_TABLE_COLS columns) or as pipe-delimited text lines (for
// very wide sheets). Numbers and dates that ExcelJS already typed are formatted
// automatically. This is a "light" reconstruction — no cell styles/colours, no
// merged cells — but it produces a fully readable, compact PDF for data sheets.

import ExcelJS from "exceljs";
import PDFDocument from "pdfkit";

// A4 page dimensions (pt)
const PAGE_W = 595.28;
const PAGE_H = 841.89;
const MARGIN = 45;
const CONTENT_W = PAGE_W - 2 * MARGIN;
const CONTENT_H = PAGE_H - 2 * MARGIN;

const ROW_H = 16;
const FONT_SIZE = 7.5;
const HEADER_SIZE = 8;
const MAX_TABLE_COLS = 12; // beyond this, fall back to text lines

/**
 * Format a raw ExcelJS cell value to a display string.
 * Handles numbers, dates, rich-text, boolean, formulas.
 */
function fmt(value) {
  if (value == null) return "";
  if (value instanceof Date) return value.toLocaleDateString("fr-FR");
  if (typeof value === "object") {
    // Rich text
    if (value.richText) return value.richText.map((r) => r.text).join("");
    // Formula with result
    if (value.result != null) return String(value.result);
    // Shared-formula wrapper
    if (value.formula != null) return "";
  }
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

/**
 * @param {Buffer} buffer  Raw .xlsx bytes.
 * @returns {Promise<Buffer>}  PDF bytes.
 */
export async function convertExcelToPdf(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({
      size: "A4",
      margins: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN },
      autoFirstPage: false,
      info: { Creator: "pdfprosafe.com" },
    });
    const chunks = [];
    doc.on("data", (c) => chunks.push(c));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);

    const sheets = wb.worksheets;
    if (sheets.length === 0) {
      doc.addPage();
      doc.font("Helvetica").fontSize(11).text("(empty workbook)");
      doc.end();
      return;
    }

    for (const ws of sheets) {
      doc.addPage();

      // Sheet title
      doc.font("Helvetica-Bold").fontSize(10)
        .text(ws.name, MARGIN, MARGIN, { lineGap: 4 });
      let y = MARGIN + 18;

      // Collect rows (skip completely empty rows)
      const rawRows = [];
      ws.eachRow({ includeEmpty: false }, (row) => {
        const cells = [];
        for (let c = 1; c <= row.cellCount; c++) {
          cells.push(fmt(row.getCell(c).value));
        }
        if (cells.some((v) => v.trim())) rawRows.push(cells);
      });

      if (rawRows.length === 0) {
        doc.font("Helvetica").fontSize(FONT_SIZE).text("(empty sheet)", MARGIN, y);
        continue;
      }

      const numCols = Math.max(...rawRows.map((r) => r.length));

      if (numCols <= MAX_TABLE_COLS) {
        // ── Bordered table mode ──────────────────────────────────────────
        const colW = CONTENT_W / numCols;

        for (let r = 0; r < rawRows.length; r++) {
          // Page break
          if (y + ROW_H > PAGE_H - MARGIN) {
            doc.addPage();
            y = MARGIN;
          }

          const isHeader = r === 0;
          doc.font(isHeader ? "Helvetica-Bold" : "Helvetica").fontSize(isHeader ? HEADER_SIZE : FONT_SIZE);
          doc.strokeColor("#cccccc").lineWidth(0.5);

          for (let c = 0; c < numCols; c++) {
            const x = MARGIN + c * colW;
            const text = rawRows[r][c] ?? "";

            // Border
            doc.rect(x, y, colW, ROW_H).stroke();

            // Text (2 pt padding inside cell)
            doc.fillColor("#000000").text(text, x + 2, y + 3, {
              width: colW - 4,
              height: ROW_H - 4,
              ellipsis: true,
              lineBreak: false,
            });
          }
          y += ROW_H;
        }
      } else {
        // ── Text-line fallback for very wide sheets ──────────────────────
        doc.font("Helvetica").fontSize(FONT_SIZE - 1);
        for (let r = 0; r < rawRows.length; r++) {
          if (y + 12 > PAGE_H - MARGIN) {
            doc.addPage();
            y = MARGIN;
          }
          const line = rawRows[r].join("  |  ");
          if (r === 0) doc.font("Helvetica-Bold");
          else doc.font("Helvetica");
          doc.text(line, MARGIN, y, { lineBreak: false, ellipsis: true, width: CONTENT_W });
          y += 11;
        }
      }
    }

    doc.end();
  });
}
