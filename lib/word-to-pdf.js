// Reconstruct a PDF from a .docx file using mammoth (text extraction) and
// pdfkit (PDF generation). This is a "light" reconstruction — text flows as
// paragraphs with basic heading detection; the exact visual layout of the
// original Word document is NOT reproduced. That is an honest trade-off for
// zero Docker / zero LibreOffice dependency.
//
// Heading heuristic: mammoth returns HTML; we tag <h1>-<h6> lines as headings
// and render them larger + bold, which gives a clean, readable result for the
// vast majority of documents (reports, letters, forms, manuals).

import mammoth from "mammoth";
import PDFDocument from "pdfkit";

const MARGIN = 72; // 1 inch
const FONT_BODY = 11;
const LINE_GAP = 4;

// Heading level → { size (pt), bold, spaceBefore }
const HEADING = {
  h1: { size: 20, spaceBefore: 12 },
  h2: { size: 16, spaceBefore: 10 },
  h3: { size: 13, spaceBefore: 8 },
  h4: { size: 12, spaceBefore: 6 },
  h5: { size: 11, spaceBefore: 4 },
  h6: { size: 10, spaceBefore: 4 },
};

/**
 * @param {Buffer} buffer  Raw .docx bytes.
 * @returns {Promise<Buffer>}  PDF bytes.
 */
export async function convertWordToPdf(buffer) {
  // mammoth.convertToHtml gives us heading-aware HTML; extractRawText loses structure.
  const { value: html } = await mammoth.convertToHtml({ buffer });

  // Parse blocks in order: headings, paragraphs, list items.
  const blocks = [];
  const blockRe = /<(h[1-6]|p|li)([^>]*)>([\s\S]*?)<\/\1>/gi;
  let m;
  while ((m = blockRe.exec(html)) !== null) {
    const tag = m[1].toLowerCase();
    const raw = m[3];
    // Strip inner HTML tags, decode basic HTML entities.
    const text = raw
      .replace(/<[^>]+>/g, "")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&nbsp;/g, " ")
      .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)))
      .replace(/\s+/g, " ")
      .trim();
    if (text) blocks.push({ tag, text });
  }

  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({
      size: "A4",
      margins: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN },
      info: { Creator: "pdfprosafe.com" },
    });
    const chunks = [];
    doc.on("data", (c) => chunks.push(c));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);

    if (blocks.length === 0) {
      doc.font("Helvetica").fontSize(FONT_BODY).text("(empty document)");
    } else {
      for (const { tag, text } of blocks) {
        const h = HEADING[tag];
        if (h) {
          doc.moveDown(h.spaceBefore / FONT_BODY);
          doc.font("Helvetica-Bold").fontSize(h.size).text(text, { lineGap: LINE_GAP });
          doc.moveDown(0.3);
          doc.font("Helvetica").fontSize(FONT_BODY);
        } else {
          // p or li
          const prefix = tag === "li" ? "• " : "";
          doc.font("Helvetica").fontSize(FONT_BODY).text(prefix + text, { lineGap: LINE_GAP, paragraphGap: 6 });
        }
      }
    }

    doc.end();
  });
}
