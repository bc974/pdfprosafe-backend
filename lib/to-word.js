// Reconstruct an editable .docx from extracted PDF geometry.
//
// Strategy: per page, group items into visual lines, emit each line as a
// paragraph, and approximate the font size from glyph height. Pages are
// separated by a page break so the document paginates like the source. This
// produces flowing, editable text — not a pixel clone — which is what a
// "PDF to Word" conversion is actually used for.

import { Document, Packer, Paragraph, TextRun, PageBreak } from "docx";

import { groupLines, lineToText } from "./group.js";

/**
 * @param {{pages: Array}} extracted  Output of extractPdf().
 * @returns {Promise<Buffer>}  .docx bytes.
 */
export async function buildWord(extracted) {
  const children = [];

  extracted.pages.forEach((page, pageIdx) => {
    if (pageIdx > 0) {
      children.push(new Paragraph({ children: [new PageBreak()] }));
    }

    const lines = groupLines(page.items);
    if (lines.length === 0) {
      children.push(new Paragraph({ children: [] }));
      return;
    }

    for (const line of lines) {
      const text = lineToText(line);
      if (text.length === 0) {
        children.push(new Paragraph({ children: [] }));
        continue;
      }
      // Median glyph height on the line → point size. docx sizes are in
      // half-points, so multiply by 2. Clamp to a sane range.
      const heights = line.items.map((it) => it.h).sort((a, b) => a - b);
      const medianH = heights[Math.floor(heights.length / 2)] || 11;
      const pt = Math.min(48, Math.max(7, Math.round(medianH)));
      children.push(
        new Paragraph({
          children: [new TextRun({ text, size: pt * 2 })],
        }),
      );
    }
  });

  if (children.length === 0) {
    children.push(new Paragraph({ children: [] }));
  }

  const doc = new Document({
    creator: "pdfprosafe.com",
    description: "Converted from PDF",
    sections: [{ children }],
  });

  return Packer.toBuffer(doc);
}
