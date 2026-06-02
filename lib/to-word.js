// Reconstruct an editable .docx from extracted PDF geometry.
//
// Per page we decide: does it read as a table or as prose?
//   - Table page  → a real Word table (rows × column cells).
//   - Prose page  → flowing paragraphs; lines clearly larger than the body text
//     are emitted bold (headings).
// Font size is approximated from glyph height. Pages are separated by a page
// break. The result is editable content — not a pixel clone — which is what a
// "PDF to Word" conversion is actually used for.

import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  PageBreak,
  Table,
  TableRow,
  TableCell,
  WidthType,
} from "docx";

import {
  groupLines,
  deriveColumns,
  lineToCells,
  medianHeight,
  bodyFontSize,
  isTabularPage,
} from "./group.js";

/** Clamp a glyph height to a sane Word point size. */
function ptOf(height) {
  return Math.min(48, Math.max(7, Math.round(height)));
}

/** A line of prose → a paragraph (bold when it's a heading). */
function proseParagraph(line, bodySize) {
  const text = line.items
    .map((it) => it.str)
    .join("")
    .replace(/\s+/g, " ")
    .trim();
  if (text.length === 0) return new Paragraph({ children: [] });
  const h = medianHeight(line);
  const isHeading = h >= bodySize * 1.15;
  return new Paragraph({
    children: [new TextRun({ text, size: ptOf(h) * 2, bold: isHeading })],
  });
}

/** A set of table-like lines → a Word table. */
function buildTable(lines, bounds, bodySize) {
  const rows = lines.map(
    (line) =>
      new TableRow({
        children: lineToCells(line, bounds).map(
          (val) =>
            new TableCell({
              children: [
                new Paragraph({
                  children: val
                    ? [new TextRun({ text: val, size: ptOf(bodySize) * 2 })]
                    : [],
                }),
              ],
            }),
        ),
      }),
  );
  return new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } });
}

/**
 * @param {{pages: Array}} extracted  Output of extractPdf().
 * @returns {Promise<Buffer>}  .docx bytes.
 */
export async function buildWord(extracted) {
  // Body text size is computed across the whole document so heading detection
  // is consistent page to page.
  const allLines = extracted.pages.flatMap((p) => groupLines(p.items));
  const bodySize = bodyFontSize(allLines);

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

    const { bounds } = deriveColumns(page.items);
    if (isTabularPage(lines, bounds)) {
      children.push(buildTable(lines, bounds, bodySize));
    } else {
      for (const line of lines) children.push(proseParagraph(line, bodySize));
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
