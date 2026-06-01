// PDF text extraction with positional geometry, using pdfjs-dist in Node.
//
// We only ever read text + coordinates — we never rasterize a page — so there
// is no dependency on `canvas`/`DOMMatrix`. pdfjs falls back to a main-thread
// "fake worker" automatically in Node, which is fine for getTextContent().
//
// Output shape (per document):
//   { pages: [ { width, height, items: [ { x, y, w, h, str } ] } ], charCount }
// Coordinates are in PDF user units with a TOP-LEFT origin (we flip pdfjs's
// bottom-left y so downstream row/line grouping reads naturally).

import { getDocument } from "pdfjs-dist/legacy/build/pdf.mjs";

/**
 * @param {Buffer|Uint8Array} buffer  Raw PDF bytes.
 * @returns {Promise<{pages: Array, charCount: number}>}
 */
export async function extractPdf(buffer) {
  // pdfjs wants a Uint8Array it can detach; copy so we never mutate the caller's
  // buffer (and so multer's pooled memory isn't held hostage by pdfjs).
  const data = new Uint8Array(buffer);

  const loadingTask = getDocument({
    data,
    // Node-friendly hardening: no eval, no network font fetches, no system fonts.
    isEvalSupported: false,
    useSystemFonts: false,
    disableFontFace: true,
    // Keep memory bounded on the free tier.
    cMapUrl: undefined,
    verbosity: 0,
  });

  const doc = await loadingTask.promise;
  const pages = [];
  let charCount = 0;

  try {
    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p);
      const viewport = page.getViewport({ scale: 1 });
      const pageHeight = viewport.height;
      const content = await page.getTextContent({
        includeMarkedContent: false,
        disableNormalization: false,
      });

      const items = [];
      for (const item of content.items) {
        // Marked-content items have no `str`.
        if (typeof item.str !== "string") continue;
        if (item.str.length === 0) continue;
        const t = item.transform; // [a, b, c, d, e, f]
        const x = t[4];
        const yBottom = t[5];
        // Glyph height ≈ vertical scale of the text matrix.
        const h = Math.hypot(t[1], t[3]) || item.height || 10;
        const w = item.width || 0;
        // Flip to a top-left origin so smaller y == higher on the page.
        const y = pageHeight - yBottom;
        items.push({ x, y, w, h, str: item.str });
        charCount += item.str.trim().length;
      }

      pages.push({ width: viewport.width, height: pageHeight, items });
      // Let pdfjs reclaim per-page memory immediately.
      page.cleanup();
    }
  } finally {
    await doc.destroy();
  }

  return { pages, charCount };
}
