// Pure geometry helpers shared by the Word and Excel builders. No I/O, no
// pdfjs — just clustering of positioned text items so both converters agree on
// what a "line", a "row", and a "column" are.

/**
 * Group a page's items into visual lines (top-to-bottom), each line's items
 * left-to-right. Items are considered the same line when their vertical centers
 * fall within `tol` (derived from glyph height).
 *
 * @returns {Array<{ y: number, items: Array }>}
 */
export function groupLines(items) {
  if (items.length === 0) return [];
  const sorted = [...items].sort((a, b) => a.y - b.y || a.x - b.x);
  const lines = [];
  let current = null;

  for (const it of sorted) {
    const tol = Math.max(3, it.h * 0.6);
    if (current && Math.abs(it.y - current.y) <= tol) {
      current.items.push(it);
      // Running average keeps the band centered as it grows.
      current.y = (current.y * (current.items.length - 1) + it.y) / current.items.length;
    } else {
      current = { y: it.y, items: [it] };
      lines.push(current);
    }
  }

  for (const line of lines) {
    line.items.sort((a, b) => a.x - b.x);
  }
  return lines;
}

/**
 * Join a line's items into a single string, inserting a space when the
 * horizontal gap between two runs is wider than a fraction of the glyph height
 * (i.e. a real word break that pdfjs split across items).
 */
export function lineToText(line) {
  let out = "";
  let prevEnd = null;
  for (const it of line.items) {
    if (prevEnd != null) {
      const gap = it.x - prevEnd;
      const needsSpace = gap > Math.max(1.5, it.h * 0.25);
      const endsWithSpace = /\s$/.test(out);
      const startsWithSpace = /^\s/.test(it.str);
      if (needsSpace && !endsWithSpace && !startsWithSpace) out += " ";
    }
    out += it.str;
    prevEnd = it.x + it.w;
  }
  return out.replace(/\s+/g, " ").trim();
}

/**
 * Derive a column model for a page from the x-start of every item. Nearby
 * x-positions are merged into a single column band (within `tol`). Returns the
 * sorted band start positions; map an item to a column with `columnIndexFor`.
 *
 * @returns {{ bounds: number[], tol: number }}
 */
export function deriveColumns(items, tol = 12) {
  // pdfjs emits a trailing whitespace run after many words; its x sits just
  // after the word and would spawn a phantom column. Only real (non-blank)
  // runs define columns.
  const xs = items
    .filter((it) => it.str.trim().length > 0)
    .map((it) => it.x)
    .sort((a, b) => a - b);
  const bounds = [];
  for (const x of xs) {
    const last = bounds[bounds.length - 1];
    if (last == null || x - last > tol) bounds.push(x);
  }
  return { bounds, tol };
}

/** Index (0-based) of the column band whose start is closest to `x`. */
export function columnIndexFor(x, bounds) {
  let best = 0;
  let bestDist = Infinity;
  for (let i = 0; i < bounds.length; i++) {
    const d = Math.abs(x - bounds[i]);
    if (d < bestDist) {
      bestDist = d;
      best = i;
    }
  }
  return best;
}
