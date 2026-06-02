// PDF → PowerPoint (.pptx), in memory, pure Node.
//
// Strategy: there is no faithful pure-JS PPTX layout engine, but a PDF *is*
// already a fixed-layout document — so we rasterize each page to a PNG and drop
// it onto its own slide, full-bleed. This is exactly what mainstream "PDF to
// PPT" converters do, and it works on ANY pdf (including scanned ones, unlike
// the text-extraction converters).
//
// Rendering uses `pdf-to-img`, which bundles pdfjs + @napi-rs/canvas (prebuilt
// native binaries, no system libraries) — so it runs on Render's plain Node
// runtime with no Docker/apt step.

import { pdf } from "pdf-to-img";
import pptxgen from "pptxgenjs";

// Note on fonts: pdfjs in Node can't fetch its standard-14 font files over the
// file:// scheme (undici has no file:// support), so PDFs that reference
// non-embedded fonts render with pdfjs's built-in metric-compatible fallbacks.
// That's fine for a rasterized slide; PDFs with embedded fonts are pixel-exact.

// Image rendering is heavier than text extraction, so the ceilings are tighter
// than the text converters' MAX_PAGES to stay within the request timeout and
// the free instance's memory.
export const MAX_PPT_PAGES = Number(process.env.MAX_PPT_PAGES || 40);
const RENDER_SCALE = Number(process.env.PPT_RENDER_SCALE || 1.5);

/** Read width/height (px) from a PNG buffer's IHDR chunk (big-endian uint32). */
function pngSize(buf) {
  return { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) };
}

/**
 * Convert a PDF buffer into a .pptx buffer (one slide per page, page rendered
 * full-bleed). Throws `{ code: "too_many_pages" }` past the page ceiling and
 * lets pdfjs's PasswordException propagate (the route maps it to `encrypted`).
 */
export async function convertPdfToPowerpoint(buffer) {
  const doc = await pdf(buffer, { scale: RENDER_SCALE });

  const total = doc.length;
  if (!total || total < 1) {
    const e = new Error("empty pdf");
    e.code = "bad_pdf";
    throw e;
  }
  if (total > MAX_PPT_PAGES) {
    const e = new Error(`too many pages (${total} > ${MAX_PPT_PAGES})`);
    e.code = "too_many_pages";
    throw e;
  }

  const pptx = new pptxgen();
  pptx.author = "pdfprosafe.com";

  let layoutW = 10;
  let layoutH = 7.5;
  let layoutSet = false;

  for await (const page of doc) {
    // `page` is a PNG Buffer.
    if (!layoutSet) {
      const { width, height } = pngSize(page);
      // Match the slide aspect ratio to the first page so there's no
      // letterboxing on uniform-size PDFs. Base the long edge on 10in.
      const aspect = height / width;
      layoutW = 10;
      layoutH = Math.max(1, Math.min(56, +(layoutW * aspect).toFixed(2)));
      pptx.defineLayout({ name: "PDFPAGE", width: layoutW, height: layoutH });
      pptx.layout = "PDFPAGE";
      layoutSet = true;
    }

    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    slide.addImage({
      data: `data:image/png;base64,${page.toString("base64")}`,
      x: 0,
      y: 0,
      w: layoutW,
      h: layoutH,
      // `contain` centers + preserves aspect, so any page whose size differs
      // from the first is letterboxed rather than stretched.
      sizing: { type: "contain", w: layoutW, h: layoutH },
    });
  }

  return pptx.write({ outputType: "nodebuffer" });
}
