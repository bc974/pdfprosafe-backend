// pdfprosafe.com — stateless conversion service.
//
// Two endpoints turn an uploaded PDF into an editable Office file:
//   POST /convert/word   → .docx
//   POST /convert/excel  → .xlsx
//
// Privacy by construction: multer keeps the upload in memory (never on disk),
// pdfjs reads it, the result buffer is streamed back, and every buffer is
// garbage-collected when the request ends. Nothing is logged or persisted.

import express from "express";
import cors from "cors";
import multer from "multer";

import { extractPdf } from "./lib/extract.js";
import { buildWord } from "./lib/to-word.js";
import { buildExcel } from "./lib/to-excel.js";

const PORT = process.env.PORT || 8080;
const MAX_BYTES = Number(process.env.MAX_UPLOAD_BYTES || 25 * 1024 * 1024); // 25 MB

// Origins allowed to call the API. Override in prod via ALLOWED_ORIGINS
// (comma-separated). Defaults cover the live site and local Next dev server.
const ALLOWED = (
  process.env.ALLOWED_ORIGINS ||
  "https://pdfprosafe.com,https://www.pdfprosafe.com,http://localhost:3333"
)
  .split(",")
  .map((s) => s.trim())
  .filter(Boolean);

const app = express();
app.disable("x-powered-by");

app.use(
  cors({
    origin(origin, cb) {
      // Allow same-origin/no-origin (curl, health checks) and whitelisted sites.
      if (!origin || ALLOWED.includes(origin)) return cb(null, true);
      cb(new Error("Origin not allowed"));
    },
    methods: ["POST", "GET", "OPTIONS"],
    maxAge: 86400,
  }),
);

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: MAX_BYTES, files: 1 },
  fileFilter(_req, file, cb) {
    const ok =
      file.mimetype === "application/pdf" ||
      file.originalname.toLowerCase().endsWith(".pdf");
    // Second arg MUST be `true` to accept; cb(null) alone silently drops it.
    if (ok) cb(null, true);
    else cb(httpError(415, "bad_file", "Only PDF files are accepted."));
  },
});

// ─── Health ───────────────────────────────────────────────────────────────
app.get(["/", "/health"], (_req, res) => {
  res.json({ ok: true, service: "pdfprosafe-backend", time: Date.now() });
});

// ─── Conversion routes ──────────────────────────────────────────────────────
const CONVERTERS = {
  word: {
    build: buildWord,
    ext: "docx",
    mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  },
  excel: {
    build: buildExcel,
    ext: "xlsx",
    mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  },
};

for (const [target, cfg] of Object.entries(CONVERTERS)) {
  app.post(`/convert/${target}`, upload.single("file"), async (req, res, next) => {
    try {
      if (!req.file) throw httpError(400, "no_file", "No file was uploaded.");

      let extracted;
      try {
        extracted = await extractPdf(req.file.buffer);
      } catch (parseErr) {
        // pdfjs throws on corrupt or encrypted PDFs — surface a clean 4xx
        // rather than a 500, and point encrypted files at the Unlock tool.
        if (parseErr?.name === "PasswordException") {
          throw httpError(
            422,
            "encrypted",
            "This PDF is password-protected. Remove the password first (Unlock PDF), then convert.",
          );
        }
        throw httpError(
          422,
          "bad_pdf",
          "This file could not be read as a PDF — it may be corrupt.",
        );
      }
      // A scanned/image-only PDF has (almost) no extractable text — there is
      // nothing to reconstruct. Tell the client to OCR first.
      if (extracted.charCount < 8) {
        throw httpError(
          422,
          "no_text",
          "This PDF has no extractable text (it may be scanned). Run OCR first.",
        );
      }

      const out = await cfg.build(extracted);
      const base = safeBase(req.file.originalname);
      res.setHeader("Content-Type", cfg.mime);
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="${base}.${cfg.ext}"`,
      );
      res.setHeader("Cache-Control", "no-store");
      res.send(out);
    } catch (err) {
      next(err);
    }
  });
}

// ─── Error handling ─────────────────────────────────────────────────────────
// eslint-disable-next-line no-unused-vars
app.use((err, _req, res, _next) => {
  if (err instanceof multer.MulterError) {
    const code = err.code === "LIMIT_FILE_SIZE" ? "too_large" : "upload_error";
    const status = err.code === "LIMIT_FILE_SIZE" ? 413 : 400;
    return res.status(status).json({ error: code, message: err.message });
  }
  const status = err.status || 500;
  const code = err.code || "server_error";
  const message =
    status === 500 ? "Conversion failed. Please try another file." : err.message;
  res.status(status).json({ error: code, message });
});

function httpError(status, code, message) {
  const e = new Error(message);
  e.status = status;
  e.code = code;
  return e;
}

/** Strip path/extension and unsafe chars from an uploaded filename. */
function safeBase(name) {
  const noPath = name.replace(/^.*[\\/]/, "");
  const noExt = noPath.replace(/\.[^.]+$/, "");
  const clean = noExt.replace(/[^\w\-. ]+/g, "_").trim();
  return clean || "converted";
}

app.listen(PORT, () => {
  console.log(`pdfprosafe-backend listening on :${PORT}`);
  console.log(`allowed origins: ${ALLOWED.join(", ")}`);
});
