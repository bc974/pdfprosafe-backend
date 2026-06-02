// End-to-end HTTP test using the same primitives a browser uses
// (fetch + FormData + File). Server must already be running on $BASE.
import { readFile, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { tmpdir } from "node:os";
import { join } from "node:path";
// (readFile already imported above — just used again for smoke outputs)

const BASE = process.env.BASE || "http://localhost:8080";
const pdf = fileURLToPath(new URL("../sample.pdf", import.meta.url));
const bytes = await readFile(pdf);
const TMP = tmpdir();
let failures = 0;
const ok = (c, m) => { if (!c) { failures++; console.error("  ✗ " + m); } else console.log("  ✓ " + m); };

async function post(path, blob, filename) {
  const fd = new FormData();
  if (blob) fd.append("file", blob, filename);
  const res = await fetch(BASE + path, { method: "POST", body: fd });
  return res;
}

// Word
{
  const res = await post("/convert/word", new Blob([bytes], { type: "application/pdf" }), "sample.pdf");
  const buf = Buffer.from(await res.arrayBuffer());
  await writeFile(join(TMP, "http.docx"), buf);
  console.log(`WORD  ${res.status}  ${res.headers.get("content-type")}`);
  console.log(`      ${res.headers.get("content-disposition")}`);
  ok(res.status === 200, "word status 200");
  ok(buf[0] === 0x50 && buf[1] === 0x4b, `word is OOXML zip (${buf.length}B)`);
  ok((res.headers.get("content-disposition") || "").includes("sample.docx"), "word filename");
}
// Excel
{
  const res = await post("/convert/excel", new Blob([bytes], { type: "application/pdf" }), "sample.pdf");
  const buf = Buffer.from(await res.arrayBuffer());
  await writeFile(join(TMP, "http.xlsx"), buf);
  console.log(`EXCEL ${res.status}  ${res.headers.get("content-type")}`);
  ok(res.status === 200, "excel status 200");
  ok(buf[0] === 0x50 && buf[1] === 0x4b, `excel is OOXML zip (${buf.length}B)`);
}
// no_file
{
  const res = await post("/convert/word", null);
  const j = await res.json();
  ok(res.status === 400 && j.error === "no_file", `no-file → 400 ${j.error}`);
}
// bad type (text masquerading)
{
  const res = await post("/convert/word", new Blob(["hello"], { type: "text/plain" }), "notes.txt");
  const j = await res.json().catch(() => ({}));
  ok(res.status === 415 && j.error === "bad_file", `non-pdf → 415 ${j.error}`);
}
// no_text (valid PDF magic but no extractable text → a tiny fake won't parse;
// instead send a pdf-looking blob to confirm graceful 4xx, not a 500 crash)
{
  const res = await post("/convert/word", new Blob([Buffer.from("%PDF-1.4\n%%EOF")], { type: "application/pdf" }), "empty.pdf");
  ok(res.status >= 400 && res.status < 500, `unparseable pdf → ${res.status} (graceful, no 500)`);
}

// ── Reverse: Word→PDF and Excel→PDF ──────────────────────────────────────────
// Re-use the smoke outputs (they exist after the forward conversions ran earlier
// in this same Node session — for the HTTP test we need them as file buffers).
const TMP2 = tmpdir();
const docxPath = join(TMP2, "smoke.docx");
const xlsxPath = join(TMP2, "smoke.xlsx");
let docxBuf, xlsxBuf;
try {
  docxBuf = await readFile(docxPath);
  xlsxBuf = await readFile(xlsxPath);
} catch {
  // Smoke files not present — skip reverse tests (run smoke first)
  console.log("  (skipping reverse tests: run smoke.mjs first)");
}

if (docxBuf) {
  const res = await post("/convert/word-to-pdf",
    new Blob([docxBuf], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }),
    "test.docx");
  const buf = Buffer.from(await res.arrayBuffer());
  ok(res.status === 200, `word-to-pdf status 200 (${res.status})`);
  ok(buf[0] === 0x25 && buf[1] === 0x50, `word-to-pdf returns a PDF (${buf.length}B)`);
}
if (xlsxBuf) {
  const res = await post("/convert/excel-to-pdf",
    new Blob([xlsxBuf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    "test.xlsx");
  const buf = Buffer.from(await res.arrayBuffer());
  ok(res.status === 200, `excel-to-pdf status 200 (${res.status})`);
  ok(buf[0] === 0x25 && buf[1] === 0x50, `excel-to-pdf returns a PDF (${buf.length}B)`);
}
// Wrong type for reverse endpoint
{
  const res = await post("/convert/word-to-pdf", new Blob([bytes], { type: "application/pdf" }), "test.pdf");
  ok(res.status === 415, `word-to-pdf rejects non-docx → ${res.status}`);
}

console.log(failures ? `\n✗ ${failures} HTTP assertion(s) failed` : "\n✓ ALL HTTP TESTS PASSED");
// Let undici's keep-alive sockets finish closing before exit, otherwise libuv
// trips an assertion on Windows when process.exit races a closing handle.
setTimeout(() => process.exit(failures ? 1 : 0), 300);
