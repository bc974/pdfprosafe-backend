// End-to-end HTTP test using the same primitives a browser uses
// (fetch + FormData + File). Server must already be running on $BASE.
import { readFile, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { tmpdir } from "node:os";
import { join } from "node:path";

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

console.log(failures ? `\n✗ ${failures} HTTP assertion(s) failed` : "\n✓ ALL HTTP TESTS PASSED");
// Let undici's keep-alive sockets finish closing before exit, otherwise libuv
// trips an assertion on Windows when process.exit races a closing handle.
setTimeout(() => process.exit(failures ? 1 : 0), 300);
