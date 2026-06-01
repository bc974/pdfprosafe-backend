# pdfprosafe-backend

Stateless Node.js conversion service for **pdfprosafe.com**. Turns an uploaded
PDF into an editable Office file:

| Endpoint | Method | Field | Returns |
|---|---|---|---|
| `/convert/word`  | `POST` | `file` (multipart) | `.docx` |
| `/convert/excel` | `POST` | `file` (multipart) | `.xlsx` |
| `/health`        | `GET`  | —                  | `{ ok: true }` |

## Privacy by construction

The upload lives **only in memory** (`multer.memoryStorage`). pdfjs reads the
buffer, the result is streamed back, and every buffer is garbage-collected when
the request ends. Nothing is written to disk and nothing is logged. This is what
backs the "processed then immediately deleted" notice on the website.

## How it works

1. `lib/extract.js` — pdfjs-dist pulls every text run with its `(x, y)` position.
2. `lib/group.js` — pure geometry: clusters runs into lines, rows, and columns.
3. `lib/to-word.js` — emits flowing paragraphs (`docx`), one page break per page.
4. `lib/to-excel.js` — emits a real cell grid (`exceljs`), one sheet per page.

It reconstructs **editable** content — not a pixel clone of the layout. Scanned
/ image-only PDFs have no text to extract and return `422 no_text` (OCR first).

## Run locally

```bash
npm install
npm start            # listens on :8080 (or $PORT)
npm run smoke ./some.pdf   # offline conversion check, writes /tmp/smoke.{docx,xlsx}
```

Test the live endpoints:

```bash
curl -F "file=@some.pdf" http://localhost:8080/convert/word  -o out.docx
curl -F "file=@some.pdf" http://localhost:8080/convert/excel -o out.xlsx
```

## Deploy to Render (free)

**Option A — Blueprint (recommended):** push this folder to its own GitHub repo,
then in Render: **New → Blueprint**, pick the repo. `render.yaml` provisions the
service automatically.

**Option B — Manual:** **New → Web Service** → connect the repo →
- Runtime: **Node**
- Build command: `npm install`
- Start command: `npm start`
- Health check path: `/health`
- Env var `ALLOWED_ORIGINS` = `https://pdfprosafe.com,https://www.pdfprosafe.com`

After the first deploy, Render gives you a URL like
`https://pdfprosafe-backend.onrender.com`. Put it in the frontend env var
`NEXT_PUBLIC_CONVERT_API` (see the website repo).

### Free-tier note

The free instance sleeps after ~15 min idle and cold-starts (~30–50 s) on the
next request. The website shows a "waking the server" hint so the first user of
the hour isn't confused. Conversions themselves take 1–5 s.

## Environment variables

| Var | Default | Purpose |
|---|---|---|
| `PORT` | `8080` | Render sets this automatically. |
| `ALLOWED_ORIGINS` | localhost + pdfprosafe.com | CORS allowlist (comma-separated). |
| `MAX_UPLOAD_BYTES` | `26214400` (25 MB) | Reject larger uploads. |
