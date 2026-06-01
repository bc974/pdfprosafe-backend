// Generates ../sample.pdf for the smoke test, borrowing @cantoo/pdf-lib from
// the website repo so we don't add it as a backend dependency.
import { createRequire } from "node:module";

const require = createRequire(
  "C:/Users/causs/Desktop/Files Acces/allpapersfree/package.json",
);
const { PDFDocument, StandardFonts } = require("@cantoo/pdf-lib");

const doc = await PDFDocument.create();
const font = await doc.embedFont(StandardFonts.Helvetica);
const bold = await doc.embedFont(StandardFonts.HelveticaBold);

// Page 1 — flowing prose (exercises the Word path).
{
  const page = doc.addPage([595, 842]); // A4
  let y = 800;
  const line = (text, size = 12, f = font) => {
    page.drawText(text, { x: 50, y, size, font: f });
    y -= size + 8;
  };
  line("Rapport de conversion", 20, bold);
  y -= 6;
  line("Ceci est un document de test pour pdfprosafe.com.");
  line("Il contient plusieurs paragraphes de texte normal afin de");
  line("verifier la reconstruction en Word (.docx) avec des lignes");
  line("qui s'enchainent correctement.");
}

// Page 2 — an aligned table (exercises the Excel path).
{
  const page = doc.addPage([595, 842]);
  const cols = [50, 230, 360, 470];
  const rows = [
    ["Produit", "Categorie", "Quantite", "Prix"],
    ["Clavier", "Peripherique", "12", "29.90"],
    ["Souris", "Peripherique", "30", "15.50"],
    ["Ecran 27", "Affichage", "5", "199.00"],
  ];
  let y = 780;
  for (const row of rows) {
    row.forEach((cell, c) => {
      page.drawText(cell, { x: cols[c], y, size: 12, font });
    });
    y -= 26;
  }
}

const bytes = await doc.save();
const { writeFile } = await import("node:fs/promises");
await writeFile(new URL("../sample.pdf", import.meta.url), bytes);
console.log(`wrote sample.pdf (${bytes.length} bytes)`);
