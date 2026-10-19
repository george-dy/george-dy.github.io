/* tools/build-bank.mjs — Builds the published question bank from the real
 * .docx uploads using the SAME shared parser code the browser uses.
 *
 * Outputs (stage 1):
 *   ../assets-raw/<docslug>-img-<n>.<ext>   extracted images (uncompressed)
 *   ../question-bank.staging.json           bank with src="assets-raw/..."
 *
 * Then run:  python3 tools/compress-assets.py
 * which compresses/resizes assets-raw/ -> assets/ and writes the final
 * ../question-bank.json + ../question-bank.js with src="assets/...".
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";
import { extractIRFromFile } from "./ir-extract.mjs";

const require = createRequire(import.meta.url);
const parser = require("../js/parser-core.js");

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const UPLOADS = process.env.UPLOADS_DIR || path.join(ROOT, "..", "uploads");

const FILES = [
  "1 Newborn_Infant_Practice_Examinations.docx",
  "2 Newborn Exam cases.docx",
  "3 Toddler and School Age questions.docx",
  "4 Developmental_Behavioral_ADHD.docx",
].map((f) => path.join(UPLOADS, f));

function slug(s) {
  return String(s).toLowerCase().normalize("NFKD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 48);
}

async function main() {
  const irs = [];
  for (const f of FILES) {
    if (!fs.existsSync(f)) { console.error("Missing source file:", f); process.exit(1); }
    console.log("Reading:", path.basename(f));
    irs.push(await extractIRFromFile(f));
  }

  const result = parser.parseExamDocuments(irs);
  console.log("Parsed:", JSON.stringify(result.stats));

  // Externalize images to assets-raw/
  const rawDir = path.join(ROOT, "assets-raw");
  fs.rmSync(rawDir, { recursive: true, force: true });
  fs.mkdirSync(rawDir, { recursive: true });
  let imgCount = 0;
  const seenData = new Map(); // dedupe identical images by dataUrl
  for (const q of result.questions) {
    for (const im of q.images || []) {
      if (!im.src || !im.src.startsWith("data:")) continue;
      let assetName = seenData.get(im.src);
      if (!assetName) {
        const ext = (im.contentType.split("/")[1] || "png").replace("jpeg", "jpg");
        imgCount++;
        assetName = `${slug(q.sourceFile)}-img-${String(imgCount).padStart(3, "0")}.${ext}`;
        const b64 = im.src.split(",")[1];
        fs.writeFileSync(path.join(rawDir, assetName), Buffer.from(b64, "base64"));
        seenData.set(im.src, assetName);
      }
      im.src = `assets-raw/${assetName}`;
    }
  }
  console.log("Extracted", imgCount, "unique images to assets-raw/ (", seenData.size, "deduped refs )");

  const bank = parser.assembleBank(result, {
    title: "Advanced Pediatric Primary Care — Practice Examination Bank",
  });
  bank.info = {
    sourceDocuments: result.documents.map((d) => ({ file: d.file, title: d.title, sets: d.setIds.length })),
    note: "Built from the course .docx documents with the same parser that powers the in-app Import tool. Rationales and answer keys are reproduced exactly as written in the source files.",
  };

  const v = parser.validateBank(bank);
  if (!v.valid) {
    console.error("Bank validation FAILED:");
    v.errors.slice(0, 20).forEach((e) => console.error("  -", e));
    process.exit(1);
  }
  console.log("Validation OK. warnings:", v.warnings.length);
  v.warnings.slice(0, 10).forEach((w) => console.log("  warn:", w));

  // attach parse warnings summary to bank for transparency
  bank.info.parseWarnings = result.warnings.map((w) => ({
    level: w.level, code: w.code, file: w.file || null, setId: w.setId || null, message: w.message,
  }));

  fs.writeFileSync(path.join(ROOT, "question-bank.staging.json"), JSON.stringify(bank));
  console.log("Wrote question-bank.staging.json (" + (fs.statSync(path.join(ROOT, "question-bank.staging.json")).size / 1e6).toFixed(1) + " MB)");
  console.log("Next: python3 tools/compress-assets.py");
}

main().catch((e) => { console.error(e); process.exit(1); });
