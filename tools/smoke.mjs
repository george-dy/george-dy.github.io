/* tools/smoke.mjs — quick smoke test: parse a real document and print a
 * structural summary. Usage: node tools/smoke.mjs <file.docx> [setIdx] [qIdx] */
import { extractIRFromFile } from "./ir-extract.mjs";
import { createRequire } from "node:module";
const require = createRequire(import.meta.url);
const parser = require("../js/parser-core.js");

const file = process.argv[2];
const ir = await extractIRFromFile(file);
console.log("paragraphs:", ir.paragraphs.length);
const t0 = Date.now();
const result = parser.parseExamDocuments([ir]);
console.log("parse ms:", Date.now() - t0);
console.log("stats:", JSON.stringify(result.stats, null, 1));
console.log("sets:");
result.sets.forEach((s) => console.log(`  [${s.id}] "${s.title}" q=${s.questionIds.length} conf=${s.metadata.confidence} by=${s.metadata.detectedBy}`));
const warns = {};
result.warnings.forEach((w) => { const k = w.level + ":" + w.code; warns[k] = (warns[k] || 0) + 1; });
console.log("warnings:", warns);

// detailed dump of one question if requested
const si = parseInt(process.argv[3] || "0", 10);
const qi = parseInt(process.argv[4] || "0", 10);
const set = result.sets[si];
if (set) {
  const q = result.questions.find((q) => q.id === set.questionIds[qi]);
  if (q) console.log(JSON.stringify(q, null, 1).slice(0, 3000));
}
