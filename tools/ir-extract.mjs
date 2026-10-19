/* tools/ir-extract.mjs — Node-side .docx -> IR using the SAME shared modules
 * as the browser (js/docx-ir.js + js/parser-core.js). Used by build-bank.mjs
 * and by the test suite. */
import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const JSZip = require(process.env.JSZIP_PATH || "jszip");
const { DOMParser } = require(process.env.XMDOM_PATH || "@xmldom/xmldom");
const DocxIR = require(path.join(path.dirname(new URL(import.meta.url).pathname), "..", "js", "docx-ir.js"));

export async function extractIRFromFile(filePath, opts = {}) {
  const buf = fs.readFileSync(filePath);
  const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  return await DocxIR.readDocxIR(ab, { JSZip, DOMParser }, path.basename(filePath), opts);
}

export { JSZip, DOMParser, DocxIR };
