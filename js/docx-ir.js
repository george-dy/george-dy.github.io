/* ============================================================================
 * docx-ir.js — Converts a .docx (via JSZip) into the neutral "document IR"
 * consumed by parser-core.js. Works in the browser (native DOMParser) and in
 * Node (pass @xmldom/xmldom's DOMParser). No other dependencies besides JSZip.
 *
 * IR paragraph: { text, bold, italic, sizePt, style, bullet, images, table }
 * IR image:     { relId, name, descr, contentType, dataUrl, widthPx, heightPx,
 *                 error? }
 * ==========================================================================*/
(function (root, factory) {
  if (typeof module === "object" && module.exports) { module.exports = factory(); }
  else { root.DocxIR = factory(); }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
  const V = "urn:schemas-microsoft-com:vml";

  const SUP_MAP = {
    "0": "⁰", "1": "¹", "2": "²", "3": "³", "4": "⁴", "5": "⁵", "6": "⁶", "7": "⁷", "8": "⁸", "9": "⁹",
    "+": "⁺", "-": "⁻", "=": "⁼", "(": "⁽", ")": "⁾",
    "a": "ᵃ", "b": "ᵇ", "c": "ᶜ", "d": "ᵈ", "e": "ᵉ", "f": "ᶠ", "g": "ᵍ", "h": "ʰ", "i": "ⁱ",
    "j": "ʲ", "k": "ᵏ", "l": "ˡ", "m": "ᵐ", "n": "ⁿ", "o": "ᵒ", "p": "ᵖ", "r": "ʳ", "s": "ˢ",
    "t": "ᵗ", "u": "ᵘ", "v": "ᵛ", "w": "ʷ", "x": "ˣ", "y": "ʸ", "z": "ᶻ",
  };
  const SUB_MAP = {
    "0": "₀", "1": "₁", "2": "₂", "3": "₃", "4": "₄", "5": "₅", "6": "₆", "7": "₇", "8": "₈", "9": "₉",
    "+": "₊", "-": "₋", "=": "₌", "(": "₍", ")": "₎",
    "a": "ₐ", "e": "ₑ", "h": "ₕ", "i": "ᵢ", "j": "ⱼ", "k": "ₖ", "l": "ₗ", "m": "ₘ",
    "n": "ₙ", "o": "ₒ", "p": "ₚ", "r": "ᵣ", "s": "ₛ", "t": "ₜ", "u": "ᵤ", "v": "ᵥ", "x": "ₓ",
  };

  function toSup(text) {
    let out = "";
    for (const ch of text.toLowerCase()) out += SUP_MAP[ch] || "";
    return out || text;
  }
  function toSub(text) {
    let out = "";
    for (const ch of text.toLowerCase()) out += SUB_MAP[ch] || "";
    return out || text;
  }

  const MIME = {
    png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif",
    bmp: "image/bmp", webp: "image/webp", tif: "image/tiff", tiff: "image/tiff",
    svg: "image/svg+xml", emf: "image/x-emf", wmf: "image/x-wmf",
  };

  // --- minimal DOM helpers that work with both DOMParser implementations ----
  function localName(node) {
    const n = node.nodeName || "";
    const i = n.indexOf(":");
    return i === -1 ? n : n.slice(i + 1);
  }
  function childrenOf(node) {
    const out = [];
    const kids = node.childNodes;
    if (!kids) return out;
    for (let i = 0; i < kids.length; i++) out.push(kids[i]);
    return out;
  }
  function elementChildren(node) {
    return childrenOf(node).filter((c) => c.nodeType === 1);
  }
  function firstChildNamed(node, name) {
    for (const c of elementChildren(node)) if (localName(c) === name) return c;
    return null;
  }
  function descendantsNamed(node, name) {
    const out = [];
    (function walk(n) {
      for (const c of elementChildren(n)) {
        if (localName(c) === name) out.push(c);
        walk(c);
      }
    })(node);
    return out;
  }
  function attr(node, name) {
    if (!node || !node.getAttribute) return null;
    // try qualified names used in OOXML
    const v = node.getAttribute(name);
    if (v != null) return v;
    return null;
  }

  function normalizeStyleId(id) {
    return String(id || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  }

  /* ------------------------------------------------------------------ */

  async function readRelationships(zip) {
    const rels = {};
    const candidates = ["word/_rels/document.xml.rels"];
    for (const path of candidates) {
      const f = zip.file(path);
      if (!f) continue;
      const xml = await f.async("string");
      const doc = readRelationships._parser.parseFromString(xml, "application/xml");
      for (const rel of descendantsNamed(doc, "Relationship")) {
        const id = attr(rel, "Id");
        const target = attr(rel, "Target");
        const mode = attr(rel, "TargetMode");
        if (id && target) rels[id] = { target, external: mode === "External" };
      }
    }
    return rels;
  }

  function resolveMediaPath(target) {
    // targets are relative to word/ (e.g. "media/image1.png" or "/word/media/..")
    let t = String(target).replace(/\\/g, "/");
    if (t.startsWith("/")) t = t.replace(/^\/word\//, "word/");
    else if (!t.startsWith("word/")) t = "word/" + t;
    // collapse ../ segments
    const parts = [];
    t.split("/").forEach((seg) => {
      if (seg === "..") parts.pop();
      else if (seg !== "." && seg !== "") parts.push(seg);
    });
    return parts.join("/");
  }

  async function loadImage(zip, rels, relId, seen) {
    const rec = { relId };
    const rel = rels[relId];
    if (!rel) { rec.error = "relationship-not-found:" + relId; return rec; }
    if (rel.external) { rec.error = "external-image-not-embedded:" + rel.target; return rec; }
    const path = resolveMediaPath(rel.target);
    rec.name = path.split("/").pop();
    const ext = (rec.name.split(".").pop() || "").toLowerCase();
    rec.contentType = MIME[ext] || "application/octet-stream";
    const entry = zip.file(path);
    if (!entry) { rec.error = "media-file-missing:" + path; return rec; }
    if (ext === "emf" || ext === "wmf") {
      rec.error = "unsupported-image-format:" + ext.toUpperCase() + " (Windows metafile cannot be displayed in browsers)";
      // still keep raw dataUrl so nothing is lost
    }
    try {
      const b64 = await entry.async("base64");
      rec.dataUrl = "data:" + rec.contentType + ";base64," + b64;
      rec.bytes = Math.floor((b64.length * 3) / 4);
    } catch (e) {
      rec.error = "media-read-failed:" + (e && e.message ? e.message : String(e));
    }
    return rec;
  }

  /* ------------------------- paragraph extraction ------------------------- */

  function extractParagraph(pNode, imageSink) {
    const pPr = firstChildNamed(pNode, "pPr");
    let style = "normal";
    let bullet = null;
    if (pPr) {
      const pStyle = firstChildNamed(pPr, "pStyle");
      if (pStyle) style = normalizeStyleId(attr(pStyle, "w:val") || attr(pStyle, "val"));
      const numPr = firstChildNamed(pPr, "numPr");
      if (numPr) {
        const ilvlEl = firstChildNamed(numPr, "ilvl");
        const ilvl = ilvlEl ? parseInt(attr(ilvlEl, "w:val") || attr(ilvlEl, "val") || "0", 10) : 0;
        bullet = { ilvl };
      }
    }

    const runs = { text: "", boldRuns: 0, textRuns: 0, italicRuns: 0, maxSize: 0 };
    const images = [];

    // Walk inline content: runs may live inside w:hyperlink / w:smartTag too.
    (function walkRuns(container) {
      for (const child of elementChildren(container)) {
        const ln = localName(child);
        if (ln === "r") {
          processRun(child);
        } else if (ln === "hyperlink" || ln === "smartTag" || ln === "ins") {
          walkRuns(child);
        } else if (ln === "fldSimple") {
          walkRuns(child);
        }
      }
    })(pNode);

    function processRun(rNode) {
      const rPr = firstChildNamed(rNode, "rPr");
      let bold = false, italic = false, sizePt = 0, vertAlign = "";
      if (rPr) {
        const b = firstChildNamed(rPr, "b");
        if (b) {
          const v = attr(b, "w:val") || attr(b, "val");
          bold = !(v === "0" || v === "false" || v === "none");
        }
        const it = firstChildNamed(rPr, "i");
        if (it) {
          const v = attr(it, "w:val") || attr(it, "val");
          italic = !(v === "0" || v === "false" || v === "none");
        }
        const sz = firstChildNamed(rPr, "sz");
        if (sz) sizePt = parseInt(attr(sz, "w:val") || attr(sz, "val") || "0", 10) / 2;
        const va = firstChildNamed(rPr, "vertAlign");
        if (va) vertAlign = attr(va, "w:val") || attr(va, "val") || "";
      }

      for (const child of elementChildren(rNode)) {
        const ln = localName(child);
        if (ln === "t") {
          const txt = child.textContent || "";
          if (txt) {
            runs.textRuns++;
            if (bold) runs.boldRuns++;
            if (italic) runs.italicRuns++;
            if (sizePt > runs.maxSize) runs.maxSize = sizePt;
            if (vertAlign === "superscript") runs.text += toSup(txt);
            else if (vertAlign === "subscript") runs.text += toSub(txt);
            else runs.text += txt;
          }
        } else if (ln === "tab") {
          runs.text += "\t";
        } else if (ln === "br" || ln === "cr") {
          runs.text += "\n";
        } else if (ln === "drawing" || ln === "object") {
          collectImages(child, "drawing");
        } else if (ln === "pict") {
          collectImages(child, "pict");
        }
      }
    }

    function collectImages(node) {
      // Modern: a:blip @r:embed ; dimensions from wp:extent (EMU)
      for (const blip of descendantsNamed(node, "blip")) {
        const relId = attr(blip, "r:embed") || attr(blip, "r:id") || attr(blip, "embed") || attr(blip, "id");
        if (relId) {
          const rec = { relId, descr: "", docPrName: "" };
          const extent = firstChildNamed(node, "extent") || descendantsNamed(node, "extent")[0];
          if (extent) {
            const cx = parseInt(attr(extent, "cx") || "0", 10);
            const cy = parseInt(attr(extent, "cy") || "0", 10);
            if (cx) rec.widthPx = Math.round(cx / 9525);
            if (cy) rec.heightPx = Math.round(cy / 9525);
          }
          const docPr = descendantsNamed(node, "docPr")[0];
          if (docPr) {
            rec.descr = attr(docPr, "descr") || "";
            rec.docPrName = attr(docPr, "name") || "";
          }
          images.push(rec);
          imageSink.push(rec);
        }
      }
      // Legacy VML: v:imagedata @r:id
      for (const imdata of descendantsNamed(node, "imagedata")) {
        const relId = attr(imdata, "r:id") || attr(imdata, "id");
        if (relId) {
          const rec = { relId, descr: "", docPrName: "" };
          images.push(rec);
          imageSink.push(rec);
        }
      }
    }

    const text = runs.text.replace(/[ \t]+$/g, "");
    return {
      text,
      bold: runs.textRuns > 0 && runs.boldRuns === runs.textRuns,
      italic: runs.textRuns > 0 && runs.italicRuns === runs.textRuns,
      sizePt: runs.maxSize || null,
      style,
      bullet,
      _images: images, // relIds resolved later
    };
  }

  function extractTable(tblNode, imageSink, outParas) {
    // Flatten tables: one IR paragraph per row, cells joined with " | ".
    for (const tr of elementChildren(tblNode)) {
      if (localName(tr) !== "tr") continue;
      const cellTexts = [];
      for (const tc of elementChildren(tr)) {
        if (localName(tc) !== "tc") continue;
        const parts = [];
        for (const p of descendantsNamed(tc, "p")) {
          const info = extractParagraph(p, imageSink);
          if (info._images.length) {
            // keep images attached to a synthetic paragraph of their own
            outParas.push({ text: "", bold: false, italic: false, sizePt: null, style: "normal", bullet: null, _images: info._images, table: true });
          }
          if (info.text.trim()) parts.push(info.text.replace(/\s+/g, " ").trim());
        }
        cellTexts.push(parts.join(" "));
      }
      outParas.push({
        text: cellTexts.join(" | "),
        bold: false, italic: false, sizePt: null,
        style: "tablerow", bullet: null, _images: [], table: true,
      });
    }
  }

  /* ------------------------------ main entry ------------------------------ */

  async function readDocxIR(arrayBuffer, deps, fileName, opts) {
    opts = opts || {};
    const JSZip = deps.JSZip;
    const DOMParserCtor = deps.DOMParser;
    if (!JSZip || !DOMParserCtor) throw new Error("readDocxIR requires { JSZip, DOMParser }");

    const zip = await JSZip.loadAsync(arrayBuffer);
    const docEntry = zip.file("word/document.xml");
    if (!docEntry) throw new Error("Not a valid .docx file (word/document.xml is missing).");
    const docXml = await docEntry.async("string");
    const parser = new DOMParserCtor();
    readRelationships._parser = parser;
    const rels = await readRelationships(zip);
    const doc = parser.parseFromString(docXml, "application/xml");

    const bodies = descendantsNamed(doc, "body");
    const body = bodies[0] || doc.documentElement;

    const imageSink = []; // every {relId,...} placeholder found in order
    const outParas = [];

    (function walkBody(container) {
      for (const child of elementChildren(container)) {
        const ln = localName(child);
        if (ln === "p") {
          outParas.push(extractParagraph(child, imageSink));
        } else if (ln === "tbl") {
          extractTable(child, imageSink, outParas);
        } else if (ln === "sdt") {
          const content = firstChildNamed(child, "sdtContent");
          if (content) walkBody(content);
        }
      }
    })(body);

    // Resolve every referenced image once (dedupe by relId).
    const byRelId = new Map();
    const maxImages = opts.maxImages || 400;
    let imgCount = 0;
    for (const rec of imageSink) {
      if (!byRelId.has(rec.relId)) {
        if (imgCount >= maxImages) {
          byRelId.set(rec.relId, Object.assign({}, rec, { error: "image-limit-exceeded" }));
        } else {
          imgCount++;
          const loaded = await loadImage(zip, rels, rec.relId);
          // merge metadata (descr/dimensions) from the first occurrence
          Object.assign(loaded, { descr: rec.descr, docPrName: rec.docPrName, widthPx: rec.widthPx || loaded.widthPx, heightPx: rec.heightPx || loaded.heightPx });
          byRelId.set(rec.relId, loaded);
        }
      }
    }

    // Attach resolved images to paragraphs; drop internal fields.
    const paragraphs = outParas.map(function (p) {
      const imgs = (p._images || []).map((rec) => {
        const loaded = byRelId.get(rec.relId) || { error: "image-unresolved", name: rec.relId };
        return Object.assign({}, loaded, {
          descr: rec.descr || loaded.descr || "",
          widthPx: rec.widthPx || loaded.widthPx || null,
          heightPx: rec.heightPx || loaded.heightPx || null,
        });
      });
      const out = {
        text: p.text || "",
        bold: !!p.bold,
        italic: !!p.italic,
        sizePt: p.sizePt || null,
        style: p.style || "normal",
        bullet: p.bullet || null,
        images: imgs,
      };
      if (p.table) out.table = true;
      return out;
    });

    return { fileName: fileName || "untitled.docx", paragraphs };
  }

  return { readDocxIR, resolveMediaPath, MIME, _internals: { extractParagraph, toSup, toSub, normalizeStyleId } };
});
