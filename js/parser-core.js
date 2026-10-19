/* ============================================================================
 * parser-core.js — Deterministic exam-document parser (shared: browser + Node)
 * ----------------------------------------------------------------------------
 * Input : one or more "document IR" objects produced by docx-ir.js:
 *   {
 *     fileName: "Week 1.docx",
 *     paragraphs: [
 *       {
 *         text:   "plain text (unicode sup/sub already applied)",
 *         bold:   true|false,          // whole paragraph bold
 *         sizePt: 14|null,             // largest run size in the paragraph
 *         style:  "heading1"|"title"|"normal"|...,   // normalized style id
 *         bullet: null | { ilvl: 0 },  // list paragraph
 *         images: [ { name, contentType, dataUrl?, error? , widthPx?, heightPx? } ]
 *       }, ...
 *     ]
 *   }
 *
 * Output: { documents, sets, questions, warnings, stats }
 *
 * Design rules:
 *  - Deterministic: same input -> same output. No AI, no randomness.
 *  - Faithful: source text is preserved verbatim; nothing is invented.
 *  - Nothing is silently dropped: unparsed content is captured and flagged.
 * ==========================================================================*/
(function (root, factory) {
  if (typeof module === "object" && module.exports) { module.exports = factory(); }
  else { root.ExamParser = factory(); }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  /* ----------------------------- helpers -------------------------------- */

  function normWS(s) { return String(s == null ? "" : s).replace(/\s+/g, " ").trim(); }

  function normForMatch(s) {
    return normWS(s).toLowerCase()
      .replace(/[\u2010-\u2015\u2212]/g, "-")     // dashes -> hyphen
      .replace(/[\u2018\u2019\u02bc]/g, "'")       // quotes
      .replace(/[\u201c\u201d]/g, '"')
      .replace(/\s+/g, " ");
  }

  // Cheap token-set Dice similarity (0..1) used only to decide whether the
  // leading text of a rationale line is an echo of the option text.
  function similarity(a, b) {
    const ta = new Set(normForMatch(a).split(/[^a-z0-9]+/).filter(Boolean));
    const tb = new Set(normForMatch(b).split(/[^a-z0-9]+/).filter(Boolean));
    if (!ta.size || !tb.size) return 0;
    let inter = 0;
    ta.forEach((t) => { if (tb.has(t)) inter++; });
    return (2 * inter) / (ta.size + tb.size);
  }

  // FNV-1a 32-bit string hash (deterministic, dependency-free).
  function fnv1a(str) {
    let h = 0x811c9dc5;
    for (let i = 0; i < str.length; i++) {
      h ^= str.charCodeAt(i);
      h = Math.imul(h, 0x01000193) >>> 0;
    }
    return h >>> 0;
  }
  function hashText(str) {
    // Two rounds -> 64-bit-ish hex string, stable across platforms.
    const h1 = fnv1a(str).toString(16).padStart(8, "0");
    const h2 = fnv1a("\u0001" + str).toString(16).padStart(8, "0");
    return h1 + h2;
  }

  function slug(s, max) {
    max = max || 40;
    return normWS(s).toLowerCase()
      .normalize("NFKD").replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .slice(0, max)
      .replace(/-+$/g, "");
  }

  function stripFileExt(name) {
    return String(name || "").replace(/\.[a-z0-9]+$/i, "");
  }

  /* --------------------------- line classifiers --------------------------- */

  const RE_QHEAD = /^\s*(?:Question|Q)\.?\s*(\d+)\s*[:.)\-–—]?\s*(.*)$/i;

  const SET_HEAD_PATTERNS = [
    /^\s*(?:Examination|Exam)\s+(\d+)(?:\s+of\s+\d+\s*\+?)?\s*[-–—:.]\s*(.+)$/i,
    /^\s*(?:Examination|Exam)\s+(\d+)(?:\s+of\s+\d+\s*\+?)?\s*$/i,
    /^\s*(?:Form|Set)\s+([A-Za-z0-9]+)\s*[-–—:.]?\s*(.*)$/i,
    /^\s*Part\s+([IVXLCDM]+|\d+)\s*[-–—:.]\s*(.+)$/i,
    /^\s*Section\s+([IVXLCDM]+|\d+|[A-Z])\s*[-–—:.]\s*(.+)$/i,
    /^\s*Practice\s+(?:Exams?|Examinations?|Tests?)\s*(?:(\d+)\s*)?[-–—:.]\s*(.+)$/i,
    /^\s*Unit\s+(\d+|[A-Z])\s*[-–—:.]\s*(.+)$/i,
    /^\s*Week\s+(\d+)\s*[-–—:.]\s*(.+)$/i,
  ];

  const RE_KEY = /^\s*(?:Correct\s+Answers?|Correct\s+Options?|Answer\s+Key|Answers?)\s*[:\-–—]\s*(.+)$/i;
  const RE_KEY_LABELS = /^\(?[A-J]\)?(?:\s*(?:,|;|and|&|\+)\s*\(?\s*[A-J]\)?)*\.?\s*$/i;
  const RE_RATHDR = /^\s*(?:Clinical\s+Rationale|Rationales?|Answer\s+Rationale|Explanations?|Why\s+each\s+option\s+is\s+(?:in)?correct)\s*[:\-–—]?\s*$/i;
  const RE_SRCMAP = /^\s*Source\s+(?:Mapping|Map|References?)\s*[:\-–—]\s*(.*)$/i;
  const RE_OPTION = /^\s*\(?([A-J])\)\s+(.+)$/;
  const RE_CAPTION = /^\s*(?:Figure|Photo|Photograph|Image|Diagram|Exhibit|Chart|Graphic)\s*[A-Za-z0-9]{0,10}\s*[:\-–—]\s*(.+)$/i;
  const RE_OPEN_ANSWER = /^\s*(?:Answer|Answer\s+Key|Model\s+Answer|Faculty\s+Answer)\s*[:\-–—]?\s*$/i;
  const RE_RAT_BULLET = /^[\s•◦‣·∙*\-–—]+/;
  const RE_RAT_CORRECT_MARK = /^\((?:the\s+)?(?:correct|incorrect|right|wrong)(?:\s+answer(?:s)?)?\)\s*[:\-–—]?\s*/i;
  const RE_RAT_LABEL = /^\(?([A-J])\)?\s*[:\)–—-]?\s*(.*)$/;

  const HEADING_STYLE = /(?:^|[^a-z])heading\s*[1-4]|^title$|^subtitle$/i;

  function classifyParagraph(p, ctx) {
    // ctx: { sawOptions, sawKey, sawRatHdr, sawSrcMap, sawOpenAnswer, lastWasImage, type }
    if (p.images && p.images.length) return "IMAGE";
    const t = normWS(p.text);
    if (!t) return "BLANK";

    if (!ctx.sawKey && RE_KEY.test(t)) {
      const m = t.match(RE_KEY);
      const val = m[1].trim();
      if (RE_KEY_LABELS.test(val)) return "KEY";
      if (ctx.type === "tf" && /^(true|false)\b\.?$/i.test(val)) return "KEY";
    }
    if (!ctx.sawRatHdr && RE_RATHDR.test(t)) return "RATHDR";
    if (RE_SRCMAP.test(t)) return "SRCMAP";
    if (!ctx.sawOpenAnswer && !ctx.sawOptions && RE_OPEN_ANSWER.test(t) &&
        (p.bold || t.toLowerCase() === "answer")) return "OPENANSWER";
    if (ctx.lastWasImage && RE_CAPTION.test(t)) return "CAPTION";
    if (RE_OPTION.test(t)) return "OPTION";
    return "TEXT";
  }

  /* ------------------------- rationale line parsing ----------------------- */

  // Parse one rationale-section line into { label, rationale, general } etc.
  function parseRationaleLine(line, options, prevLabel) {
    let s = normWS(line);
    s = s.replace(RE_RAT_BULLET, "");
    s = s.replace(RE_RAT_CORRECT_MARK, "");
    s = normWS(s);
    if (!s) return null;

    const m = s.match(/^\(?([A-J])\)\s*[:\-–—]?\s*(.*)$/);
    if (m) {
      const label = m[1].toUpperCase();
      let rest = normWS(m[2]);
      const opt = options.find((o) => o.label === label);
      if (opt) {
        let rationale = null;
        if (rest.startsWith(":")) {
          rationale = normWS(rest.slice(1));
        } else {
          const nRest = normForMatch(rest);
          const nOpt = normForMatch(opt.text);
          if (nOpt && nRest.startsWith(nOpt)) {
            rationale = normWS(rest.slice(findEchoEnd(rest, opt.text)));
            rationale = rationale.replace(/^[:\-–—]\s*/, "");
          } else {
            // Try "echo: rationale" split at first colon when the prefix
            // strongly resembles the option text.
            const ci = rest.search(/:\s/);
            if (ci > 0 && ci < 400) {
              const prefix = rest.slice(0, ci);
              if (similarity(prefix, opt.text) >= 0.6) rationale = normWS(rest.slice(ci + 1));
            }
            if (rationale === null) {
              // "A): rationale" style already handled; otherwise treat the
              // remainder (minus a leading colon) as the rationale.
              rationale = normWS(rest.replace(/^[:\-–—]\s*/, ""));
            }
          }
        }
        return { label, rationale: rationale || "" };
      }
      // Label not among options: keep as general note (do not lose text).
      return { label: null, general: s, nearLabel: label };
    }
    return { label: null, general: s, continuationOf: isStandaloneNote(s) ? null : (prevLabel || null) };
  }

  // Lines like "Note: ...", "Remember: ..." are standalone notes rather than
  // continuations of the previous option's rationale.
  function isStandaloneNote(s) {
    return /^(?:note|nb|comment|tip|remember|important|caution|warning|safety|teaching point|critical|key point|clinical pearl)\b/i.test(s);
  }

  // Number of characters of `rest` that echo `optText` (whitespace tolerant).
  function findEchoEnd(rest, optText) {
    // Walk both strings ignoring whitespace differences.
    let i = 0, j = 0;
    const a = rest, b = normForMatch(optText);
    const aN = normForMatch(rest);
    while (j < b.length && i < aN.length) {
      if (aN[i] !== b[j]) {
        // tolerate small mismatch only if remainder still lines up; else stop
        break;
      }
      i++; j++;
    }
    if (j >= b.length) {
      // Map normalized index back to raw index in `rest`.
      let raw = 0, norm = 0;
      const normChars = normForMatch(rest);
      while (norm < i && raw < rest.length) {
        const ch = rest[raw];
        if (/\s/.test(ch)) {
          // normalized collapses whitespace to single spaces
          if (norm < normChars.length && normChars[norm] === " ") { norm++; }
          raw++;
          while (raw < rest.length && /\s/.test(rest[raw])) raw++;
          continue;
        }
        norm++; raw++;
      }
      return raw;
    }
    return -1; // no full echo
  }

  /* ------------------------------ type mapping ---------------------------- */

  function mapTypeFromSuffix(suffix) {
    const s = normWS(suffix).toLowerCase();
    if (/select\s*all|\bs\.?a\.?t\.?a\b/i.test(s)) return "sata";
    if (/true\s*[\/\-]?\s*(?:or\s*)?false/i.test(s)) return "tf";
    if (/multiple\s*choice|\bmc\b|single\s*best\s*answer/i.test(s)) return "mc";
    return null; // suffix is a title, not a type
  }

  /* --------------------------- main parse routine ------------------------- */

  function parseExamDocuments(irs, opts) {
    opts = opts || {};
    const warnings = [];
    const documents = [];
    const sets = [];
    const questions = [];

    function warn(level, code, message, ref) {
      warnings.push(Object.assign({ level, code, message }, ref || {}));
    }

    irs.forEach(function (ir) {
      const fileResult = parseOneDocument(ir, warn);
      documents.push(fileResult.document);
      fileResult.sets.forEach((s) => sets.push(s));
      fileResult.questions.forEach((q) => questions.push(q));
    });

    // Cross-document duplicate question detection (informational).
    detectDuplicates(questions, warn);

    const stats = {
      files: irs.length,
      sets: sets.length,
      questions: questions.length,
      byType: questions.reduce((acc, q) => { acc[q.type] = (acc[q.type] || 0) + 1; return acc; }, {}),
      withImages: questions.filter((q) => (q.images || []).length).length,
      warningCount: warnings.length,
      errorCount: warnings.filter((w) => w.level === "error").length,
    };

    return { ok: stats.errorCount === 0, documents, sets, questions, warnings, stats };
  }

  function parseOneDocument(ir, warn) {
    const fileName = ir.fileName || "untitled.docx";
    const fileSlug = slug(stripFileExt(fileName), 48) || "doc";
    const paras = ir.paragraphs || [];

    // ---- Document front matter & title ------------------------------------
    let docTitle = "";
    const docDescParts = [];
    const firstQIdx = paras.findIndex((p) => RE_QHEAD.test(normWS(p.text)));

    // Front matter = paragraphs before first question heading / set heading.
    let frontEnd = firstQIdx === -1 ? paras.length : firstQIdx;
    for (let i = 0; i < frontEnd; i++) {
      if (looksLikeSetHeading(paras[i], paras, i, warn, fileName)) { frontEnd = i; break; }
    }
    for (let i = 0; i < Math.min(frontEnd, 12); i++) {
      const t = normWS(paras[i].text);
      if (!t) continue;
      if (!docTitle && (HEADING_STYLE.test(paras[i].style || "") || paras[i].bold || i === 0)) docTitle = t;
      else docDescParts.push(t);
    }
    if (!docTitle) docTitle = stripFileExt(fileName);
    const fileLabel = docTitle || stripFileExt(fileName);

    // ---- Walk paragraphs with a state machine ------------------------------
    const fileSets = [];
    const fileQuestions = [];
    let currentSet = null;
    let currentQ = null;
    let setCounter = 0;
    let lastQNumInSet = 0;
    let lastWasImageWhere = null; // 'stem' | 'rationale'

    function ensureSet(explicitHeadingPara, detectedBy, confidence) {
      setCounter++;
      const headingText = explicitHeadingPara ? normWS(explicitHeadingPara.text) : "";
      const headSlug = slug(headingText, 30);
      const id = fileSlug + "--s" + setCounter + (headSlug ? "-" + headSlug : "");
      currentSet = {
        id,
        title: headingText || (fileLabel + " — Set " + setCounter),
        sourceFile: fileName,
        documentTitle: docTitle,
        description: "",
        questionIds: [],
        metadata: {
          setIndex: setCounter,
          headingText: headingText,
          detectedBy: detectedBy || [],
          confidence: confidence || "high",
          autoTitle: !headingText,
        },
      };
      fileSets.push(currentSet);
      lastQNumInSet = 0;
      return currentSet;
    }

    function closeQuestion() {
      if (!currentQ) return;
      finalizeQuestion(currentQ, warn, fileName, currentSet ? currentSet.id : null);
      if (currentSet) currentSet.questionIds.push(currentQ.id);
      fileQuestions.push(currentQ);
      currentQ = null;
    }

    function startQuestion(num, suffix, paraIdx) {
      // Numbering-restart signal => implicit set boundary (flagged for review)
      if (currentSet && num <= lastQNumInSet) {
        closeQuestion();
        warn("warn", "numbering-restart",
          "Question numbering restarted at " + num + " without an explicit set heading. A new set was created automatically — please review the boundary.",
          { file: fileName, atParagraph: paraIdx });
        ensureSet(null, ["numbering-restart"], "review");
      }
      closeQuestion();
      if (!currentSet) ensureSet(null, ["implicit-first"], currentSetAutoConfidence(paras, paraIdx));
      lastQNumInSet = num;

      const type = mapTypeFromSuffix(suffix);
      const title = type ? null : normWS(suffix) || null;
      const ordinal = currentSet.questionIds.length + fileQuestions.filter((q) => q.setId === (currentSet && currentSet.id)).length + 1;
      currentQ = {
        id: currentSet.id + "--q" + String(ordinal).padStart(3, "0"),
        setId: currentSet.id,
        sourceFile: fileName,
        questionNumber: num,
        ordinal: currentSet.questionIds.length + 1,
        type: type || null, // resolved in finalizeQuestion if null
        title: title,
        stem: [],
        images: [],
        options: [],
        correctAnswers: [],
        answerBlock: null,
        sourceMapping: "",
        rationaleNotes: [],
        trailingNotes: [],
        issues: [],
        metadata: { headingSuffix: normWS(suffix) },
      };
      // parse context
      currentQ._ctx = {
        section: "stem", // stem|options|postkey|rationale|srcmap|openanswer|terminal
        sawOptions: false, sawKey: false, sawRatHdr: false, sawSrcMap: false, sawOpenAnswer: false,
        lastWasImage: false, type: type,
      };
    }

    function currentOrdinalForSet() {
      return currentSet.questionIds.length + 1;
    }

    for (let i = 0; i < paras.length; i++) {
      const p = paras[i];
      const t = normWS(p.text);

      // --- question heading?
      const qm = t.match(RE_QHEAD);
      if (qm && !(currentQ && currentQ._ctx.section === "openanswer" && false)) {
        // Guard: inside an open-answer block, a line like "Question ..." cannot
        // occur in practice; treat every match as a heading.
        startQuestion(parseInt(qm[1], 10), qm[2] || "", i);
        // recompute ordinal/id properly now that set is known
        currentQ.ordinal = currentOrdinalForSet();
        currentQ.id = currentSet.id + "--q" + String(currentQ.ordinal).padStart(3, "0");
        lastWasImageWhere = null;
        continue;
      }

      // --- set heading?
      if (looksLikeSetHeading(p, paras, i, warn, fileName) &&
          (!currentQ || isTerminalZone(currentQ._ctx.section))) {
        closeQuestion();
        ensureSet(p, ["heading"], "high");
        // description = following non-question paragraphs until first question
        const desc = [];
        for (let j = i + 1; j < Math.min(i + 6, paras.length); j++) {
          const tj = normWS(paras[j].text);
          if (!tj) continue;
          if (RE_QHEAD.test(tj) || looksLikeSetHeading(paras[j], paras, j, null, fileName)) break;
          desc.push(tj);
        }
        currentSet.description = desc.join(" ");
        lastWasImageWhere = null;
        continue;
      }

      // --- front matter (before any set/question)
      if (!currentQ && !currentSet) { continue; } // already captured above
      if (!currentQ && currentSet) {
        // paragraphs between set heading and first question: description handled
        continue;
      }

      // --- inside a question block
      const ctx = currentQ._ctx;
      const kind = classifyParagraph(p, {
        sawOptions: ctx.sawOptions, sawKey: ctx.sawKey, sawRatHdr: ctx.sawRatHdr,
        sawSrcMap: ctx.sawSrcMap, sawOpenAnswer: ctx.sawOpenAnswer,
        lastWasImage: ctx.lastWasImage, type: ctx.type,
      });

      switch (kind) {
        case "BLANK":
          ctx.lastWasImage = false;
          break;

        case "IMAGE": {
          const pos = (ctx.section === "rationale" || ctx.section === "srcmap" || ctx.section === "terminal" || ctx.section === "postkey")
            ? "rationale" : "stem";
          p.images.forEach(function (img) {
            const rec = {
              src: img.dataUrl || (img.assetPath || ""),
              alt: img.descr ? normWS(img.descr) : "",
              caption: "",
              position: pos,
              sourceName: img.name || "",
              contentType: img.contentType || "",
              widthPx: img.widthPx || null,
              heightPx: img.heightPx || null,
            };
            if (img.error) { rec.error = img.error; currentQ.issues.push("image-extraction-failed:" + (img.name || "?")); }
            currentQ.images.push(rec);
            ctx.lastWasImage = true;
            lastWasImageWhere = pos;
          });
          if (!p.images.length) { ctx.lastWasImage = false; }
          break;
        }

        case "CAPTION": {
          const cap = normWS(t);
          const imgs = currentQ.images.filter((im) => !im.caption);
          const target = imgs.length ? imgs[imgs.length - 1] : null;
          if (target && ctx.lastWasImage) {
            target.caption = cap;
            target.alt = cap;
          } else if (target && lastWasImageWhere) {
            target.caption = cap; target.alt = cap;
          } else {            // Caption-like line with no adjacent image: keep as note.
            if (ctx.section === "rationale" || ctx.section === "postkey" || ctx.section === "srcmap") currentQ.rationaleNotes.push(cap);
            else currentQ.stem.push(cap);
          }
          ctx.lastWasImage = false;
          break;
        }

        case "OPTION": {
          if (ctx.section === "rationale" || ctx.section === "postkey" || ctx.section === "srcmap" || ctx.section === "terminal") {
            // Option-like line inside the rationale zone = rationale entry.
            handleRationaleLine(currentQ, t, warn);
            break;
          }
          if (ctx.section === "openanswer") {
            currentQ.answerBlock.push({ text: t, bold: !!p.bold, bullet: p.bullet ? p.bullet.ilvl : null });
            break;
          }
          const om = t.match(RE_OPTION);
          const label = om[1].toUpperCase();
          const text = normWS(om[2]);
          const existing = currentQ.options.find((o) => o.label === label);
          if (existing) {
            currentQ.issues.push("duplicate-option-label:" + label);
          } else {
            currentQ.options.push({ label, text, isCorrect: false, rationale: "" });
          }
          ctx.section = "options";
          ctx.sawOptions = true;
          ctx.lastWasImage = false;
          break;
        }

        case "KEY": {
          const km = t.match(RE_KEY);
          const val = km[1].trim();
          ctx.keyRaw = val;
          ctx.sawKey = true;
          ctx.section = "postkey";
          ctx.lastWasImage = false;
          break;
        }

        case "RATHDR":
          ctx.sawRatHdr = true;
          ctx.section = "rationale";
          ctx.lastWasImage = false;
          break;

        case "SRCMAP": {
          const sm = t.match(RE_SRCMAP);
          currentQ.sourceMapping = normWS(sm[1] || "");
          ctx.sawSrcMap = true;
          ctx.section = "srcmap";
          ctx.lastWasImage = false;
          break;
        }

        case "OPENANSWER":
          ctx.sawOpenAnswer = true;
          ctx.section = "openanswer";
          currentQ.answerBlock = [];
          ctx.lastWasImage = false;
          break;

        case "TEXT":
        default: {
          ctx.lastWasImage = false;
          switch (ctx.section) {
            case "stem":
              currentQ.stem.push(t);
              break;
            case "options": {
              // Continuation paragraph of the previous option.
              const last = currentQ.options[currentQ.options.length - 1];
              if (last) {
                last.text = normWS(last.text + " " + t);
                currentQ.issues.push("option-continuation-merged:" + last.label);
              } else {
                currentQ.stem.push(t);
              }
              break;
            }
            case "postkey":
            case "rationale":
              handleRationaleLine(currentQ, t, warn);
              if (ctx.section === "postkey") ctx.section = "rationale";
              break;
            case "srcmap":
              // additional source-mapping continuation or trailing note
              if (currentQ.sourceMapping && !currentQ._srcMapClosed) {
                currentQ.sourceMapping = normWS(currentQ.sourceMapping + " " + t);
              } else {
                currentQ.trailingNotes.push(t);
              }
              break;
            case "openanswer":
              currentQ.answerBlock.push({ text: t, bold: !!p.bold, bullet: p.bullet ? p.bullet.ilvl : null });
              break;
            case "terminal":
              currentQ.trailingNotes.push(t);
              break;
          }
          break;
        }
      }
    }
    closeQuestion();

    // ---- Document record ---------------------------------------------------
    const document = {
      file: fileName,
      title: docTitle,
      description: docDescParts.join(" ").slice(0, 600),
      importedAt: new Date().toISOString(),
      paragraphCount: paras.length,
      setIds: fileSets.map((s) => s.id),
    };

    // ---- Post-pass: set descriptions for implicit sets ---------------------
    fileSets.forEach(function (s) {
      if (s.metadata.autoTitle) {
        s.title = fileLabel + " — Set " + s.metadata.setIndex;
      }
    });

    return { document, sets: fileSets, questions: fileQuestions };
  }

  function currentSetAutoConfidence(paras, paraIdx) {
    // Questions appearing before any heading: low-confidence implicit set.
    return "review";
  }

  function isTerminalZone(section) {
    return section === "srcmap" || section === "terminal" || section === "postkey" ||
      section === "rationale" || section === "openanswer";
  }

  function looksLikeSetHeading(p, paras, idx, warn, fileName) {
    const t = normWS(p.text);
    if (!t || t.length > 300) return false;
    let matched = null;
    for (let k = 0; k < SET_HEAD_PATTERNS.length; k++) {
      const m = t.match(SET_HEAD_PATTERNS[k]);
      if (m) { matched = m; break; }
    }
    if (!matched) return false;
    // Do not treat question headings as set headings.
    if (RE_QHEAD.test(t) && /^Question\s+\d+/i.test(t)) return false;

    const styleSignal = HEADING_STYLE.test(p.style || "");
    const boldSignal = !!p.bold;
    const sizeSignal = (p.sizePt || 0) >= 13;
    // Proximity signal: a Question heading appears within the next 8 paragraphs.
    let prox = false;
    for (let j = idx + 1; j < Math.min(idx + 9, paras.length); j++) {
      if (RE_QHEAD.test(normWS(paras[j].text))) { prox = true; break; }
      if (paras[j].images && paras[j].images.length) break;
    }
    if (styleSignal || boldSignal || sizeSignal) return true;
    if (prox) return true;
    return false;
  }

  /* ---------------------- rationale line integration ---------------------- */

  function handleRationaleLine(q, t, warn) {
    const lastLabel = q._lastRationaleLabel || null;
    const res = parseRationaleLine(t, q.options, lastLabel);
    if (!res) return;
    if (res.label) {
      const opt = q.options.find((o) => o.label === res.label);
      if (opt) {
        opt.rationale = opt.rationale ? normWS(opt.rationale + " " + res.rationale) : res.rationale;
        q._lastRationaleLabel = res.label;
      } else {
        q.rationaleNotes.push(t);
        q._lastRationaleLabel = null;
      }
    } else if (res.general) {
      if (res.continuationOf && q._lastRationaleLabel) {
        const opt = q.options.find((o) => o.label === q._lastRationaleLabel);
        if (opt) { opt.rationale = normWS(opt.rationale + " " + res.general); return; }
      }
      q.rationaleNotes.push(res.general);
    }
  }

  /* --------------------------- question finalize -------------------------- */

  function finalizeQuestion(q, warn, fileName, setId) {
    const ctx = q._ctx;
    const issues = q.issues;

    // -- resolve type
    if (!q.type) {
      if (q.options.length >= 2) q.type = "mc";
      else if (q.answerBlock && q.answerBlock.length) q.type = "open";
      else q.type = "mc";
    }
    if (q.type === "tf" && q.options.length === 0) {
      // Synthesize True/False options (never invent clinical content).
      q.options = [
        { label: "A", text: "True", isCorrect: false, rationale: "" },
        { label: "B", text: "False", isCorrect: false, rationale: "" },
      ];
      if (ctx.keyRaw && /^(true|false)$/i.test(normWS(ctx.keyRaw))) {
        const isTrue = /^true$/i.test(normWS(ctx.keyRaw));
        q.options[0].isCorrect = isTrue;
        q.options[1].isCorrect = !isTrue;
        q.correctAnswers = [isTrue ? "A" : "B"];
        ctx.keyRaw = null;
      }
      q.type = "tf";
    }

    // -- answer key
    if (ctx.keyRaw && q.type !== "open") {
      const labels = extractKeyLabels(ctx.keyRaw);
      if (labels.length) {
        q.correctAnswers = labels;
      } else {
        issues.push("unparseable-answer-key");
        q.metadata.keyRaw = ctx.keyRaw;
      }
    }

    // -- validate key vs options
    if (q.type === "open") {
      if (!q.answerBlock || !q.answerBlock.length) issues.push("missing-answer-block");
      q.correctAnswers = [];
      q.options = [];
    } else {
      if (!q.options.length) issues.push("missing-options");
      if (!q.correctAnswers.length) issues.push("missing-key");
      const labelSet = new Set(q.options.map((o) => o.label));
      const bad = q.correctAnswers.filter((L) => !labelSet.has(L));
      if (bad.length) issues.push("key-label-not-in-options:" + bad.join(""));
      // option label sequence check
      const expected = q.options.map((_, k) => String.fromCharCode(65 + k));
      const got = q.options.map((o) => o.label);
      if (got.join("") !== expected.join("")) issues.push("option-sequence:" + got.join(""));
      // type vs key-count consistency
      if (q.type === "mc" && q.correctAnswers.length > 1) {
        q.type = "sata";
        issues.push("type-corrected-to-sata");
      }
      if (q.type === "sata" && q.correctAnswers.length === 1) {
        issues.push("sata-single-key");
      }
      // mark isCorrect
      q.options.forEach((o) => { o.isCorrect = q.correctAnswers.indexOf(o.label) !== -1; });
      // per-option rationale completeness
      const noRat = q.options.filter((o) => !o.rationale).map((o) => o.label);
      if (noRat.length && q.options.length) issues.push("missing-rationale:" + noRat.join(""));
      if (!ctx.sawRatHdr && !q.rationaleNotes.length && q.options.some((o) => o.rationale)) {
        // rationales found without a header: fine, but note it
      }
    }

    // -- stem sanity
    if (!q.stem.length && !q.images.some((im) => im.position === "stem")) issues.push("missing-stem");

    // -- content hash (duplicate detection / stable identity)
    q.metadata.contentHash = hashText(
      q.stem.join(" ").toLowerCase() + "|" +
      q.options.map((o) => o.label + o.text.toLowerCase()).join("|") + "|" +
      q.correctAnswers.join("") + "|" +
      (q.images || []).map((im) => im.sourceName || im.src.slice(0, 64)).join(",") + "|" +
      (q.answerBlock || []).map((b) => b.text.toLowerCase()).join(" ")
    );

    // -- warnings for serious issues
    issues.forEach(function (iss) {
      const code = iss.split(":")[0];
      const level = (code === "missing-key" || code === "missing-options" || code === "missing-stem" ||
        code === "key-label-not-in-options" || code === "unparseable-answer-key" ||
        code === "duplicate-option-label" || code.startsWith("image-extraction-failed")) ? "error"
        : (code.startsWith("option-continuation") ? "info" : "warn");
      warn(level, code, describeIssue(iss, q), { file: fileName, setId, questionNumber: q.questionNumber, questionId: q.id });
    });

    if (q.trailingNotes.length) {
      warn("info", "trailing-content", "Question " + q.questionNumber + ": " + q.trailingNotes.length +
        " trailing paragraph(s) were preserved in 'trailingNotes'.", { file: fileName, setId, questionId: q.id });
    }

    // cleanup internal fields
    delete q._ctx;
    delete q._lastRationaleLabel;
    q.answerBlock = q.answerBlock && q.answerBlock.length ? q.answerBlock : (q.type === "open" ? q.answerBlock : null);
    if (!q.answerBlock) delete q.answerBlock;
    if (!q.rationaleNotes.length) delete q.rationaleNotes;
    if (!q.trailingNotes.length) delete q.trailingNotes;
    if (!q.images.length) delete q.images;
    if (!q.issues.length) delete q.issues;
    if (!q.title) delete q.title;
    if (!q.sourceMapping) delete q.sourceMapping;
  }

  function extractKeyLabels(raw) {
    const s = normWS(raw).toUpperCase();
    if (RE_KEY_LABELS.test(raw.trim())) {
      const labels = s.match(/[A-J]/g) || [];
      // dedupe preserving order
      return labels.filter((L, i) => labels.indexOf(L) === i);
    }
    return [];
  }

  function describeIssue(iss, q) {
    const code = iss.split(":")[0];
    const arg = iss.slice(code.length + 1);
    const where = "Question " + q.questionNumber + (q.title ? " (" + q.title + ")" : "");
    switch (code) {
      case "missing-key": return where + " has no detectable answer key.";
      case "missing-options": return where + " has no answer options and no open-answer block.";
      case "missing-stem": return where + " has no question text (stem) and no stem image.";
      case "unparseable-answer-key": return where + " answer key could not be parsed (\"" + (q.metadata.keyRaw || "") + "\").";
      case "key-label-not-in-options": return where + " answer key references option(s) " + arg + " that do not exist.";
      case "duplicate-option-label": return where + " has duplicate option label " + arg + ".";
      case "option-sequence": return where + " option labels are not a clean A..N sequence (found: " + arg + ").";
      case "type-corrected-to-sata": return where + " was labeled Multiple Choice but has multiple correct answers; reclassified as Select All That Apply.";
      case "sata-single-key": return where + " is labeled Select All That Apply but the key lists a single answer.";
      case "missing-rationale": return where + " has no rationale for option(s) " + arg + ".";
      case "option-continuation-merged": return where + ": extra paragraph merged into option " + arg + ".";
      case "image-extraction-failed": return where + ": an embedded image could not be extracted (" + arg + ").";
      case "missing-answer-block": return where + " (open-ended) has no answer/explanation block.";
      default: return where + ": " + iss;
    }
  }

  /* ------------------------- duplicate detection -------------------------- */

  function detectDuplicates(questions, warn) {
    const byHash = new Map();
    questions.forEach(function (q) {
      const h = q.metadata && q.metadata.contentHash;
      if (!h) return;
      if (!byHash.has(h)) byHash.set(h, []);
      byHash.get(h).push(q);
    });
    byHash.forEach(function (group) {
      if (group.length > 1) {
        warn("warn", "duplicate-question",
          "Identical question content appears " + group.length + " times: " +
          group.map((q) => {
            const sm = String(q.setId || "").match(/--s(\d+)/);
            return (q.sourceFile ? q.sourceFile.replace(/\.docx$/i, "") + " " : "") +
              "[set " + (sm ? sm[1] : "?") + ", Q" + q.questionNumber + "]";
          }).join("; ") +
          ". (Parallel forms often repeat items intentionally — review before deleting.)",
          { questionIds: group.map((q) => q.id) });
      }
    });
  }

  /* --------------------------- bank assembly ------------------------------ */

  function assembleBank(parseResult, opts) {
    opts = opts || {};
    const bankId = "bank-" + hashText(
      parseResult.sets.map((s) => s.id + ":" + s.questionIds.length).join("|") + "#" +
      parseResult.questions.map((q) => q.metadata.contentHash).join("|")
    ).slice(0, 12);

    const bank = {
      format: "exam-bank",
      version: 1,
      bankId,
      title: opts.title || "Practice Examination Bank",
      generatedAt: opts.generatedAt || new Date().toISOString(),
      generator: "exam-practice-platform/parser-core v1",
      documents: parseResult.documents,
      sets: parseResult.sets,
      questions: parseResult.questions,
    };
    return bank;
  }

  /* ------------------------------- validator ------------------------------ */

  function validateBank(bank) {
    const errors = [];
    const warnings = [];
    if (!bank || typeof bank !== "object") { errors.push("Bank is not an object."); return { valid: false, errors, warnings }; }
    if (bank.format !== "exam-bank") errors.push("Unrecognized format field (expected \"exam-bank\").");
    if (typeof bank.version !== "number") errors.push("Missing numeric 'version'.");
    if (!Array.isArray(bank.sets)) errors.push("Missing 'sets' array.");
    if (!Array.isArray(bank.questions)) errors.push("Missing 'questions' array.");
    if (errors.length) return { valid: false, errors, warnings };

    const qIds = new Set();
    const setIds = new Set();
    const qById = new Map();

    bank.sets.forEach(function (s) {
      if (!s.id) errors.push("Set without id: " + JSON.stringify(s).slice(0, 80));
      else if (setIds.has(s.id)) errors.push("Duplicate set id: " + s.id);
      setIds.add(s.id);
      if (!s.title) warnings.push("Set " + s.id + " has no title.");
    });

    const TYPES = new Set(["mc", "sata", "tf", "open"]);
    bank.questions.forEach(function (q) {
      if (!q.id) { errors.push("Question without id."); return; }
      if (qById.has(q.id)) errors.push("Duplicate question id: " + q.id);
      qById.set(q.id, q);
      if (!setIds.has(q.setId)) errors.push("Question " + q.id + " references unknown set " + q.setId);
      if (!TYPES.has(q.type)) errors.push("Question " + q.id + " has invalid type '" + q.type + "'.");
      if (!Array.isArray(q.stem) || (!q.stem.length && !(q.images || []).some((im) => im.position === "stem"))) {
        warnings.push("Question " + q.id + " has an empty stem.");
      }
      if (q.type === "open") {
        if (!Array.isArray(q.answerBlock) || !q.answerBlock.length) warnings.push("Open question " + q.id + " has no answer block.");
      } else {
        const labels = (q.options || []).map((o) => o.label);
        if (!labels.length) errors.push("Question " + q.id + " has no options.");
        if (new Set(labels).size !== labels.length) errors.push("Question " + q.id + " has duplicate option labels.");
        const correct = (q.options || []).filter((o) => o.isCorrect).map((o) => o.label);
        const key = q.correctAnswers || [];
        const a = correct.slice().sort().join("");
        const b = key.slice().sort().join("");
        if (a !== b) errors.push("Question " + q.id + ": correctAnswers (" + (b || "none") + ") and option.isCorrect flags (" + (a || "none") + ") disagree.");
        if (!key.length) warnings.push("Question " + q.id + " has no answer key.");
        if (q.type === "mc" && key.length > 1) errors.push("Question " + q.id + " is 'mc' but has multiple correct answers.");
        if (q.type === "sata" && key.length < 2) warnings.push("Question " + q.id + " is 'sata' but has fewer than 2 correct answers.");
      }
      (q.images || []).forEach(function (im, k) {
        if (!im.src) errors.push("Question " + q.id + " image #" + (k + 1) + " has no src.");
      });
    });

    bank.sets.forEach(function (s) {
      (s.questionIds || []).forEach(function (qid) {
        if (!qById.has(qid)) errors.push("Set " + s.id + " references unknown question " + qid);
      });
      const owned = bank.questions.filter((q) => q.setId === s.id).length;
      if ((s.questionIds || []).length !== owned) {
        warnings.push("Set " + s.id + ": questionIds length (" + (s.questionIds || []).length + ") != questions with setId (" + owned + ").");
      }
    });

    return { valid: errors.length === 0, errors, warnings };
  }

  /* ------------------------------ exports --------------------------------- */

  return {
    parseExamDocuments,
    assembleBank,
    validateBank,
    // exported for tests / reuse
    _internals: {
      normWS, normForMatch, similarity, slug, hashText, fnv1a,
      parseRationaleLine, findEchoEnd, mapTypeFromSuffix,
      RE_QHEAD, RE_KEY, RE_OPTION, SET_HEAD_PATTERNS, classifyParagraph,
      extractKeyLabels, looksLikeSetHeading,
    },
  };
});
