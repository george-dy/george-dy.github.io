/* tests/parser.test.mjs — parser-core unit tests (synthetic fixtures) +
 * integration tests against the real .docx uploads (skipped if absent). */
import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import { createRequire } from "node:module";
import { extractIRFromFile } from "../tools/ir-extract.mjs";

const require = createRequire(import.meta.url);
const parser = require("../js/parser-core.js");
const { parseExamDocuments, assembleBank, validateBank, _internals } = parser;

/* ------------------------- IR fixture builders ---------------------------- */
function P(text, opts = {}) {
  return {
    text,
    bold: !!opts.bold,
    italic: false,
    sizePt: opts.sizePt || null,
    style: opts.style || "normal",
    bullet: opts.bullet != null ? { ilvl: opts.bullet } : null,
    images: [],
  };
}
function BLANK() { return P(""); }
function IMG(name = "image1.png", descr = "") {
  return {
    text: "", bold: false, italic: false, sizePt: null, style: "normal", bullet: null,
    images: [{ name, contentType: "image/png", dataUrl: "data:image/png;base64,QUJD", descr, widthPx: 100, heightPx: 50 }],
  };
}
function IR(fileName, paragraphs) { return { fileName, paragraphs }; }
function parse1(paras, fileName = "Week 1.docx") {
  return parseExamDocuments([IR(fileName, paras)]);
}
function q(result, n = 1) {
  return result.questions.find((q) => q.questionNumber === n);
}

/* =============== CASE 1: standard multiple-choice question =============== */
test("case 1: standard MC question parses stem/options/key/rationales/source", () => {
  const r = parse1([
    P("Week 1 Practice Exam", { bold: true }),
    P("Form A - 50 Items", { bold: true }),
    P("Question 1: Multiple Choice", { bold: true }),
    P("A nurse assesses a client. Which finding is normal?"),
    BLANK(),
    P("A) Finding one"),
    P("B) Finding two"),
    P("C) Finding three"),
    P("D) Finding four"),
    BLANK(), BLANK(), BLANK(),
    P("Correct Answer: B"),
    P("Clinical Rationale:"),
    P("• A): One is wrong because X."),
    P("• B): Two is correct because Y."),
    P("• C): Three is wrong because Z."),
    P("• D): Four is wrong because W."),
    P("Source Mapping: Slide 7; Transcript 2"),
  ]);
  assert.equal(r.stats.questions, 1);
  assert.equal(r.stats.errorCount, 0);
  const qq = q(r);
  assert.equal(qq.type, "mc");
  assert.deepEqual(qq.stem, ["A nurse assesses a client. Which finding is normal?"]);
  assert.equal(qq.options.length, 4);
  assert.deepEqual(qq.correctAnswers, ["B"]);
  assert.equal(qq.options[1].isCorrect, true);
  assert.equal(qq.options[0].isCorrect, false);
  assert.match(qq.options[0].rationale, /One is wrong because X\./);
  assert.match(qq.options[1].rationale, /Two is correct because Y\./);
  assert.equal(qq.sourceMapping, "Slide 7; Transcript 2");
  assert.equal(r.sets.length, 1);
  assert.equal(r.sets[0].title, "Form A - 50 Items");
});

/* ============ CASE 2: SATA with several correct answers ================== */
test("case 2: SATA grading data — exact key set, per-option rationales", () => {
  const r = parse1([
    P("Form A - 50 Items", { bold: true }),
    P("Question 1: Select All That Apply", { bold: true }),
    P("Which vaccines are due at 12 months? Select all that apply."),
    P("A) Hep A dose 2"),
    P("B) PCV dose 4"),
    P("C) DTaP dose 5"),
    P("D) Varicella dose 1"),
    P("E) MMR dose 1"),
    BLANK(), BLANK(), BLANK(),
    P("Correct Answer: B, D, E"),
    P("Clinical Rationale:"),
    P("• B): PCV4 completes the series."),
    P("• D): Varicella 1 is due."),
    P("• E): MMR 1 is due."),
    P("• A): Hep A dose 2 is later."),
    P("• C): DTaP 5 is a school-entry booster."),
    P("Source Mapping: Slide 5"),
  ]);
  const qq = q(r);
  assert.equal(qq.type, "sata");
  assert.deepEqual(qq.correctAnswers, ["B", "D", "E"]);
  assert.deepEqual(qq.options.filter((o) => o.isCorrect).map((o) => o.label), ["B", "D", "E"]);
  assert.equal(qq.options.length, 5);
  assert.ok(qq.options.every((o) => o.rationale.length > 0), "every option has rationale");
  assert.equal(r.stats.errorCount, 0);
  assert.ok(!qq.issues, "no issues");
});

/* ============ CASE 3: more than four options ============================= */
test("case 3: question with six options A-F", () => {
  const r = parse1([
    P("Form A - 50 Items", { bold: true }),
    P("Question 1: Multiple Choice", { bold: true }),
    P("Which statement is correct?"),
    P("A) one"), P("B) two"), P("C) three"), P("D) four"), P("E) five"), P("F) six"),
    BLANK(),
    P("Correct Answer: F"),
    P("Clinical Rationale:"),
    P("F) six: correct because reasons."),
    P("A) one: wrong."), P("B) two: wrong."), P("C) three: wrong."), P("D) four: wrong."), P("E) five: wrong."),
    P("Source Mapping: deck"),
  ]);
  const qq = q(r);
  assert.equal(qq.options.length, 6);
  assert.deepEqual(qq.correctAnswers, ["F"]);
  assert.equal(r.stats.errorCount, 0);
});

/* ============ CASE 4: multi-paragraph stem =============================== */
test("case 4: multi-paragraph stem preserved as separate paragraphs", () => {
  const r = parse1([
    P("Form A - 50 Items", { bold: true }),
    P("Question 1: Multiple Choice", { bold: true }),
    P("First stem paragraph with a long clinical scenario."),
    P("Second stem paragraph asking the actual question?"),
    P("A) opt1"), P("B) opt2"), P("C) opt3"), P("D) opt4"),
    P("Correct Answer: C"),
    P("Clinical Rationale:"),
    P("C) opt3: right."), P("A) opt1: wrong."), P("B) opt2: wrong."), P("D) opt4: wrong."),
    P("Source Mapping: x"),
  ]);
  assert.deepEqual(q(r).stem, [
    "First stem paragraph with a long clinical scenario.",
    "Second stem paragraph asking the actual question?",
  ]);
});

/* ============ CASE 5: rationale for every option (echo format) =========== */
test("case 5: 'label) optionText: rationale' echo format splits correctly", () => {
  const r = parse1([
    P("Form A - 50 Items", { bold: true }),
    P("Question 1: Multiple Choice", { bold: true }),
    P("Stem?"),
    P("A) Jaundice appearing within the first 24 hours of life is pathologic until proven otherwise."),
    P("B) Physiologic jaundice; reassure and repeat visual check."),
    BLANK(),
    P("Correct Answer: A"),
    P("Clinical Rationale:"),
    P("A) Jaundice appearing within the first 24 hours of life is pathologic until proven otherwise.: Core course fact: early jaundice needs labs."),
    P("B) Physiologic jaundice; reassure and repeat visual check.: Visual checks are insufficient."),
    P("Source Mapping: Slide 2"),
  ]);
  const qq = q(r);
  assert.equal(qq.options[0].rationale, "Core course fact: early jaundice needs labs.");
  assert.equal(qq.options[1].rationale, "Visual checks are insufficient.");
  assert.equal(r.stats.errorCount, 0);
});

/* ============ CASE 6: missing rationale ================================== */
test("case 6: missing rationale is flagged, not invented", () => {
  const r = parse1([
    P("Form A - 50 Items", { bold: true }),
    P("Question 1: Multiple Choice", { bold: true }),
    P("Stem?"),
    P("A) a"), P("B) b"), P("C) c"), P("D) d"),
    P("Correct Answer: B"),
    P("Clinical Rationale:"),
    P("B) b: because."),
    P("A) a: wrong."),
    P("Source Mapping: x"),
  ]);
  const qq = q(r);
  assert.equal(qq.options[1].rationale, "because.");
  assert.equal(qq.options[2].rationale, "");
  assert.equal(qq.options[3].rationale, "");
  assert.ok(qq.issues.includes("missing-rationale:CD"), "flags missing rationales: " + JSON.stringify(qq.issues));
  assert.ok(r.warnings.some((w) => w.code === "missing-rationale"));
});

/* ============ CASE 7: embedded image ===================================== */
test("case 7: embedded image attaches to the question with caption; rationale-zone image marked separately", () => {
  const r = parse1([
    P("Form A - 50 Items", { bold: true }),
    P("Question 1: Multiple Choice", { bold: true }),
    P("The cropped photograph shows a rash. Which diagnosis?"),
    IMG("photo1.png"),
    P("Figure 1: Clinical photograph - scalp scales."),
    P("A) cradle cap"), P("B) tinea"), P("C) psoriasis"), P("D) eczema"),
    P("Correct Answer: A"),
    P("Clinical Rationale:"),
    P("A) cradle cap: correct."), P("B) tinea: wrong."), P("C) psoriasis: wrong."), P("D) eczema: wrong."),
    IMG("ref.png", "reference chart"),
    P("Figure 2: Reference asset chart."),
    P("Source Mapping: Slide 75"),
  ]);
  const qq = q(r);
  assert.equal(qq.images.length, 2);
  assert.equal(qq.images[0].position, "stem");
  assert.equal(qq.images[0].alt, "Figure 1: Clinical photograph - scalp scales.");
  assert.equal(qq.images[1].position, "rationale");
  assert.equal(qq.images[1].alt, "Figure 2: Reference asset chart.");
  assert.deepEqual(qq.stem, ["The cropped photograph shows a rash. Which diagnosis?"]);
});

/* ============ CASE 8: multiple sets in one document ====================== */
test("case 8: three headed sets in one document, questions split correctly", () => {
  const paras = [P("Practice Exams", { bold: true })];
  for (const [name, count] of [["Form A - 50 Items", 2], ["Form B - 50 Items", 3], ["Form C - 50 Items", 1]]) {
    paras.push(P(name, { bold: true }));
    for (let i = 1; i <= count; i++) {
      paras.push(P(`Question ${i}: Multiple Choice`, { bold: true }));
      paras.push(P(`Stem ${name} ${i}?`));
      paras.push(P("A) a"), P("B) b"), P("C) c"), P("D) d"));
      paras.push(P("Correct Answer: A"));
      paras.push(P("Clinical Rationale:"));
      paras.push(P("A) a: yes."), P("B) b: no."), P("C) c: no."), P("D) d: no."));
      paras.push(P("Source Mapping: x"));
    }
  }
  const r = parse1(paras);
  assert.equal(r.sets.length, 3);
  assert.deepEqual(r.sets.map((s) => s.questionIds.length), [2, 3, 1]);
  assert.equal(r.stats.questions, 6);
  assert.ok(r.sets.every((s) => s.metadata.confidence === "high"));
  // ids unique & stable-looking
  const ids = new Set(r.questions.map((x) => x.id));
  assert.equal(ids.size, 6);
  // each question points at the right set
  r.sets.forEach((s) => s.questionIds.forEach((qid) => {
    const qq = r.questions.find((x) => x.id === qid);
    assert.equal(qq.setId, s.id);
  }));
});

/* ============ CASE 9: numbering restarts without heading ================= */
test("case 9: numbering restart at 1 creates a flagged implicit set boundary", () => {
  const paras = [];
  for (let s = 0; s < 2; s++) {
    for (let i = 1; i <= 2; i++) {
      paras.push(P(`Question ${i}: Multiple Choice`, { bold: true }));
      paras.push(P(`Stem ${s}-${i}?`));
      paras.push(P("A) a"), P("B) b"));
      paras.push(P("Correct Answer: A"));
      paras.push(P("A) a: yes."), P("B) b: no."));
      paras.push(P("Source Mapping: x"));
    }
  }
  const r = parse1(paras);
  assert.equal(r.sets.length, 2, "two sets from restart");
  assert.equal(r.sets[0].metadata.confidence, "review");
  assert.equal(r.sets[1].metadata.detectedBy[0], "numbering-restart");
  assert.ok(r.warnings.some((w) => w.code === "numbering-restart"));
  assert.deepEqual(r.sets.map((s) => s.questionIds.length), [2, 2]);
});

/* ============ CASE 10: numbering continues within one set ================ */
test("case 10: continuous numbering does NOT split the set", () => {
  const paras = [P("Part I - Cases", { style: "heading1" })];
  for (let i = 1; i <= 4; i++) {
    paras.push(P(`Question ${i}. Case ${i} - topic`, { bold: true }));
    paras.push(P("Stem?"));
    paras.push(P("Answer", { bold: true }));
    paras.push(P("Expected diagnosis: something."));
  }
  const r = parse1(paras);
  assert.equal(r.sets.length, 1);
  assert.equal(r.sets[0].questionIds.length, 4);
  assert.equal(r.stats.errorCount, 0);
});

/* ============ CASE 11: invalid / missing answer key ====================== */
test("case 11a: missing answer key is flagged as error, question preserved", () => {
  const r = parse1([
    P("Form A - Items", { bold: true }),
    P("Question 1: Multiple Choice", { bold: true }),
    P("Stem?"),
    P("A) a"), P("B) b"), P("C) c"), P("D) d"),
    P("Clinical Rationale:"),
    P("A) a: yes."),
    P("Source Mapping: x"),
  ]);
  const qq = q(r);
  assert.ok(qq.issues.includes("missing-key"));
  assert.deepEqual(qq.correctAnswers, []);
  assert.equal(r.stats.errorCount >= 1, true);
  assert.ok(r.warnings.some((w) => w.code === "missing-key" && w.level === "error"));
  // stem/options still preserved
  assert.equal(qq.options.length, 4);
});

test("case 11b: key referencing nonexistent option is flagged", () => {
  const r = parse1([
    P("Form A - Items", { bold: true }),
    P("Question 1: Multiple Choice", { bold: true }),
    P("Stem?"),
    P("A) a"), P("B) b"), P("C) c"), P("D) d"),
    P("Correct Answer: E"),
    P("Source Mapping: x"),
  ]);
  const qq = q(r);
  assert.ok(qq.issues.some((i) => i.startsWith("key-label-not-in-options")));
});

test("case 11c: 'Answer source:' line is NOT treated as an answer key", () => {
  const r = parse1([
    P("Part I - Cases", { style: "heading1" }),
    P("Question 1. Case - topic", { bold: true }),
    P("Stem?"),
    P("Answer", { bold: true }),
    P("Answer source: Faculty Answer Key, Card 1."),
    P("Expected diagnosis: Caput succedaneum"),
  ]);
  const qq = q(r);
  assert.equal(qq.type, "open");
  assert.ok(qq.answerBlock.some((b) => b.text.startsWith("Answer source:")));
  assert.equal(r.stats.errorCount, 0);
});

/* ============ CASE 12: inconsistent blank lines / spacing ================ */
test("case 12: inconsistent spacing does not break parsing", () => {
  const r = parse1([
    P("Form A - Items", { bold: true }),
    BLANK(), BLANK(),
    P("Question 1: Multiple Choice", { bold: true }),
    BLANK(),
    P("Stem?"),
    BLANK(), BLANK(),
    P("A) a"),
    BLANK(),
    P("B) b"),
    P("C) c"), BLANK(), BLANK(),
    P("D) d"),
    BLANK(), BLANK(), BLANK(), BLANK(), BLANK(),
    P("Correct Answer: C"),
    BLANK(),
    P("Clinical Rationale:"),
    BLANK(),
    P("C) c: right."),
    BLANK(),
    P("A) a: wrong."),
    P("B) b: wrong."),
    BLANK(),
    P("D) d: wrong."),
    BLANK(), BLANK(),
    P("Source Mapping: x"),
    BLANK(),
  ]);
  const qq = q(r);
  assert.equal(qq.options.length, 4);
  assert.deepEqual(qq.correctAnswers, ["C"]);
  assert.equal(r.stats.errorCount, 0);
  assert.equal(qq.options[2].rationale, "right.");
});

/* ============ EXTRA robustness cases ===================================== */
test("extra: option label inside rationale does not create a new option", () => {
  const r = parse1([
    P("Form A - Items", { bold: true }),
    P("Question 1: Multiple Choice", { bold: true }),
    P("Stem?"),
    P("A) a"), P("B) b"), P("C) c"), P("D) d"),
    P("Correct Answer: A"),
    P("Clinical Rationale:"),
    P("A) a: right."),
    P("Note: some sources list A) differently but the course teaches this."),
    P("B) b: wrong."), P("C) c: wrong."), P("D) d: wrong."),
    P("Source Mapping: x"),
  ]);
  const qq = q(r);
  assert.equal(qq.options.length, 4, "still four options");
  assert.ok((qq.rationaleNotes || []).some((n) => n.startsWith("Note:")), "note preserved as rationale note");
});

test("extra: '(Correct)' marker in rationale line is stripped", () => {
  const res = _internals.parseRationaleLine(
    "• (Correct) D) Assess misuse and diversion: Ongoing assessment is required.",
    [{ label: "D", text: "Assess misuse and diversion" }], null);
  assert.equal(res.label, "D");
  assert.equal(res.rationale, "Ongoing assessment is required.");
});

test("extra: MC heading with multi-letter key is reclassified as SATA", () => {
  const r = parse1([
    P("Form A - Items", { bold: true }),
    P("Question 1: Multiple Choice", { bold: true }),
    P("Stem?"),
    P("A) a"), P("B) b"), P("C) c"), P("D) d"),
    P("Correct Answer: A, C"),
    P("Source Mapping: x"),
  ]);
  const qq = q(r);
  assert.equal(qq.type, "sata");
  assert.ok(qq.issues.includes("type-corrected-to-sata"));
});

test("extra: true/false question gets synthesized options", () => {
  const r = parse1([
    P("Form A - Items", { bold: true }),
    P("Question 1: True/False", { bold: true }),
    P("The statement is correct?"),
    P("Correct Answer: True"),
    P("Source Mapping: x"),
  ]);
  const qq = q(r);
  assert.equal(qq.type, "tf");
  assert.deepEqual(qq.options.map((o) => o.text), ["True", "False"]);
  assert.deepEqual(qq.correctAnswers, ["A"]);
});

test("extra: open question captures structured answer block with bullets/bold", () => {
  const r = parse1([
    P("Part II - Picture Cases", { style: "heading1" }),
    P("Question 16. Picture Case - Color", { bold: true }),
    P("Examine the photograph and describe."),
    IMG("p.png"),
    P("Photo A: newborn supine."),
    P("Answer", { bold: true }),
    P("Answer source: Faculty Answer Key, Card 23."),
    P("Key reasoning", { bold: true }),
    P("Hands and feet blue.", { bullet: 0 }),
    P("Lips pink.", { bullet: 0 }),
  ]);
  const qq = q(r, 16);
  assert.equal(qq.type, "open");
  assert.equal(qq.title, "Picture Case - Color");
  assert.equal(qq.images[0].alt, "Photo A: newborn supine.");
  assert.equal(qq.answerBlock.length, 4);
  assert.equal(qq.answerBlock[2].bullet, 0);
  assert.equal(qq.answerBlock[0].bold, false);
  assert.equal(qq.answerBlock[1].bold, true);   // "Key reasoning" subheading
  assert.equal(qq.answerBlock[1].text, "Key reasoning");
});

test("extra: multi-paragraph option continuation is merged and flagged", () => {
  const r = parse1([
    P("Form A - Items", { bold: true }),
    P("Question 1: Multiple Choice", { bold: true }),
    P("Stem?"),
    P("A) first part of option"),
    P("second part of option A"),
    P("B) b"), P("C) c"), P("D) d"),
    P("Correct Answer: A"),
    P("Source Mapping: x"),
  ]);
  const qq = q(r);
  assert.equal(qq.options.length, 4);
  assert.equal(qq.options[0].text, "first part of option second part of option A");
  assert.ok(qq.issues.some((i) => i.startsWith("option-continuation-merged")));
});

test("extra: set heading without heading style but bold is detected", () => {
  const r = parse1([
    P("Examination 2 of 3 — ADHD Management", { bold: true, sizePt: 14 }),
    P("50 questions - provider level"),
    P("Question 1: Multiple Choice", { bold: true }),
    P("Stem?"),
    P("A) a"), P("B) b"),
    P("Correct Answer: B"),
    P("Source Mapping: x"),
  ]);
  assert.equal(r.sets.length, 1);
  assert.equal(r.sets[0].title, "Examination 2 of 3 — ADHD Management");
  assert.equal(r.sets[0].description, "50 questions - provider level");
  assert.equal(r.sets[0].metadata.confidence, "high");
});

test("extra: ids are stable for identical input (two parses)", () => {
  const mk = () => parse1([
    P("Form A - Items", { bold: true }),
    P("Question 1: Multiple Choice", { bold: true }),
    P("Stem?"), P("A) a"), P("B) b"),
    P("Correct Answer: A"), P("Source Mapping: x"),
  ]);
  const r1 = mk(), r2 = mk();
  assert.equal(r1.sets[0].id, r2.sets[0].id);
  assert.equal(r1.questions[0].id, r2.questions[0].id);
  assert.equal(r1.questions[0].metadata.contentHash, r2.questions[0].metadata.contentHash);
});

test("extra: duplicate questions across sets are detected", () => {
  const paras = [];
  for (const form of ["Form A - Items", "Form B - Items"]) {
    paras.push(P(form, { bold: true }));
    paras.push(P("Question 1: Multiple Choice", { bold: true }));
    paras.push(P("Identical stem?"));
    paras.push(P("A) a"), P("B) b"));
    paras.push(P("Correct Answer: A"));
    paras.push(P("Source Mapping: x"));
  }
  const r = parse1(paras);
  assert.ok(r.warnings.some((w) => w.code === "duplicate-question"));
});

/* -------------------------- validator & bank ----------------------------- */
test("validator: catches id duplicates, key/isCorrect disagreement, bad types", () => {
  const good = parse1([
    P("Form A - Items", { bold: true }),
    P("Question 1: Multiple Choice", { bold: true }),
    P("Stem?"), P("A) a"), P("B) b"),
    P("Correct Answer: A"), P("Source Mapping: x"),
  ]);
  const bank = assembleBank(good);
  let v = validateBank(bank);
  assert.equal(v.valid, true, JSON.stringify(v.errors));

  // corrupt: disagreement
  bank.questions[0].correctAnswers = ["B"];
  v = validateBank(bank);
  assert.equal(v.valid, false);
  assert.ok(v.errors.some((e) => /disagree/.test(e)));

  // corrupt: duplicate id
  bank.questions[0].correctAnswers = ["A"];
  bank.questions.push(JSON.parse(JSON.stringify(bank.questions[0])));
  v = validateBank(bank);
  assert.ok(v.errors.some((e) => /Duplicate question id/.test(e)));
});

test("validator: rejects non-bank JSON gracefully", () => {
  const v = validateBank({ hello: "world" });
  assert.equal(v.valid, false);
  assert.ok(v.errors.length > 0);
  const v2 = validateBank(null);
  assert.equal(v2.valid, false);
});

/* --------------------- integration with real documents -------------------- */
const UPLOADS = new URL("../../uploads/", import.meta.url).pathname;
const hasUploads = fs.existsSync(UPLOADS);

test("integration: doc 3 (Toddler) -> 3 sets x 50, no errors", { skip: !hasUploads }, async () => {
  const ir = await extractIRFromFile(UPLOADS + "3 Toddler and School Age questions.docx");
  const r = parseExamDocuments([ir]);
  assert.equal(r.sets.length, 3);
  assert.deepEqual(r.sets.map((s) => s.questionIds.length), [50, 50, 50]);
  assert.equal(r.stats.byType.mc, 139);
  assert.equal(r.stats.byType.sata, 11);
  assert.equal(r.stats.errorCount, 0);
  assert.equal(r.stats.warningCount, 0);
  const bank = assembleBank(r);
  assert.equal(validateBank(bank).valid, true, JSON.stringify(validateBank(bank).errors.slice(0, 3)));
});

test("integration: doc 4 (ADHD) -> 3 sets x 50, no errors", { skip: !hasUploads }, async () => {
  const ir = await extractIRFromFile(UPLOADS + "4 Developmental_Behavioral_ADHD.docx");
  const r = parseExamDocuments([ir]);
  assert.equal(r.sets.length, 3);
  assert.deepEqual(r.sets.map((s) => s.questionIds.length), [50, 50, 50]);
  assert.equal(r.stats.byType.mc, 144);
  assert.equal(r.stats.byType.sata, 6);
  assert.equal(r.stats.errorCount, 0);
});

test("integration: doc 1 (Newborn/Infant) -> 6 sets (5x50+34), 40 images, no errors", { skip: !hasUploads }, async () => {
  const ir = await extractIRFromFile(UPLOADS + "1 Newborn_Infant_Practice_Examinations.docx");
  const r = parseExamDocuments([ir]);
  assert.equal(r.sets.length, 6);
  assert.deepEqual(r.sets.map((s) => s.questionIds.length), [50, 50, 50, 50, 50, 34]);
  assert.ok(r.sets.every((s) => s.metadata.confidence === "high"), "all sets high confidence");
  assert.equal(r.stats.byType.mc, 260);
  assert.equal(r.stats.byType.sata, 24);
  assert.equal(r.stats.errorCount, 0);
  const imgs = r.questions.reduce((a, x) => a + (x.images ? x.images.length : 0), 0);
  assert.equal(imgs, 40);
  // every question has a key & options
  for (const x of r.questions) {
    assert.ok(x.correctAnswers.length >= 1, "key for " + x.id);
    assert.ok(x.options.length >= 2, "options for " + x.id);
    const labels = new Set(x.options.map((o) => o.label));
    x.correctAnswers.forEach((L) => assert.ok(labels.has(L), "key label valid in " + x.id));
  }
});

test("integration: doc 2 (Newborn cases) -> 3 open-ended sets (15/12/8), 15 images", { skip: !hasUploads }, async () => {
  const ir = await extractIRFromFile(UPLOADS + "2 Newborn Exam cases.docx");
  const r = parseExamDocuments([ir]);
  assert.equal(r.sets.length, 3);
  assert.deepEqual(r.sets.map((s) => s.questionIds.length), [15, 12, 8]);
  assert.equal(r.stats.byType.open, 35);
  assert.equal(r.stats.errorCount, 0);
  const imgs = r.questions.reduce((a, x) => a + (x.images ? x.images.length : 0), 0);
  assert.equal(imgs, 15);
  for (const x of r.questions) {
    assert.ok(x.answerBlock && x.answerBlock.length >= 2, "answer block for " + x.id);
  }
});

test("integration: combined 4-doc bank validates", { skip: !hasUploads }, async () => {
  const files = [
    "1 Newborn_Infant_Practice_Examinations.docx",
    "2 Newborn Exam cases.docx",
    "3 Toddler and School Age questions.docx",
    "4 Developmental_Behavioral_ADHD.docx",
  ];
  const irs = [];
  for (const f of files) irs.push(await extractIRFromFile(UPLOADS + f));
  const r = parseExamDocuments(irs);
  assert.equal(r.sets.length, 15);
  assert.equal(r.stats.questions, 619);
  const bank = assembleBank(r, { title: "Test bank" });
  const v = validateBank(bank);
  assert.equal(v.valid, true, JSON.stringify(v.errors.slice(0, 5)));
  // ids globally unique
  assert.equal(new Set(bank.questions.map((x) => x.id)).size, 619);
  assert.equal(new Set(bank.sets.map((x) => x.id)).size, 15);
});
