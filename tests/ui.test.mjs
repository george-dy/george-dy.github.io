/* tests/ui.test.mjs — jsdom-driven UI tests for the single-page app.
 * Boots index.html (scripts stripped, then evaluated manually) and exercises
 * library / exam / grading / persistence / reset / import flows. */
import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { JSDOM, VirtualConsole } from "jsdom";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const JSZip = require("jszip");

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const HTML = fs.readFileSync(path.join(ROOT, "index.html"), "utf8")
  .replace(/<script src="[^"]+"><\/script>/g, "");
const PARSER_SRC = fs.readFileSync(path.join(ROOT, "js", "parser-core.js"), "utf8");
const DOCXIR_SRC = fs.readFileSync(path.join(ROOT, "js", "docx-ir.js"), "utf8");
const APP_SRC = fs.readFileSync(path.join(ROOT, "js", "app.js"), "utf8");
const JSZIP_SRC = fs.readFileSync(path.join(ROOT, "vendor", "jszip.min.js"), "utf8");

function tick(w, n = 6, ms = 8) {
  return new Promise(function (res) {
    let i = 0;
    (function step() {
      if (i++ >= n) return res();
      setTimeout(step, ms);
    })();
  });
}

function makeWindow(bank, opts = {}) {
  const errors = [];
  const vc = new VirtualConsole();
  vc.on("error", (...a) => errors.push(a.map(String).join(" ")));
  vc.on("jsdomError", (e) => {
    if (/Could not parse CSS/i.test(e.message)) return; // jsdom CSSOM lacks modern CSS
    errors.push("jsdomError: " + e.message);
  });
  const dom = new JSDOM(HTML, {
    url: "http://localhost/app/",
    runScripts: "outside-only",
    pretendToBeVisual: true,
    virtualConsole: vc,
  });
  const w = dom.window;
  // jsdom lacks TextEncoder/TextDecoder (all real browsers have them);
  // JSZip needs them for .async("string"). jsdom's MutationObserver also does
  // not pump JSZip's bundled microtask scheduler, so expose setImmediate.
  if (!w.TextDecoder) w.TextDecoder = TextDecoder;
  if (!w.TextEncoder) w.TextEncoder = TextEncoder;
  if (!w.setImmediate) { w.setImmediate = setImmediate; w.clearImmediate = clearImmediate; }
  if (bank) w.EXAM_BANK = bank;
  if (opts.seedStorage) {
    for (const [k, v] of Object.entries(opts.seedStorage)) w.localStorage.setItem(k, v);
  }
  w.URL.createObjectURL = () => "blob:stub";
  w.URL.revokeObjectURL = () => {};
  w.__PEP_SKIP_MIRROR__ = true; // jsdom cannot load external scripts
  // Image stub: never loads (like jsdom), fails fast so downscale() resolves.
  w.Image = class {
    set src(v) { setTimeout(() => { if (this.onerror) this.onerror(new Error("stub")); }, 0); }
  };
  w.eval(JSZIP_SRC);
  w.eval(PARSER_SRC);
  w.eval(DOCXIR_SRC);
  w.eval(APP_SRC);
  return { w, errors, dom };
}

function testBank() {
  return JSON.parse(JSON.stringify({
    format: "exam-bank", version: 1, bankId: "test-bank", title: "Test",
    generatedAt: "2026-01-01T00:00:00Z",
    documents: [{ file: "Week 1.docx", title: "Week 1", importedAt: "2026-01-01T00:00:00Z", paragraphCount: 10, setIds: ["week-1-docx--s1-form-a"] }],
    sets: [
      {
        id: "week-1-docx--s1-form-a", title: "Form A", sourceFile: "Week 1.docx",
        description: "first set", questionIds: ["week-1-docx--s1-form-a--q001", "week-1-docx--s1-form-a--q002", "week-1-docx--s1-form-a--q003"],
        metadata: { setIndex: 1, headingText: "Form A", detectedBy: ["heading"], confidence: "high", autoTitle: false },
      },
      {
        id: "week-1-docx--s2-form-b", title: "Form B", sourceFile: "Week 1.docx",
        description: "second set", questionIds: ["week-1-docx--s2-form-b--q001"],
        metadata: { setIndex: 2, headingText: "Form B", detectedBy: ["heading"], confidence: "high", autoTitle: false },
      },
    ],
    questions: [
      {
        id: "week-1-docx--s1-form-a--q001", setId: "week-1-docx--s1-form-a", sourceFile: "Week 1.docx",
        questionNumber: 1, ordinal: 1, type: "mc",
        stem: ["A client arrives. Which finding is expected?"],
        options: [
          { label: "A", text: "Finding one", isCorrect: false, rationale: "One is wrong because alpha." },
          { label: "B", text: "Finding two", isCorrect: true, rationale: "Two is right because beta." },
          { label: "C", text: "Finding three", isCorrect: false, rationale: "" },
        ],
        correctAnswers: ["B"], sourceMapping: "Slide 3",
        metadata: { contentHash: "h1" },
      },
      {
        id: "week-1-docx--s1-form-a--q002", setId: "week-1-docx--s1-form-a", sourceFile: "Week 1.docx",
        questionNumber: 2, ordinal: 2, type: "sata",
        stem: ["Select all that apply."],
        options: [
          { label: "A", text: "alpha", isCorrect: true, rationale: "alpha yes." },
          { label: "B", text: "beta", isCorrect: false, rationale: "beta no." },
          { label: "C", text: "gamma", isCorrect: true, rationale: "gamma yes." },
          { label: "D", text: "delta", isCorrect: false, rationale: "delta no." },
        ],
        correctAnswers: ["A", "C"], metadata: { contentHash: "h2" },
      },
      {
        id: "week-1-docx--s1-form-a--q003", setId: "week-1-docx--s1-form-a", sourceFile: "Week 1.docx",
        questionNumber: 3, ordinal: 3, type: "open",
        stem: ["Describe the assessment sequence."],
        answerBlock: [{ text: "Answer source: faculty.", bold: false, bullet: null }, { text: "Inspect, palpate, auscultate.", bold: false, bullet: 0 }],
        correctAnswers: [], metadata: { contentHash: "h3" },
      },
      {
        id: "week-1-docx--s2-form-b--q001", setId: "week-1-docx--s2-form-b", sourceFile: "Week 1.docx",
        questionNumber: 1, ordinal: 1, type: "mc",
        stem: ["Second set question?"],
        options: [
          { label: "A", text: "yes", isCorrect: true, rationale: "yes." },
          { label: "B", text: "no", isCorrect: false, rationale: "no." },
        ],
        correctAnswers: ["A"], metadata: { contentHash: "h4" },
      },
    ],
  }));
}

/* ============================ library view ============================== */

test("library: renders set cards with counts, badges, and actions", async () => {
  const { w, errors } = makeWindow(testBank());
  await tick(w);
  const cards = w.document.querySelectorAll(".set-card");
  assert.equal(cards.length, 2);
  assert.match(cards[0].textContent, /Form A/);
  assert.match(cards[0].textContent, /3 questions/);
  assert.match(cards[0].textContent, /Not started/);
  assert.match(cards[0].textContent, /Set 1 of 2/);
  assert.ok(cards[0].querySelector("button.btn-primary").textContent.includes("Start"));
  assert.equal(errors.length, 0, errors.join("\n"));
});

test("library: search filters sets by title and source", async () => {
  const { w } = makeWindow(testBank());
  await tick(w);
  const search = w.document.querySelector(".search-input");
  search.value = "form b";
  search.dispatchEvent(new w.Event("input", { bubbles: true }));
  await tick(w, 8, 30);
  const cards = w.document.querySelectorAll(".set-card");
  assert.equal(cards.length, 1);
  assert.match(cards[0].textContent, /Form B/);
});

test("library: status filter chips work", async () => {
  const { w } = makeWindow(testBank());
  await tick(w);
  const chips = [...w.document.querySelectorAll(".chip")];
  const completed = chips.find((c) => c.textContent === "Completed");
  completed.click();
  await tick(w);
  assert.equal(w.document.querySelectorAll(".set-card").length, 0);
  const untouched = [...w.document.querySelectorAll(".chip")].find((c) => c.textContent === "Not started");
  untouched.click();
  await tick(w);
  assert.equal(w.document.querySelectorAll(".set-card").length, 2);
});

/* ============================= exam view ================================ */

async function openExam(w, setId = "week-1-docx--s1-form-a") {
  w.location.hash = "#/exam/" + setId;
  w.dispatchEvent(new w.Event("hashchange"));
  await tick(w);
}

test("exam: shows stem, type badge, options; rationales hidden before submit", async () => {
  const { w, errors } = makeWindow(testBank());
  await tick(w);
  await openExam(w);
  assert.match(w.document.querySelector(".q-stem").textContent, /Which finding is expected/);
  assert.match(w.document.querySelector(".type-badge").textContent, /Multiple Choice/);
  const inputs = w.document.querySelectorAll(".opts input");
  assert.equal(inputs.length, 3);
  assert.equal(inputs[0].type, "radio");
  // rationales must NOT be in the DOM before submission
  assert.equal(w.document.querySelectorAll(".rat-item").length, 0);
  assert.ok(!w.document.body.textContent.includes("One is wrong because alpha"));
  assert.equal(w.document.querySelectorAll(".verdict").length, 0);
  // prev disabled at first question
  assert.equal(w.document.querySelector(".q-prev").disabled, true);
  assert.equal(w.document.querySelector(".q-next").disabled, false);
  assert.equal(errors.length, 0, errors.join("\n"));
});

test("exam: empty submit is blocked with a warning and no grading", async () => {
  const { w } = makeWindow(testBank());
  await tick(w);
  await openExam(w);
  w.document.querySelector(".q-submit").click();
  await tick(w);
  assert.ok(w.document.querySelector(".toast-warn"), "warning toast shown");
  assert.equal(w.document.querySelectorAll(".verdict").length, 0);
  const rec = w.ExamApp.answerRec("week-1-docx--s1-form-a", "week-1-docx--s1-form-a--q001");
  assert.equal(rec, null, "nothing recorded");
});

test("exam: selection highlight, submit correct, rationale reveal, options locked", async () => {
  const { w, errors } = makeWindow(testBank());
  await tick(w);
  await openExam(w);
  const labels = w.document.querySelectorAll(".opt");
  const inputB = labels[1].querySelector("input");
  inputB.checked = true;
  inputB.dispatchEvent(new w.Event("change", { bubbles: true }));
  await tick(w, 2);
  assert.ok(labels[1].classList.contains("opt-selected"), "highlighted");

  w.document.querySelector(".q-submit").click();
  await tick(w);
  const verdict = w.document.querySelector(".verdict");
  assert.ok(verdict.classList.contains("verdict-ok"), "correct verdict");
  assert.match(verdict.textContent, /Correct/);
  const rats = w.document.querySelectorAll(".rat-item");
  assert.equal(rats.length, 3, "rationale per option revealed");
  assert.match(rats[0].textContent, /One is wrong because alpha/);
  assert.match(rats[2].textContent, /No rationale provided in the source material\./);
  // options locked after submit
  assert.equal(w.document.querySelectorAll(".opts input")[0].disabled, true);
  // navigator marks question correct
  const navBtn = w.document.querySelectorAll(".qnav-btn")[0];
  assert.ok(navBtn.classList.contains("qnav-ok"));
  // source mapping present
  assert.match(w.document.querySelector(".src-map").textContent, /Slide 3/);
  assert.equal(errors.length, 0, errors.join("\n"));
});

test("exam: SATA exact-match grading — partial selection is incorrect with partial credit note", async () => {
  const { w } = makeWindow(testBank());
  await tick(w);
  await openExam(w);
  w.document.querySelector(".q-next").click();
  await tick(w);
  assert.equal(w.document.querySelectorAll(".opts input")[0].type, "checkbox");
  const inputs = w.document.querySelectorAll(".opts input");
  inputs[0].checked = true;
  inputs[0].dispatchEvent(new w.Event("change", { bubbles: true }));
  await tick(w, 2);
  w.document.querySelector(".q-submit").click();
  await tick(w);
  const verdict = w.document.querySelector(".verdict");
  assert.ok(verdict.classList.contains("verdict-no"), "partial is NOT correct");
  assert.match(verdict.textContent, /partial credit 50%/);
  const rec = w.ExamApp.answerRec("week-1-docx--s1-form-a", "week-1-docx--s1-form-a--q002");
  assert.equal(rec.correct, false);
  assert.equal(rec.partial, 0.5);
});

test("exam: navigation preserves selection and submission state", async () => {
  const { w } = makeWindow(testBank());
  await tick(w);
  await openExam(w);
  const inputs = w.document.querySelectorAll(".opts input");
  inputs[2].checked = true;
  inputs[2].dispatchEvent(new w.Event("change", { bubbles: true }));
  await tick(w, 2);
  w.document.querySelector(".q-next").click();   // to Q2 (unanswered)
  await tick(w);
  assert.match(w.document.querySelector(".q-number").textContent, /Question 2 of 3/);
  // jump directly to Q3 via navigator
  w.document.querySelectorAll(".qnav-btn")[2].click();
  await tick(w);
  assert.match(w.document.querySelector(".q-number").textContent, /Question 3 of 3/);
  assert.equal(w.document.querySelector(".q-next").disabled, true, "next disabled on last");
  // back to Q1: selection preserved, still unsubmitted
  w.document.querySelectorAll(".qnav-btn")[0].click();
  await tick(w);
  const q1inputs = w.document.querySelectorAll(".opts input");
  assert.equal(q1inputs[2].checked, true, "selection preserved");
  assert.equal(w.document.querySelectorAll(".verdict").length, 0, "not submitted");
});

test("exam: submitted question shows recorded result on return without resubmission", async () => {
  const { w } = makeWindow(testBank());
  await tick(w);
  await openExam(w);
  const inputs = w.document.querySelectorAll(".opts input");
  inputs[1].checked = true;
  inputs[1].dispatchEvent(new w.Event("change", { bubbles: true }));
  await tick(w, 2);
  w.document.querySelector(".q-submit").click();
  await tick(w);
  w.document.querySelector(".q-next").click();
  await tick(w);
  w.document.querySelectorAll(".qnav-btn")[0].click();
  await tick(w);
  assert.ok(w.document.querySelector(".verdict-ok"), "result shown again");
  assert.equal(w.document.querySelectorAll(".q-submit").length, 0, "no double submit button");
  assert.ok(w.document.querySelector(".rat-item"), "rationales still revealed");
});

test("exam: open-ended question reveal + self-grade affects stats", async () => {
  const { w } = makeWindow(testBank());
  await tick(w);
  await openExam(w);
  w.document.querySelectorAll(".qnav-btn")[2].click();
  await tick(w);
  assert.match(w.document.querySelector(".type-badge").textContent, /Open-Ended/);
  assert.ok(!w.document.querySelector(".open-answer"), "answer hidden pre-reveal");
  // notes textarea persists
  const ta = w.document.querySelector(".open-note");
  ta.value = "my study note";
  ta.dispatchEvent(new w.Event("input", { bubbles: true }));
  await tick(w, 10, 20);
  w.document.querySelector(".q-reveal").click();
  await tick(w);
  const block = w.document.querySelector(".open-answer");
  assert.ok(block, "faculty answer revealed");
  assert.match(block.textContent, /Inspect, palpate, auscultate/);
  // self grade
  const knew = [...w.document.querySelectorAll(".self-grade button")].find((b) => /knew/.test(b.textContent));
  knew.click();
  await tick(w);
  const stats = w.ExamApp.setStats("week-1-docx--s1-form-a");
  assert.equal(stats.correct, 1);
  assert.equal(stats.answered, 1);
});

/* =========================== persistence ================================ */

test("persistence: progress survives a simulated reload and resumes position", async () => {
  const { w } = makeWindow(testBank());
  await tick(w);
  await openExam(w);
  const inputs = w.document.querySelectorAll(".opts input");
  inputs[1].checked = true;
  inputs[1].dispatchEvent(new w.Event("change", { bubbles: true }));
  await tick(w, 2);
  w.document.querySelector(".q-submit").click();
  await tick(w);
  w.document.querySelector(".q-next").click();  // last = index 1
  await tick(w);

  // capture storage, rebuild window
  const seed = {};
  for (let i = 0; i < w.localStorage.length; i++) {
    const k = w.localStorage.key(i);
    seed[k] = w.localStorage.getItem(k);
  }
  const second = makeWindow(testBank(), { seedStorage: seed });
  await tick(second.w);
  // library shows in-progress stats
  const card = second.w.document.querySelector(".set-card");
  assert.match(card.textContent, /In progress/);
  assert.match(card.textContent, /1\/3 answered · 1 correct/);
  assert.ok(card.querySelector("button.btn-primary").textContent.includes("Continue"));
  // resume at saved index
  await openExam(second.w);
  assert.match(second.w.document.querySelector(".q-number").textContent, /Question 2 of 3/);
  // previously submitted Q1 still shows feedback when revisited
  second.w.document.querySelectorAll(".qnav-btn")[0].click();
  await tick(second.w);
  assert.ok(second.w.document.querySelector(".verdict-ok"));
});

test("reset: resetting one set leaves other sets untouched", async () => {
  const { w } = makeWindow(testBank());
  await tick(w);
  // answer Q in set A and set B
  await openExam(w, "week-1-docx--s1-form-a");
  let inputs = w.document.querySelectorAll(".opts input");
  inputs[1].checked = true;
  inputs[1].dispatchEvent(new w.Event("change", { bubbles: true }));
  await tick(w, 2);
  w.document.querySelector(".q-submit").click();
  await tick(w);
  await openExam(w, "week-1-docx--s2-form-b");
  inputs = w.document.querySelectorAll(".opts input");
  inputs[0].checked = true;
  inputs[0].dispatchEvent(new w.Event("change", { bubbles: true }));
  await tick(w, 2);
  w.document.querySelector(".q-submit").click();
  await tick(w);
  assert.equal(w.ExamApp.setStats("week-1-docx--s2-form-b").answered, 1);

  // back to library, reset set A via confirm dialog
  w.location.hash = "#/";
  w.dispatchEvent(new w.Event("hashchange"));
  await tick(w);
  const resetBtns = [...w.document.querySelectorAll(".set-card .btn-reset")];
  resetBtns[0].click();
  await tick(w);
  const dialog = w.document.querySelector(".modal");
  assert.ok(dialog, "confirmation dialog appears");
  assert.match(dialog.textContent, /permanently cleared/);
  dialog.querySelectorAll(".modal-actions button")[1].click(); // confirm
  await tick(w, 6, 15);
  assert.equal(w.ExamApp.setStats("week-1-docx--s1-form-a").done, 0, "set A cleared");
  assert.equal(w.ExamApp.setStats("week-1-docx--s2-form-b").answered, 1, "set B untouched");
});

test("storage: quota failure degrades gracefully (toast, no crash)", async () => {
  const { w } = makeWindow(testBank());
  await tick(w);
  const orig = w.localStorage.setItem.bind(w.localStorage);
  w.localStorage.setItem = function (k, v) {
    const e = new Error("quota"); e.name = "QuotaExceededError"; throw e;
  };
  await openExam(w);
  const inputs = w.document.querySelectorAll(".opts input");
  inputs[1].checked = true;
  inputs[1].dispatchEvent(new w.Event("change", { bubbles: true }));
  await tick(w, 3);
  assert.ok(w.document.querySelector(".toast-warn"), "user warned about storage");
  w.localStorage.setItem = orig;
});

test("bank: missing bank shows friendly empty state with demo option", async () => {
  const { w } = makeWindow(null);
  await tick(w);
  assert.ok(w.document.querySelector(".empty-state"));
  assert.match(w.document.body.textContent, /No question bank found/);
  const demoBtn = [...w.document.querySelectorAll("button")].find((b) => /demo bank/i.test(b.textContent));
  demoBtn.click();
  await tick(w);
  assert.equal(w.document.querySelectorAll(".set-card").length, 1);
});

test("grading unit checks via ExamApp.grade", async () => {
  const { w } = makeWindow(testBank());
  await tick(w);
  const q = w.ExamApp.state.bank.questions[1]; // SATA A,C
  assert.equal(w.ExamApp.grade(q, ["A", "C"]).correct, true);
  assert.equal(w.ExamApp.grade(q, ["A"]).correct, false);
  assert.equal(w.ExamApp.grade(q, ["A", "B", "C"]).correct, false);
  assert.equal(w.ExamApp.grade(q, ["C", "A"]).correct, true, "order-insensitive");
  const mc = w.ExamApp.state.bank.questions[0];
  assert.equal(w.ExamApp.grade(mc, ["B"]).correct, true);
  assert.equal(w.ExamApp.grade(mc, ["A"]).correct, false);
  assert.equal(w.ExamApp.grade(mc, []).correct, false);
});

/* ===================== in-browser .docx import flow ===================== */

async function makeTinyDocx(w) {
  const zip = new JSZip();
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const p = (text, bold) => `<w:p>${bold ? '<w:r><w:rPr><w:b/></w:rPr>' : '<w:r>'}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const paras = [
    p("Week 9 Practice Examination", true),
    p("Form A - 2 Items", true),
    p("Question 1: Multiple Choice", true),
    p("Imported stem one?"),
    p("A) one"), p("B) two"),
    p("Correct Answer: B"),
    p("Clinical Rationale:"),
    p("B) two: because import."),
    p("A) one: nope."),
    p("Source Mapping: Imported slide 1"),
    p("Question 2: Select All That Apply", true),
    p("Imported SATA stem?"),
    p("A) a"), p("B) b"), p("C) c"),
    p("Correct Answer: A, C"),
    p("Clinical Rationale:"),
    p("A) a: yes."), p("C) c: yes."), p("B) b: no."),
    p("Source Mapping: Imported slide 2"),
  ];
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}"><w:body>${paras.join("")}<w:sectPr/></w:body></w:document>`;
  zip.file("word/document.xml", xml);
  const buf = await zip.generateAsync({ type: "arraybuffer" });
  // Re-create the bytes inside the jsdom realm (cross-realm ArrayBuffers
  // break JSZip's instanceof checks; real browsers never hit this).
  const nodeBytes = new Uint8Array(buf);
  const wBytes = new w.Uint8Array(nodeBytes.length);
  wBytes.set(nodeBytes);
  return wBytes.buffer;
}

test("import: in-browser docx parse -> review -> accept -> local bank", async () => {
  const { w, errors } = makeWindow(testBank());
  await tick(w);
  w.location.hash = "#/import";
  w.dispatchEvent(new w.Event("hashchange"));
  await tick(w);
  assert.ok(w.document.querySelector(".dropzone"), "dropzone rendered");

  const buf = await makeTinyDocx(w);
  const fake = { name: "Week 9.docx", arrayBuffer: async () => buf };
  await w.ExamApp.importFlow._api.startParse([fake]);
  await tick(w, 10, 20);

  const flow = w.ExamApp.importFlow;
  assert.equal(flow.files[0].status, "done", JSON.stringify(flow.files[0].error || ""));
  assert.ok(flow.draft, "draft built");
  assert.equal(flow.draft.sets.length, 1);
  assert.equal(flow.draft.sets[0].title, "Form A - 2 Items");
  assert.equal(flow.draft.questions.length, 2);
  assert.equal(flow.draft.questions[1].type, "sata");
  assert.deepEqual([...flow.draft.questions[1].correctAnswers], ["A", "C"]);
  // review UI present
  assert.ok(w.document.querySelector(".set-panel"), "review panel rendered");
  assert.match(w.document.body.textContent, /2 question\(s\) detected/);

  flow._api.applyImport();
  await tick(w);
  assert.equal(w.ExamApp.state.bankSource, "local");
  const sets = w.ExamApp.state.bank.sets;
  assert.ok(sets.some((s) => s.sourceFile === "Week 9.docx"), "new set merged into bank");
  assert.equal(sets.length, 3, "existing sets preserved");
  // publish step visible
  assert.match(w.document.body.textContent, /Publish your updated bank/);
  const hardErrors = errors.filter((e) => !/parse/i.test(e));
  assert.equal(hardErrors.length, 0, hardErrors.join("\n"));
});

test("import: duplicate detection offers skip/replace/copy and copy renumbers ids", async () => {
  const { w } = makeWindow(testBank());
  await tick(w);
  w.location.hash = "#/import";
  w.dispatchEvent(new w.Event("hashchange"));
  await tick(w);
  // Re-import the SAME bank content as a docx? Simulate by importing a docx
  // matching an existing set fingerprint is complex; instead verify dup logic
  // via fingerprint equality of identical parsed content:
  const buf = await makeTinyDocx(w);
  const fake = { name: "Week 9.docx", arrayBuffer: async () => buf };
  await w.ExamApp.importFlow._api.startParse([fake]);
  await tick(w, 10, 20);
  w.ExamApp.importFlow._api.applyImport();
  await tick(w);
  // second import of same file -> duplicate detected
  w.location.hash = "#/";
  w.dispatchEvent(new w.Event("hashchange"));
  await tick(w);
  w.location.hash = "#/import";
  w.dispatchEvent(new w.Event("hashchange"));
  await tick(w);
  const fake2 = { name: "Week 9.docx", arrayBuffer: async () => buf };
  w.ExamApp.importFlow.files = [];
  w.ExamApp.importFlow.draft = null;
  w.ExamApp.importFlow.step = 1;
  await w.ExamApp.importFlow._api.startParse([fake2]);
  await tick(w, 10, 20);
  const d = w.ExamApp.importFlow.draft;
  assert.equal(d.dupes.length, 1, "duplicate set detected");
  assert.ok(w.document.querySelector(".dup-box"), "dup UI rendered");
  // choose copy
  d.dupes[0].mode = "copy";
  w.ExamApp.importFlow._api.applyImport();
  await tick(w);
  const sets = w.ExamApp.state.bank.sets;
  const week9 = sets.filter((s) => s.sourceFile === "Week 9.docx");
  assert.equal(week9.length, 2, "kept both copies");
  assert.notEqual(week9[0].id, week9[1].id);
  const v = w.ExamParser.validateBank(w.ExamApp.state.bank);
  assert.equal(v.valid, true, JSON.stringify(v.errors.slice(0, 3)));
});

test("settings: reset-all requires typing RESET", async () => {
  const { w } = makeWindow(testBank());
  await tick(w);
  w.location.hash = "#/settings";
  w.dispatchEvent(new w.Event("hashchange"));
  await tick(w);
  const resetAll = [...w.document.querySelectorAll("button")].find((b) => /Reset ALL progress/.test(b.textContent));
  resetAll.click();
  await tick(w);
  const dialog = w.document.querySelector(".modal");
  assert.ok(dialog);
  const okBtn = dialog.querySelectorAll(".modal-actions button")[1];
  assert.equal(okBtn.disabled, true, "confirm disabled until text typed");
  dialog.querySelector("input").value = "RESET";
  dialog.querySelector("input").dispatchEvent(new w.Event("input", { bubbles: true }));
  await tick(w, 2);
  assert.equal(okBtn.disabled, false);
  okBtn.click();
  await tick(w, 4, 12);
  assert.deepEqual(JSON.parse(JSON.stringify(w.ExamApp.state.progress)), {});
});
