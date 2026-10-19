/* tests/e2e.test.mjs — End-to-end: serve the project over HTTP and drive the
 * real published site (real scripts, real question-bank.json via fetch) in
 * jsdom. Closest available substitute for a browser session. */
import test from "node:test";
import assert from "node:assert/strict";
import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { JSDOM, VirtualConsole, ResourceLoader } from "jsdom";

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const MIME = {
  ".html": "text/html", ".js": "text/javascript", ".json": "application/json",
  ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".md": "text/markdown",
};

function serve() {
  const server = http.createServer((req, res) => {
    let p = decodeURIComponent(req.url.split("?")[0]);
    if (p === "/") p = "/index.html";
    const file = path.join(ROOT, p);
    if (!file.startsWith(ROOT) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) {
      res.writeHead(404); res.end("not found"); return;
    }
    res.writeHead(200, { "Content-Type": MIME[path.extname(file)] || "application/octet-stream" });
    res.end(fs.readFileSync(file));
  });
  return new Promise((resolve) => {
    server.listen(0, "127.0.0.1", () => resolve(server));
  });
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
async function waitFor(fn, ms = 8000) {
  const t0 = Date.now();
  while (Date.now() - t0 < ms) {
    const v = fn();
    if (v) return v;
    await sleep(60);
  }
  throw new Error("waitFor timed out");
}

test("e2e: published site loads bank via fetch, exam flow works", async () => {
  const server = await serve();
  const port = server.address().port;
  const url = `http://127.0.0.1:${port}/`;

  const errors = [];
  const vc = new VirtualConsole();
  vc.on("error", (...a) => errors.push(a.map(String).join(" ")));
  vc.on("jsdomError", (e) => { if (!/Could not parse CSS/i.test(e.message)) errors.push("jsdomError: " + e.message); });

  const dom = await JSDOM.fromURL(url, {
    runScripts: "dangerously",
    resources: "usable",
    pretendToBeVisual: true,
    virtualConsole: vc,
    beforeParse(window) {
      // jsdom has no fetch; route to Node's fetch (same semantics the app uses)
      window.fetch = (u, o) => fetch(new URL(String(u), window.location.href), o);
      window.TextDecoder = window.TextDecoder || TextDecoder;
      window.TextEncoder = window.TextEncoder || TextEncoder;
      window.URL.createObjectURL = () => "blob:stub";
      window.URL.revokeObjectURL = () => {};
    },
  });
  const w = dom.window;
  const d = w.document;

  try {
    // 1. library renders the REAL bank
    await waitFor(() => d.querySelectorAll(".set-card").length >= 15, 15000);
    const cards = d.querySelectorAll(".set-card");
    assert.equal(cards.length, 15, "all 15 sets listed");
    assert.match(d.querySelector(".view-sub").textContent, /619 questions/);
    assert.equal(w.ExamApp.state.bankSource, "published", "loaded via fetch, not mirror");

    // 2. search works on real data
    const search = d.querySelector(".search-input");
    search.value = "ADHD";
    search.dispatchEvent(new w.Event("input", { bubbles: true }));
    await sleep(300);
    assert.equal(d.querySelectorAll(".set-card").length, 3, "ADHD filter -> 3 sets");
    search.value = "";
    search.dispatchEvent(new w.Event("input", { bubbles: true }));
    await sleep(300);

    // 3. open first set, answer Q1 correctly using the bank's own key
    const setId = w.ExamApp.state.bank.sets[0].id;
    const q1 = w.ExamApp.setQuestions(setId)[0];
    const key = q1.correctAnswers[0];
    cards[0].querySelector("button.btn-primary").click();
    await waitFor(() => d.querySelector(".q-card"));
    assert.ok(d.querySelector(".q-stem").textContent.includes(q1.stem[0].slice(0, 40)));
    assert.equal(d.querySelectorAll(".rat-item").length, 0, "no rationales pre-submit");
    const optLabels = [...d.querySelectorAll(".opt")];
    const target = optLabels.find((l) => l.querySelector(".opt-key").textContent === key);
    const input = target.querySelector("input");
    input.checked = true;
    input.dispatchEvent(new w.Event("change", { bubbles: true }));
    await sleep(80);
    d.querySelector(".q-submit").click();
    await waitFor(() => d.querySelector(".verdict"));
    assert.ok(d.querySelector(".verdict-ok"), "real bank Q1 graded correct");
    assert.equal(d.querySelectorAll(".rat-item").length, q1.options.length);

    // 4. an image question renders its <img> (doc1 set 6)
    const imgSet = w.ExamApp.state.bank.sets.find((s) => s.id.includes("--s6-"));
    w.location.hash = "#/exam/" + encodeURIComponent(imgSet.id);
    w.dispatchEvent(new w.Event("hashchange"));
    await waitFor(() => d.querySelector(".q-card"));
    await waitFor(() => d.querySelector(".q-img-el"), 5000);
    const img = d.querySelector(".q-img-el");
    assert.match(img.getAttribute("src"), /^assets\//, "relative asset path");
    assert.ok(img.alt.length > 5, "alt text present");

    // 5. progress persisted to localStorage
    const prog = JSON.parse(w.localStorage.getItem("pep.progress.v1"));
    assert.ok(prog[setId].answers[q1.id].submitted);

    // 6. console clean (ignore CSS-parse noise and image load failures in jsdom)
    const real = errors.filter((e) => !/Could not parse CSS/i.test(e));
    assert.equal(real.length, 0, real.join("\n"));
  } finally {
    w.close();
    server.close();
  }
});
