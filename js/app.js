/* ============================================================================
 * app.js — Practice Examination Platform (single-page application logic)
 * Classic script (no modules) so it works from GitHub Pages AND file://
 * double-click. Depends on: js/parser-core.js (window.ExamParser) and, for
 * the import pipeline, js/docx-ir.js + vendor/jszip.min.js.
 *
 * Views: Library (#/), Exam (#/exam/<setId>), Import (#/import),
 *        Settings (#/settings).
 * ==========================================================================*/
(function () {
  "use strict";

  /* ============================ tiny helpers ============================= */

  function el(tag, className, text) {
    const n = document.createElement(tag);
    if (className) n.className = className;
    if (text != null) n.textContent = text;
    return n;
  }
  function svgIcon(name) {
    const span = el("span", "icon icon-" + name);
    span.setAttribute("aria-hidden", "true");
    span.innerHTML = ICONS[name] || "";
    return span;
  }
  const ICONS = {
    book: '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/></svg>',
    back: '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 12H5"/><path d="M12 19l-7-7 7-7"/></svg>',
    check: '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6L9 17l-5-5"/></svg>',
    x: '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M18 6L6 18"/><path d="M6 6l12 12"/></svg>',
    half: '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M12 3a9 9 0 0 1 0 18z" fill="currentColor" stroke="none"/></svg>',
    gear: '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 1 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 1 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 1 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 1 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>',
    upload: '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="M17 8l-5-5-5 5"/><path d="M12 3v12"/></svg>',
    download: '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="M7 10l5 5 5-5"/><path d="M12 15V3"/></svg>',
    reset: '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12a9 9 0 1 0 3-6.7L3 8"/><path d="M3 3v5h5"/></svg>',
    zoom: '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.35-4.35"/><path d="M11 8v6M8 11h6"/></svg>',
    grid: '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>',
    warn: '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><path d="M12 9v4"/><path d="M12 17h.01"/></svg>',
  };

  function fmtPct(n) { return Math.round(n * 100) + "%"; }
  function nowISO() { return new Date().toISOString(); }
  function debounce(fn, ms) {
    let t = null;
    return function () {
      const args = arguments;
      clearTimeout(t);
      t = setTimeout(function () { fn.apply(null, args); }, ms);
    };
  }

  /* ============================== storage ================================ */

  const LS_KEYS = { progress: "pep.progress.v1", bank: "pep.bank.v1", prefs: "pep.prefs.v1", bankMeta: "pep.bankmeta.v1" };

  const Storage = {
    available: (function () {
      try {
        const k = "__pep_test__";
        window.localStorage.setItem(k, "1");
        window.localStorage.removeItem(k);
        return true;
      } catch (e) { return false; }
    })(),
    get(key, fallback) {
      if (!this.available) return fallback;
      try {
        const raw = window.localStorage.getItem(key);
        if (raw == null) return fallback;
        return JSON.parse(raw);
      } catch (e) {
        console.warn("Could not read localStorage key", key, e);
        return fallback;
      }
    },
    set(key, value) {
      if (!this.available) return { ok: false, reason: "unavailable" };
      try {
        window.localStorage.setItem(key, JSON.stringify(value));
        return { ok: true };
      } catch (e) {
        console.warn("Could not write localStorage key", key, e);
        return { ok: false, reason: (e && e.name === "QuotaExceededError") ? "quota" : "error" };
      }
    },
    remove(key) {
      if (!this.available) return;
      try { window.localStorage.removeItem(key); } catch (e) { /* ignore */ }
    },
  };

  /* =============================== state ================================= */

  const S = {
    bank: null,            // active question bank object
    bankSource: "none",    // 'published' | 'local' | 'demo' | 'none'
    bankNote: "",
    progress: {},          // { [setId]: { answers: {qid: rec}, last: n, updated } }
    prefs: { theme: "system", navOpen: true },
    route: { name: "library", setId: null },
    exam: null,            // { setId, index }
    libSearch: "",
    libFilter: "all",
  };

  function loadPersisted() {
    S.progress = Storage.get(LS_KEYS.progress, {}) || {};
    const prefs = Storage.get(LS_KEYS.prefs, null);
    if (prefs && typeof prefs === "object") S.prefs = Object.assign(S.prefs, prefs);
  }

  function saveProgress() {
    const res = Storage.set(LS_KEYS.progress, S.progress);
    if (!res.ok) {
      toast(res.reason === "quota"
        ? "Warning: browser storage is full — progress from this session may not be saved. Export a progress backup from Settings."
        : "Warning: browser storage is unavailable; progress will not persist after you close this tab.", "warn", 8000);
    }
    return res;
  }
  function savePrefs() { Storage.set(LS_KEYS.prefs, S.prefs); }

  /* ====================== progress / grading model ======================= */

  function setProgress(setId) {
    if (!S.progress[setId]) S.progress[setId] = { answers: {}, last: 0, updated: 0 };
    return S.progress[setId];
  }
  function answerRec(setId, qid) {
    const p = setProgress(setId);
    return p.answers[qid] || null;
  }
  function bankSet(setId) {
    if (!S.bank) return null;
    return (S.bank.sets || []).find((s) => s.id === setId) || null;
  }
  function setQuestions(setId) {
    if (!S.bank) return [];
    const set = bankSet(setId);
    if (!set) return [];
    const byId = new Map((S.bank.questions || []).map((q) => [q.id, q]));
    return (set.questionIds || []).map((id) => byId.get(id)).filter(Boolean);
  }
  function findQuestion(qid) {
    if (!S.bank) return null;
    return (S.bank.questions || []).find((q) => q.id === qid) || null;
  }

  function gradeQuestion(q, selectedLabels) {
    if (q.type === "open") return { correct: null, partial: 0 };
    const key = (q.correctAnswers || []).slice().sort();
    const sel = (selectedLabels || []).slice().sort();
    if (!key.length) return { correct: null, partial: 0 };
    if (q.type === "sata") {
      const hits = sel.filter((L) => key.indexOf(L) !== -1).length;
      const wrongs = sel.length - hits;
      const partial = Math.max(0, (hits - wrongs) / key.length);
      const exact = sel.length === key.length && hits === key.length;
      return { correct: exact, partial: exact ? 1 : partial };
    }
    const exact = sel.length === 1 && sel[0] === key[0];
    return { correct: exact, partial: exact ? 1 : 0 };
  }

  function setStats(setId) {
    const qs = setQuestions(setId);
    const p = S.progress[setId];
    let answered = 0, correct = 0, incorrect = 0, revealed = 0;
    qs.forEach(function (q) {
      const rec = p && p.answers[q.id];
      if (!rec) return;
      if (rec.submitted) {
        answered++;
        if (rec.correct === true) correct++;
        else if (rec.correct === false) incorrect++;
      } else if (rec.revealed) {
        revealed++;
      }
    });
    const total = qs.length;
    const done = answered + revealed;
    return {
      total, answered, correct, incorrect, revealed, done,
      pct: total ? done / total : 0,
      started: done > 0,
      completed: total > 0 && done === total,
    };
  }

  /* ============================ routing ================================== */

  function navigate(hash) {
    if (window.location.hash === hash) render();
    else window.location.hash = hash;
  }
  function parseHash() {
    const h = (window.location.hash || "#/").replace(/^#/, "");
    const parts = h.split("/").filter(Boolean);
    if (parts[0] === "exam" && parts[1]) return { name: "exam", setId: decodeURIComponent(parts[1]) };
    if (parts[0] === "import") return { name: "import" };
    if (parts[0] === "settings") return { name: "settings" };
    return { name: "library" };
  }

  /* ============================== toasts ================================= */

  let toastHost = null;
  function toast(msg, kind, ms) {
    if (!toastHost) {
      toastHost = el("div", "toast-host");
      toastHost.setAttribute("aria-live", "polite");
      document.body.appendChild(toastHost);
    }
    const t = el("div", "toast toast-" + (kind || "info"));
    t.setAttribute("role", kind === "error" ? "alert" : "status");
    t.textContent = msg;
    toastHost.appendChild(t);
    setTimeout(function () {
      t.classList.add("toast-out");
      setTimeout(function () { t.remove(); }, 300);
    }, ms || 4200);
  }

  /* ============================== modals ================================= */

  function confirmDialog(opts) {
    return new Promise(function (resolve) {
      const overlay = el("div", "modal-overlay");
      const dialog = el("div", "modal");
      dialog.setAttribute("role", "dialog");
      dialog.setAttribute("aria-modal", "true");
      dialog.setAttribute("aria-labelledby", "modal-title");
      const h2 = el("h2", "modal-title", opts.title);
      h2.id = "modal-title";
      dialog.appendChild(h2);
      const body = el("div", "modal-body");
      if (typeof opts.body === "string") body.appendChild(el("p", null, opts.body));
      else if (opts.body) body.appendChild(opts.body);
      dialog.appendChild(body);

      let input = null;
      if (opts.requireText) {
        const lab = el("label", "modal-require", "Type \"" + opts.requireText + "\" to confirm:");
        input = el("input", "text-input");
        input.type = "text";
        input.autocomplete = "off";
        lab.appendChild(input);
        dialog.appendChild(lab);
      }

      const actions = el("div", "modal-actions");
      const cancel = el("button", "btn " + (opts.danger ? "btn-ghost" : "btn-ghost"), opts.cancelLabel || "Cancel");
      const ok = el("button", "btn " + (opts.danger ? "btn-danger" : "btn-primary"), opts.confirmLabel || "Confirm");
      if (opts.requireText) ok.disabled = true;
      actions.appendChild(cancel);
      actions.appendChild(ok);
      dialog.appendChild(actions);
      overlay.appendChild(dialog);
      document.body.appendChild(overlay);

      function close(val) {
        overlay.remove();
        document.removeEventListener("keydown", onKey, true);
        resolve(val);
      }
      function onKey(e) {
        if (e.key === "Escape") { e.stopPropagation(); close(false); }
        if (e.key === "Tab") {
          const focusables = dialog.querySelectorAll("button, input, [tabindex]");
          if (!focusables.length) return;
          const first = focusables[0], last = focusables[focusables.length - 1];
          if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
          else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
        }
      }
      cancel.addEventListener("click", function () { close(false); });
      ok.addEventListener("click", function () { close(true); });
      if (input) input.addEventListener("input", function () {
        ok.disabled = input.value.trim() !== opts.requireText;
      });
      overlay.addEventListener("mousedown", function (e) { if (e.target === overlay) close(false); });
      document.addEventListener("keydown", onKey, true);
      (input || ok).focus();
    });
  }

  /* ============================ lightbox ================================= */

  function openLightbox(src, alt) {
    const overlay = el("div", "lightbox-overlay");
    const fig = el("figure", "lightbox-fig");
    const img = el("img", "lightbox-img");
    img.src = src;
    img.alt = alt || "Enlarged examination image";
    fig.appendChild(img);
    const cap = el("figcaption", "lightbox-cap", alt || "");
    fig.appendChild(cap);
    const closeBtn = el("button", "btn btn-ghost lightbox-close", "Close (Esc)");
    overlay.appendChild(fig);
    overlay.appendChild(closeBtn);
    document.body.appendChild(overlay);
    function close() {
      overlay.remove();
      document.removeEventListener("keydown", onKey, true);
    }
    function onKey(e) { if (e.key === "Escape") close(); }
    closeBtn.addEventListener("click", close);
    overlay.addEventListener("mousedown", function (e) { if (e.target === overlay) close(); });
    document.addEventListener("keydown", onKey, true);
    closeBtn.focus();
  }

  function safeImgSrc(src) {
    if (!src || typeof src !== "string") return null;
    if (/^data:image\/(png|jpeg|jpg|gif|webp);base64,/i.test(src)) return src;
    if (/^https?:\/\//i.test(src)) return src;
    if (/^[a-z0-9._\/%-]+$/i.test(src) && src.indexOf("..") === -1) return src; // relative path
    return null;
  }

  /* ============================ rendering ================================ */

  let mainHost = null;

  function render() {
    S.route = parseHash();
    mainHost.textContent = "";
    document.querySelectorAll(".topnav a").forEach(function (a) {
      const href = (a.getAttribute("href") || "").replace(/^#/, "");
      const active = href === (S.route.name === "library" ? "/" : "/" + S.route.name) ||
        (S.route.name === "exam" && href === "/");
      a.classList.toggle("active", href === (S.route.name === "library" ? "/" : "/" + S.route.name));
      if (active) a.setAttribute("aria-current", "page"); else a.removeAttribute("aria-current");
    });
    if (S.route.name === "exam") viewExam(S.route.setId);
    else if (S.route.name === "import") viewImport();
    else if (S.route.name === "settings") viewSettings();
    else viewLibrary();
  }

  /* ------------------------------ LIBRARY ------------------------------- */

  function docLabel(fileName) {
    return String(fileName || "unknown").replace(/\.docx$/i, "");
  }

  function viewLibrary() {
    const wrap = el("div", "view library-view");
    const head = el("header", "view-head");
    head.appendChild(el("h1", "view-title", "Examination Library"));
    const sub = el("p", "view-sub");
    if (S.bank) {
      sub.textContent = (S.bank.sets || []).length + " examination sets · " +
        (S.bank.questions || []).length + " questions · from " +
        ((S.bank.documents || []).length) + " source document(s)";
    } else {
      sub.textContent = "No question bank loaded yet.";
    }
    head.appendChild(sub);
    wrap.appendChild(head);

    if (S.bankSource === "local") {
      const banner = el("div", "banner banner-info");
      banner.appendChild(svgIcon("warn"));
      const txt = el("div", null);
      txt.appendChild(el("p", null, "You are using a question bank stored in this browser (imported locally). Classmates will not see it until you publish it to GitHub."));
      const actions = el("div", "banner-actions");
      const pub = el("a", "btn btn-small btn-ghost", "How to publish");
      pub.href = "#/import";
      const usePub = el("button", "btn btn-small btn-ghost", "Use published bank instead");
      usePub.addEventListener("click", function () {
        Storage.remove(LS_KEYS.bank);
        Storage.remove(LS_KEYS.bankMeta);
        loadBank(true).then(function () { render(); toast("Switched to the published question bank."); });
      });
      actions.appendChild(pub); actions.appendChild(usePub);
      txt.appendChild(actions);
      banner.appendChild(txt);
      wrap.appendChild(banner);
    }

    if (!S.bank) {
      wrap.appendChild(emptyState());
      mainHost.appendChild(wrap);
      return;
    }

    const controls = el("div", "lib-controls");
    const searchWrap = el("div", "search-wrap");
    const search = el("input", "text-input search-input");
    search.type = "search";
    search.placeholder = "Search by week, document, or set title…";
    search.setAttribute("aria-label", "Search examination sets");
    search.value = S.libSearch || "";
    searchWrap.appendChild(svgIcon("zoom"));
    searchWrap.appendChild(search);
    controls.appendChild(searchWrap);

    const filters = el("div", "filter-chips");
    const FILTERS = [
      ["all", "All"], ["untouched", "Not started"], ["progress", "In progress"], ["completed", "Completed"],
    ];
    FILTERS.forEach(function (f) {
      const b = el("button", "chip" + ((S.libFilter || "all") === f[0] ? " chip-active" : ""), f[1]);
      b.type = "button";
      b.addEventListener("click", function () {
        S.libFilter = f[0];
        render();
      });
      filters.appendChild(b);
    });
    controls.appendChild(filters);
    wrap.appendChild(controls);

    search.addEventListener("input", debounce(function () {
      S.libSearch = search.value;
      rerenderLibraryList(listHost);
    }, 160));

    const listHost = el("div", "lib-list");
    rerenderLibraryList(listHost);
    wrap.appendChild(listHost);

    const foot = el("p", "lib-note");
    foot.textContent = "Progress is saved only in this browser (localStorage). It does not sync across devices or browsers. Use Settings → Export progress to back it up.";
    wrap.appendChild(foot);

    mainHost.appendChild(wrap);
  }

  function emptyState() {
    const card = el("div", "empty-state");
    card.appendChild(svgIcon("book"));
    card.appendChild(el("h2", null, "No question bank found"));
    card.appendChild(el("p", null,
      "This app loads its questions from a file called question-bank.json next to index.html. That file is missing (or could not be read — this can happen when you open index.html by double-clicking it)."));
    const ol = el("ol", "empty-steps");
    [
      "To practice right away: open the Import screen and add your .docx examination documents.",
      "To publish questions for your whole class: follow the README instructions to build question-bank.json and upload it with index.html to GitHub Pages.",
      "If you opened this page by double-clicking index.html: also keep question-bank.js in the same folder, or serve the folder with a local web server instead.",
    ].forEach(function (t) { ol.appendChild(el("li", null, t)); });
    card.appendChild(ol);
    const actions = el("div", "empty-actions");
    const imp = el("a", "btn btn-primary", "Import Word documents");
    imp.href = "#/import";
    const demo = el("button", "btn btn-ghost", "Load 4-question demo bank");
    demo.addEventListener("click", function () {
      S.bank = demoBank();
      S.bankSource = "demo";
      toast("Demo bank loaded (not saved as your published bank).");
      render();
    });
    actions.appendChild(imp); actions.appendChild(demo);
    card.appendChild(actions);
    return card;
  }

  function rerenderLibraryList(host) {
    host.textContent = "";
    const term = (S.libSearch || "").trim().toLowerCase();
    const filter = S.libFilter || "all";
    const sets = (S.bank && S.bank.sets) || [];

    const byDoc = new Map();
    sets.forEach(function (set) {
      const key = set.sourceFile || "unknown";
      if (!byDoc.has(key)) byDoc.set(key, []);
      byDoc.get(key).push(set);
    });

    let visibleSets = 0;
    byDoc.forEach(function (docSets, fileName) {
      const matching = docSets.filter(function (set) {
        const stats = setStats(set.id);
        if (filter === "untouched" && stats.started) return false;
        if (filter === "progress" && (!stats.started || stats.completed)) return false;
        if (filter === "completed" && !stats.completed) return false;
        if (!term) return true;
        const hay = (set.title + " " + fileName + " " + (set.description || "")).toLowerCase();
        return hay.indexOf(term) !== -1;
      });
      if (!matching.length) return;
      visibleSets += matching.length;
      const sec = el("section", "doc-group");
      const h2 = el("h2", "doc-group-title", docLabel(fileName));
      sec.appendChild(h2);
      const grid = el("div", "set-grid");
      matching.forEach(function (set) { grid.appendChild(setCard(set)); });
      sec.appendChild(grid);
      host.appendChild(sec);
    });

    if (!visibleSets) {
      const none = el("p", "lib-empty", term
        ? "No sets match your search or filter."
        : "This bank has no sets yet.");
      host.appendChild(none);
    }
  }

  function setCard(set) {
    const stats = setStats(set.id);
    const card = el("article", "set-card");
    card.dataset.status = stats.completed ? "completed" : (stats.started ? "progress" : "untouched");

    const top = el("div", "set-card-top");
    const h3 = el("h3", "set-title", set.title || set.id);
    top.appendChild(h3);
    const badge = el("span", "badge badge-" + card.dataset.status,
      stats.completed ? "Completed" : stats.started ? "In progress" : "Not started");
    top.appendChild(badge);
    card.appendChild(top);

    const meta = el("p", "set-meta");
    const siblings = (S.bank.sets || []).filter((s) => s.sourceFile === set.sourceFile);
    const setIdx = siblings.indexOf(set) + 1;
    meta.textContent = "Set " + setIdx + " of " + siblings.length + " in “" + docLabel(set.sourceFile) + "” · " + stats.total + " questions";
    card.appendChild(meta);
    if (set.description) {
      const desc = el("p", "set-desc", set.description);
      desc.title = set.description;
      card.appendChild(desc);
    }

    const bar = el("div", "pbar");
    bar.setAttribute("role", "progressbar");
    bar.setAttribute("aria-valuemin", "0"); bar.setAttribute("aria-valuemax", "100");
    bar.setAttribute("aria-valuenow", String(Math.round(stats.pct * 100)));
    bar.setAttribute("aria-label", "Completion: " + fmtPct(stats.pct));
    const fill = el("div", "pbar-fill" + (stats.completed ? " pbar-done" : ""));
    fill.style.width = fmtPct(stats.pct);
    bar.appendChild(fill);
    card.appendChild(bar);

    const nums = el("p", "set-nums");
    nums.textContent = stats.done + "/" + stats.total + " answered · " + stats.correct + " correct" +
      (stats.incorrect ? " · " + stats.incorrect + " incorrect" : "") + " · " + fmtPct(stats.pct) + " complete";
    card.appendChild(nums);

    const actions = el("div", "set-actions");
    const go = el("button", "btn btn-primary btn-small",
      stats.completed ? "Review set" : stats.started ? "Continue" : "Start");
    go.addEventListener("click", function () { navigate("#/exam/" + encodeURIComponent(set.id)); });
    actions.appendChild(go);

    const reset = el("button", "btn btn-ghost btn-small btn-reset", "Reset");
    reset.prepend(svgIcon("reset"));
    reset.setAttribute("aria-label", "Reset saved progress for " + (set.title || set.id));
    reset.addEventListener("click", async function () {
      const yes = await confirmDialog({
        title: "Reset this set?",
        body: "All saved answers, results and your position in “" + (set.title || set.id) +
          "” will be permanently cleared from this browser. Other sets are not affected. Your question bank is NOT deleted.",
        confirmLabel: "Reset set",
        danger: true,
      });
      if (yes) {
        delete S.progress[set.id];
        saveProgress();
        render();
        toast("Progress for this set was reset.");
      }
    });
    actions.appendChild(reset);
    card.appendChild(actions);
    return card;
  }

  /* -------------------------------- EXAM --------------------------------- */

  function viewExam(setId) {
    const set = bankSet(setId);
    if (!set) {
      const wrap = el("div", "view");
      wrap.appendChild(el("h1", "view-title", "Set not found"));
      wrap.appendChild(el("p", null, "The set “" + setId + "” is not in the active question bank."));
      const back = el("a", "btn btn-primary", "Back to library");
      back.href = "#/";
      wrap.appendChild(back);
      mainHost.appendChild(wrap);
      return;
    }
    const qs = setQuestions(setId);
    if (!qs.length) {
      const wrap = el("div", "view");
      wrap.appendChild(el("h1", "view-title", "This set has no questions"));
      mainHost.appendChild(wrap);
      return;
    }

    const prog = setProgress(setId);
    let index = S.exam && S.exam.setId === setId ? S.exam.index : prog.last;
    if (typeof index !== "number" || index < 0 || index >= qs.length) index = 0;
    S.exam = { setId: setId, index: index };

    const wrap = el("div", "view exam-view");

    /* header */
    const head = el("header", "exam-head");
    const left = el("div", "exam-head-left");
    const back = el("button", "btn btn-ghost btn-small", " Library");
    back.prepend(svgIcon("back"));
    back.addEventListener("click", function () { navigate("#/"); });
    left.appendChild(back);
    const titles = el("div", "exam-titles");
    titles.appendChild(el("div", "exam-doc", docLabel(set.sourceFile)));
    titles.appendChild(el("h1", "exam-title", set.title || set.id));
    left.appendChild(titles);
    head.appendChild(left);
    const right = el("div", "exam-head-right");
    const navToggle = el("button", "btn btn-ghost btn-small exam-nav-toggle", " Questions");
    navToggle.prepend(svgIcon("grid"));
    navToggle.setAttribute("aria-expanded", "false");
    navToggle.setAttribute("aria-controls", "qnav-drawer");
    navToggle.addEventListener("click", function () {
      const drawer = wrap.querySelector(".qnav-drawer");
      const open = drawer.classList.toggle("qnav-open");
      navToggle.setAttribute("aria-expanded", String(open));
    });
    right.appendChild(navToggle);
    head.appendChild(right);
    wrap.appendChild(head);

    /* progress row */
    const prow = el("div", "exam-progress");
    const ptext = el("span", "exam-progress-text");
    const pbar = el("div", "pbar pbar-thin");
    pbar.setAttribute("role", "progressbar");
    pbar.setAttribute("aria-valuemin", "0"); pbar.setAttribute("aria-valuemax", "100");
    const pfill = el("div", "pbar-fill");
    pbar.appendChild(pfill);
    prow.appendChild(ptext); prow.appendChild(pbar);
    wrap.appendChild(prow);

    /* body: question area + navigator */
    const body = el("div", "exam-body");
    const qarea = el("div", "question-area");
    body.appendChild(qarea);

    const drawer = el("aside", "qnav-drawer");
    drawer.id = "qnav-drawer";
    drawer.setAttribute("aria-label", "Question navigator");
    const drawerHead = el("div", "qnav-head");
    drawerHead.appendChild(el("h2", "qnav-title", "Questions"));
    const legend = el("div", "qnav-legend");
    [["un", "Unanswered"], ["ok", "Correct"], ["no", "Incorrect"], ["rv", "Reviewed"]].forEach(function (L) {
      const item = el("span", "qnav-legend-item");
      item.appendChild(el("span", "qnav-dot qnav-dot-" + L[0]));
      item.appendChild(el("span", null, L[1]));
      legend.appendChild(item);
    });
    drawerHead.appendChild(legend);
    drawer.appendChild(drawerHead);
    const grid = el("div", "qnav-grid");
    grid.setAttribute("role", "group");
    grid.setAttribute("aria-label", "Jump to question");
    drawer.appendChild(grid);
    const drawerFoot = el("div", "qnav-foot");
    const resetBtn = el("button", "btn btn-ghost btn-small", " Reset set");
    resetBtn.prepend(svgIcon("reset"));
    resetBtn.addEventListener("click", async function () {
      const yes = await confirmDialog({
        title: "Reset this set?",
        body: "All saved answers and results in “" + (set.title || set.id) + "” will be cleared from this browser. The question bank itself is not changed.",
        confirmLabel: "Reset set", danger: true,
      });
      if (yes) {
        delete S.progress[setId];
        saveProgress();
        S.exam = { setId: setId, index: 0 };
        render();
        toast("Progress for this set was reset.");
      }
    });
    drawerFoot.appendChild(resetBtn);
    drawer.appendChild(drawerFoot);
    body.appendChild(drawer);
    wrap.appendChild(body);

    mainHost.appendChild(wrap);

    function updateProgressRow() {
      const st = setStats(setId);
      ptext.textContent = "Question " + (S.exam.index + 1) + " of " + qs.length + " · " +
        st.done + " answered · " + st.correct + " correct";
      pbar.setAttribute("aria-valuenow", String(Math.round(st.pct * 100)));
      pfill.style.width = fmtPct(st.pct);
    }

    function updateNavigator() {
      grid.textContent = "";
      qs.forEach(function (q, i) {
        const rec = answerRec(setId, q.id);
        let state = "un";
        if (rec && rec.submitted) state = rec.correct === true ? "ok" : rec.correct === false ? "no" : "rv";
        else if (rec && rec.revealed) state = "rv";
        const b = el("button", "qnav-btn qnav-" + state + (i === S.exam.index ? " qnav-current" : ""));
        b.type = "button";
        b.textContent = String(i + 1);
        const statusWord = state === "ok" ? "answered, correct" : state === "no" ? "answered, incorrect" : state === "rv" ? "reviewed" : "unanswered";
        b.setAttribute("aria-label", "Question " + (i + 1) + " of " + qs.length + ", " + statusWord + ". Jump to question " + (i + 1) + ".");
        if (i === S.exam.index) b.setAttribute("aria-current", "true");
        b.addEventListener("click", function () { goTo(i); });
        grid.appendChild(b);
      });
    }

    function goTo(i) {
      if (i < 0 || i >= qs.length) return;
      S.exam.index = i;
      const prog2 = setProgress(setId);
      prog2.last = i;
      prog2.updated = Date.now();
      saveProgress();
      renderQuestion();
    }

    function imageBlock(im, q) {
      const wrap2 = el("figure", "q-img");
      const src = safeImgSrc(im.src);
      if (!src) {
        wrap2.appendChild(el("figcaption", "q-img-missing",
          "[Image could not be displayed" + (im.sourceName ? ": " + im.sourceName : "") + "]"));
        return wrap2;
      }
      const btn = el("button", "q-img-btn");
      btn.type = "button";
      btn.setAttribute("aria-label", "Enlarge image" + (im.caption ? ": " + im.caption : ""));
      const img = el("img", "q-img-el");
      img.src = src;
      img.alt = im.alt || im.caption || ("Examination image for question " + (q.questionNumber || ""));
      img.loading = "lazy";
      img.addEventListener("error", function () {
        wrap2.textContent = "";
        wrap2.appendChild(el("figcaption", "q-img-missing", "[Image failed to load: " + (im.sourceName || im.src) + "]"));
      });
      btn.appendChild(img);
      const zoomHint = el("span", "q-img-zoom");
      zoomHint.appendChild(svgIcon("zoom"));
      btn.appendChild(zoomHint);
      btn.addEventListener("click", function () { openLightbox(src, im.alt || im.caption || ""); });
      wrap2.appendChild(btn);
      if (im.caption) wrap2.appendChild(el("figcaption", "q-img-cap", im.caption));
      if (im.error) wrap2.appendChild(el("figcaption", "q-img-warn", "Source warning: " + im.error));
      return wrap2;
    }

    /* ------- multiple choice / SATA ------- */
    function renderChoiceQuestion(card, q, rec, refresh) {
      const submitted = !!(rec && rec.submitted);
      const chosen = (rec && rec.sel) ? rec.sel.slice() : [];

      const fieldset = el("fieldset", "opts");
      const legendEl = el("legend", "opts-legend",
        q.type === "sata" ? "Select all answers that apply:" : "Select one answer:");
      fieldset.appendChild(legendEl);

      const inputType = q.type === "sata" ? "checkbox" : "radio";
      (q.options || []).forEach(function (opt) {
        const label = el("label", "opt");
        const input = el("input");
        input.type = inputType;
        input.name = "opts-" + q.id;
        input.value = opt.label;
        input.id = "opt-" + q.id + "-" + opt.label;
        if (chosen.indexOf(opt.label) !== -1) input.checked = true;
        if (submitted) input.disabled = true;
        input.addEventListener("change", function () {
          const p = setProgress(setId);
          const cur = p.answers[q.id] || { sel: [] };
          if (q.type === "sata") {
            const sel = (cur.sel || []).slice();
            const at = sel.indexOf(opt.label);
            if (input.checked && at === -1) sel.push(opt.label);
            if (!input.checked && at !== -1) sel.splice(at, 1);
            cur.sel = sel;
          } else {
            cur.sel = [opt.label];
          }
          cur.updatedAt = Date.now();
          p.answers[q.id] = cur;
          p.updated = Date.now();
          saveProgress();
          label.classList.toggle("opt-selected", input.checked);
          if (inputType === "radio") {
            fieldset.querySelectorAll(".opt").forEach(function (l2) {
              const i2 = l2.querySelector("input");
              l2.classList.toggle("opt-selected", !!(i2 && i2.checked));
            });
          }
          updateNavigator(); updateProgressRow();
        });
        label.appendChild(input);
        label.appendChild(el("span", "opt-key", opt.label));
        label.appendChild(el("span", "opt-text", opt.text));

        if (submitted) {
          const isKey = (q.correctAnswers || []).indexOf(opt.label) !== -1;
          const picked = chosen.indexOf(opt.label) !== -1;
          if (isKey) label.classList.add("opt-correct");
          else if (picked) label.classList.add("opt-incorrect");
          if (picked) label.classList.add("opt-picked");
          const mark = el("span", "opt-mark");
          if (picked && isKey) { mark.appendChild(svgIcon("check")); mark.appendChild(el("span", null, "Your answer — correct")); mark.classList.add("opt-mark-ok"); }
          else if (picked && !isKey) { mark.appendChild(svgIcon("x")); mark.appendChild(el("span", null, "Your answer — incorrect")); mark.classList.add("opt-mark-no"); }
          else if (!picked && isKey) { mark.appendChild(svgIcon("check")); mark.appendChild(el("span", null, "Correct answer")); mark.classList.add("opt-mark-key"); }
          if (mark.childNodes.length) label.appendChild(mark);
        } else if (chosen.indexOf(opt.label) !== -1) {
          label.classList.add("opt-selected");
        }
        fieldset.appendChild(label);
      });
      card.appendChild(fieldset);

      /* actions row */
      const actRow = el("div", "q-actions");
      if (!submitted) {
        const submit = el("button", "btn btn-primary q-submit", "Submit answer");
        submit.addEventListener("click", function () {
          const p = setProgress(setId);
          const cur = p.answers[q.id];
          const sel = cur && cur.sel ? cur.sel : [];
          if (!sel.length) {
            toast(q.type === "sata" ? "Select at least one option before submitting." : "Select an option before submitting.", "warn");
            return;
          }
          const g = gradeQuestion(q, sel);
          p.answers[q.id] = {
            sel: sel.slice(), submitted: true, correct: g.correct, partial: g.partial,
            revealed: true, submittedAt: nowISO(), attempts: ((cur && cur.attempts) || 0) + 1,
          };
          p.updated = Date.now();
          saveProgress();
          refresh();
          requestAnimationFrame(function () {
            const banner = qarea.querySelector(".verdict");
            if (banner) banner.focus();
          });
        });
        actRow.appendChild(submit);
        if (chosen.length) actRow.appendChild(el("span", "q-actions-hint", "Selection saved — you can navigate away and return."));
      } else {
        const retry = el("button", "btn btn-ghost", "Try again (clears this attempt)");
        retry.addEventListener("click", async function () {
          const yes = await confirmDialog({
            title: "Try this question again?",
            body: "Your previous result for this question will be cleared so you can re-answer it. Other questions are unaffected.",
            confirmLabel: "Clear attempt",
          });
          if (!yes) return;
          const p = setProgress(setId);
          delete p.answers[q.id];
          p.updated = Date.now();
          saveProgress();
          refresh();
        });
        actRow.appendChild(retry);
      }
      card.appendChild(actRow);

      /* feedback */
      if (submitted) {
        const g = gradeQuestion(q, chosen);
        const verdict = el("div", "verdict " + (g.correct ? "verdict-ok" : "verdict-no"));
        verdict.setAttribute("tabindex", "-1");
        verdict.setAttribute("role", "status");
        verdict.appendChild(svgIcon(g.correct ? "check" : "x"));
        const vtxt = el("div", null);
        vtxt.appendChild(el("p", "verdict-line", g.correct ? "Correct" : "Incorrect"));
        const keySorted = (q.correctAnswers || []).slice().sort().join(", ");
        const detail = el("p", "verdict-detail");
        detail.textContent = "Correct answer: " + keySorted +
          (q.type === "sata" && !g.correct
            ? "  ·  you selected " + ((chosen.slice().sort().join(", ")) || "nothing") +
              (g.partial > 0 ? " (partial credit " + Math.round(g.partial * 100) + "%)" : "")
            : "");
        vtxt.appendChild(detail);
        verdict.appendChild(vtxt);
        card.appendChild(verdict);

        card.appendChild(el("h3", "rat-head", "Rationale for each option"));
        const ratList = el("div", "rat-list");
        (q.options || []).forEach(function (opt, i) {
          const item = el("div", "rat-item reveal");
          item.style.animationDelay = (i * 60) + "ms";
          const isKey = (q.correctAnswers || []).indexOf(opt.label) !== -1;
          const picked = chosen.indexOf(opt.label) !== -1;
          const headRow = el("div", "rat-item-head");
          headRow.appendChild(el("span", "rat-key " + (isKey ? "rat-key-ok" : "rat-key-no"), opt.label));
          headRow.appendChild(el("span", "rat-verdict " + (isKey ? "rat-verdict-ok" : "rat-verdict-no"),
            isKey ? "Correct option" : "Incorrect option"));
          if (picked) headRow.appendChild(el("span", "rat-picked", "you chose this"));
          item.appendChild(headRow);
          item.appendChild(el("p", "rat-opt-text", opt.text));
          item.appendChild(el("p", "rat-text", opt.rationale
            ? opt.rationale
            : "No rationale provided in the source material."));
          ratList.appendChild(item);
        });
        (q.images || []).filter(function (im) { return im.position === "rationale"; }).forEach(function (im) {
          ratList.appendChild(imageBlock(im, q));
        });
        (q.rationaleNotes || []).forEach(function (n) {
          ratList.appendChild(el("p", "rat-note", n));
        });
        card.appendChild(ratList);
      }

      if (q.sourceMapping) {
        const sm = el("details", "src-map");
        sm.appendChild(el("summary", null, "Source mapping"));
        sm.appendChild(el("p", "src-map-text", q.sourceMapping));
        card.appendChild(sm);
      }
    }

    /* ------- open-ended (self-graded) ------- */
    function renderOpenQuestion(card, q, rec, refresh) {
      const revealed = !!(rec && rec.revealed);
      const noteWrap = el("div", "open-note-wrap");
      const noteLabel = el("label", "open-note-label", "Your answer / study notes (saved privately in this browser, never graded automatically):");
      noteLabel.htmlFor = "open-note-" + q.id;
      const ta = el("textarea", "text-input open-note");
      ta.id = "open-note-" + q.id;
      ta.rows = 3;
      ta.value = (rec && rec.note) || "";
      ta.addEventListener("input", debounce(function () {
        const p = setProgress(setId);
        const cur = p.answers[q.id] || {};
        cur.note = ta.value;
        cur.updatedAt = Date.now();
        p.answers[q.id] = cur;
        p.updated = Date.now();
        saveProgress();
      }, 300));
      noteWrap.appendChild(noteLabel);
      noteWrap.appendChild(ta);
      card.appendChild(noteWrap);

      const actRow = el("div", "q-actions");
      if (!revealed) {
        const reveal = el("button", "btn btn-primary q-reveal", "Reveal faculty answer");
        reveal.addEventListener("click", function () {
          const p = setProgress(setId);
          const cur = p.answers[q.id] || {};
          cur.revealed = true;
          cur.correct = null;
          cur.revealedAt = nowISO();
          p.answers[q.id] = cur;
          p.updated = Date.now();
          saveProgress();
          refresh();
          requestAnimationFrame(function () {
            const banner = qarea.querySelector(".verdict");
            if (banner) banner.focus();
          });
        });
        actRow.appendChild(reveal);
        actRow.appendChild(el("span", "q-actions-hint", "Think or write your answer first — the model answer stays hidden until you reveal it."));
      } else {
        const verdict = el("div", "verdict verdict-open");
        verdict.setAttribute("tabindex", "-1");
        verdict.setAttribute("role", "status");
        verdict.appendChild(svgIcon("half"));
        const vtxt = el("div", null);
        vtxt.appendChild(el("p", "verdict-line", "Faculty answer revealed"));
        vtxt.appendChild(el("p", "verdict-detail", "Grade yourself honestly — this counts toward your statistics."));
        verdict.appendChild(vtxt);
        card.appendChild(verdict);

        const selfRow = el("div", "self-grade");
        selfRow.appendChild(el("span", "self-grade-label", "Did you know this?"));
        const knew = el("button", "btn btn-small " + (rec && rec.self === "knew" ? "btn-ok-active" : "btn-ghost"), "✓ I knew this");
        const missed = el("button", "btn btn-small " + (rec && rec.self === "missed" ? "btn-no-active" : "btn-ghost"), "✗ Needs review");
        function setSelf(val) {
          const p = setProgress(setId);
          const cur = p.answers[q.id] || { revealed: true };
          cur.revealed = true;
          cur.self = val;
          cur.correct = val === "knew";
          cur.submitted = true;
          p.answers[q.id] = cur;
          p.updated = Date.now();
          saveProgress();
          refresh();
        }
        knew.addEventListener("click", function () { setSelf("knew"); });
        missed.addEventListener("click", function () { setSelf("missed"); });
        selfRow.appendChild(knew); selfRow.appendChild(missed);
        actRow.appendChild(selfRow);

        const hide = el("button", "btn btn-ghost btn-small", "Hide answer again");
        hide.addEventListener("click", function () {
          const p = setProgress(setId);
          const note = p.answers[q.id] && p.answers[q.id].note;
          delete p.answers[q.id];
          if (note) p.answers[q.id] = { note: note, updatedAt: Date.now() }; // keep study notes
          p.updated = Date.now();
          saveProgress();
          refresh();
        });
        actRow.appendChild(hide);
      }
      card.appendChild(actRow);

      if (revealed) {
        const block = el("div", "open-answer reveal");
        block.appendChild(el("h3", "rat-head", "Faculty answer & rationale"));
        (q.answerBlock || []).forEach(function (b) {
          if (b.bullet != null) block.appendChild(el("p", "open-line open-bullet", "• " + b.text));
          else if (b.bold) block.appendChild(el("h4", "open-line open-sub", b.text));
          else block.appendChild(el("p", "open-line", b.text));
        });
        if (!q.answerBlock || !q.answerBlock.length) {
          block.appendChild(el("p", "open-line", "No rationale provided in the source material."));
        }
        (q.images || []).filter(function (im) { return im.position === "rationale"; }).forEach(function (im) {
          block.appendChild(imageBlock(im, q));
        });
        card.appendChild(block);
      }

      if (q.sourceMapping) {
        const sm = el("details", "src-map");
        sm.appendChild(el("summary", null, "Source mapping"));
        sm.appendChild(el("p", "src-map-text", q.sourceMapping));
        card.appendChild(sm);
      }
    }

    function renderQuestion() {
      const q = qs[S.exam.index];
      const rec = answerRec(setId, q.id);
      qarea.textContent = "";
      updateProgressRow();
      updateNavigator();

      const card = el("article", "q-card");

      const metaRow = el("div", "q-meta");
      metaRow.appendChild(el("span", "q-number", "Question " + (S.exam.index + 1) + " of " + qs.length));
      if (q.questionNumber && q.questionNumber !== S.exam.index + 1) {
        metaRow.appendChild(el("span", "q-srcnum", "source #" + q.questionNumber));
      }
      const typeLabel = q.type === "mc" ? "Multiple Choice" : q.type === "sata" ? "Select All That Apply" : q.type === "tf" ? "True / False" : "Open-Ended Case (self-graded)";
      metaRow.appendChild(el("span", "type-badge type-" + q.type, typeLabel));
      if (rec && rec.submitted) metaRow.appendChild(el("span", "q-state q-state-sub", "Submitted"));
      card.appendChild(metaRow);

      if (q.title) card.appendChild(el("h2", "q-title", q.title));

      const stemWrap = el("div", "q-stem");
      stemWrap.setAttribute("tabindex", "-1");
      (q.stem || []).forEach(function (para) { stemWrap.appendChild(el("p", "q-stem-p", para)); });
      card.appendChild(stemWrap);

      (q.images || []).filter(function (im) { return im.position !== "rationale"; }).forEach(function (im) {
        card.appendChild(imageBlock(im, q));
      });

      if (q.type === "open") {
        renderOpenQuestion(card, q, rec, function () { renderQuestion(); });
      } else {
        renderChoiceQuestion(card, q, rec, function () { renderQuestion(); });
      }

      const nav = el("div", "q-nav");
      const prev = el("button", "btn btn-ghost q-prev", "← Previous");
      prev.disabled = S.exam.index === 0;
      prev.addEventListener("click", function () { goTo(S.exam.index - 1); });
      const next = el("button", "btn btn-ghost q-next", "Next →");
      next.disabled = S.exam.index === qs.length - 1;
      next.addEventListener("click", function () { goTo(S.exam.index + 1); });
      nav.appendChild(prev);
      const spacer = el("div", "q-nav-mid");
      if (S.exam.index === qs.length - 1) {
        const done = el("button", "btn btn-primary", "Finish & review set");
        done.addEventListener("click", function () { navigate("#/"); });
        spacer.appendChild(done);
      }
      nav.appendChild(spacer);
      nav.appendChild(next);
      card.appendChild(nav);

      qarea.appendChild(card);
      stemWrap.focus({ preventScroll: true });
    }

    renderQuestion();
  }

  /* ------------------------------- IMPORT -------------------------------- */

  const ImportFlow = {
    files: [],
    draft: null,
    step: 1,
  };

  function viewImport() {
    const wrap = el("div", "view import-view");
    wrap.appendChild(el("h1", "view-title", "Import Word Documents"));
    wrap.appendChild(el("p", "view-sub",
      "Convert your .docx examination files into the study bank used by this app. Everything happens in your browser — no upload, no AI, no server."));

    const flow = el("ol", "steps");
    [["1", "Choose documents"], ["2", "Review & correct"], ["3", "Publish"]].forEach(function (s, i) {
      const li = el("li", "step" + (ImportFlow.step === i + 1 ? " step-active" : "") + (ImportFlow.step > i + 1 ? " step-done" : ""));
      li.appendChild(el("span", "step-num", s[0]));
      li.appendChild(el("span", null, s[1]));
      flow.appendChild(li);
    });
    wrap.appendChild(flow);

    if (S.bankSource === "local") {
      const note = el("div", "banner banner-info");
      note.appendChild(el("p", null, "Reminder: importing only updates THIS browser. To share with classmates you must download the updated question-bank.json / question-bank.js and upload them to your GitHub repository (step 3)."));
      wrap.appendChild(note);
    }

    /* step 1: picker */
    const picker = el("section", "import-step");
    const drop = el("div", "dropzone");
    drop.setAttribute("tabindex", "0");
    drop.setAttribute("role", "button");
    drop.setAttribute("aria-label", "Choose or drop .docx files");
    drop.appendChild(svgIcon("upload"));
    drop.appendChild(el("p", "drop-main", "Drop .docx files here"));
    drop.appendChild(el("p", "drop-sub", "or"));
    const choose = el("button", "btn btn-primary", "Choose files…");
    const input = el("input", "file-input sr-only");
    input.type = "file";
    input.accept = ".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    input.multiple = true;
    choose.addEventListener("click", function (e) { e.stopPropagation(); input.click(); });
    drop.appendChild(choose);
    drop.appendChild(input);
    drop.addEventListener("click", function (e) { if (e.target !== choose) input.click(); });
    drop.addEventListener("keydown", function (e) { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); input.click(); } });
    ["dragenter", "dragover"].forEach(function (ev) {
      drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add("drop-hover"); });
    });
    ["dragleave", "drop"].forEach(function (ev) {
      drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove("drop-hover"); });
    });
    drop.addEventListener("drop", function (e) {
      const files = Array.prototype.slice.call((e.dataTransfer && e.dataTransfer.files) || []);
      if (files.length) startParse(files);
    });
    input.addEventListener("change", function () {
      const files = Array.prototype.slice.call(input.files || []);
      if (files.length) startParse(files);
      input.value = "";
    });
    picker.appendChild(drop);
    const fileList = el("ul", "import-files");
    picker.appendChild(fileList);
    wrap.appendChild(picker);

    const reviewHost = el("section", "import-step");
    wrap.appendChild(reviewHost);
    const publishHost = el("section", "import-step");
    wrap.appendChild(publishHost);

    function renderFileList() {
      fileList.textContent = "";
      ImportFlow.files.forEach(function (f) {
        const li = el("li", "import-file import-file-" + f.status);
        li.appendChild(el("span", "import-file-name", f.file.name));
        const st = el("span", "import-file-status");
        if (f.status === "parsing") st.textContent = "reading…";
        else if (f.status === "done") {
          st.textContent = f.result.sets.length + " set(s), " + f.result.questions.length + " question(s)" +
            (f.result.stats.errorCount ? ", " + f.result.stats.errorCount + " problem(s)" : "");
        } else if (f.status === "error") st.textContent = "error: " + f.error;
        li.appendChild(st);
        fileList.appendChild(li);
      });
    }

    async function startParse(files) {
      ImportFlow.step = 2;
      const docxFiles = files.filter(function (f) { return /\.docx$/i.test(f.name); });
      files.filter(function (f) { return !/\.docx$/i.test(f.name); }).forEach(function (f) {
        ImportFlow.files.push({ file: f, status: "error", error: "not a .docx file" });
      });
      docxFiles.forEach(function (f) { ImportFlow.files.push({ file: f, status: "parsing" }); });
      renderFileList();

      for (const entry of ImportFlow.files.filter(function (f) { return f.status === "parsing"; })) {
        try {
          const buf = await entry.file.arrayBuffer();
          const ir = await window.DocxIR.readDocxIR(buf, { JSZip: window.JSZip, DOMParser: window.DOMParser }, entry.file.name, {});
          await downscaleImages(ir);
          const result = window.ExamParser.parseExamDocuments([ir]);
          entry.ir = ir;
          entry.result = result;
          entry.status = "done";
        } catch (e) {
          entry.status = "error";
          entry.error = (e && e.message) ? e.message : String(e);
          console.error("Import failed for", entry.file.name, e);
        }
        renderFileList();
      }
      buildDraft();
      render();
    }

    async function downscaleImages(ir) {
      const MAX = 1600, LIMIT = 250000;
      for (const p of ir.paragraphs) {
        for (const im of (p.images || [])) {
          if (!im.dataUrl || im.error) continue;
          if (im.dataUrl.length < LIMIT * 1.37) continue;
          try {
            const smaller = await scaleDataUrl(im.dataUrl, MAX);
            if (smaller && smaller.length < im.dataUrl.length) im.dataUrl = smaller;
          } catch (e) { /* keep original */ }
        }
      }
    }
    function scaleDataUrl(dataUrl, maxDim) {
      return new Promise(function (resolve) {
        const img = new Image();
        img.onload = function () {
          try {
            const scale = Math.min(1, maxDim / Math.max(img.width, img.height));
            if (scale >= 1) return resolve(null);
            const canvas = document.createElement("canvas");
            canvas.width = Math.round(img.width * scale);
            canvas.height = Math.round(img.height * scale);
            const ctx = canvas.getContext("2d");
            ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
            const isPng = dataUrl.indexOf("data:image/png") === 0;
            let out = canvas.toDataURL(isPng ? "image/png" : "image/jpeg", 0.85);
            if (isPng && out.length >= dataUrl.length) out = canvas.toDataURL("image/jpeg", 0.85);
            resolve(out.length < dataUrl.length ? out : null);
          } catch (e) { resolve(null); }
        };
        img.onerror = function () { resolve(null); };
        img.src = dataUrl;
      });
    }

    function normTitle(t) { return String(t || "").toLowerCase().replace(/\s+/g, " ").trim(); }
    function setFingerprint(qs) {
      const hashes = qs.map(function (q) { return (q.metadata && q.metadata.contentHash) || ""; }).sort();
      return window.ExamParser._internals.hashText(hashes.join("|"));
    }

    function buildDraft() {
      const parsed = ImportFlow.files.filter(function (f) { return f.status === "done" && f.result; });
      if (!parsed.length) { ImportFlow.draft = null; return; }
      const sets = [], questions = [], documents = [], warnings = [];
      parsed.forEach(function (f) {
        f.result.sets.forEach(function (s) { sets.push(s); });
        f.result.questions.forEach(function (q) { questions.push(q); });
        f.result.documents.forEach(function (d) { documents.push(d); });
        f.result.warnings.forEach(function (w) { warnings.push(w); });
      });
      const dupes = [];
      const existing = (S.bank && S.bank.sets) || [];
      sets.forEach(function (s) {
        const sameFileTitle = existing.filter(function (e2) {
          return e2.sourceFile === s.sourceFile && normTitle(e2.title) === normTitle(s.title);
        });
        const fpNew = setFingerprint(questions.filter(function (q) { return q.setId === s.id; }));
        const sameFp = existing.filter(function (e2) { return setFingerprint(setQuestions(e2.id)) === fpNew; });
        if (sameFileTitle.length || sameFp.length) {
          dupes.push({
            setId: s.id,
            matches: (sameFileTitle.length ? sameFileTitle : sameFp).map(function (m) { return m.id; }),
            mode: "replace",
          });
        }
      });
      ImportFlow.draft = { sets: sets, questions: questions, documents: documents, warnings: warnings, dupes: dupes };
    }

    function questionEditor(q) {
      const box = el("div", "q-edit");
      const head = el("div", "q-edit-head");
      head.appendChild(el("span", "q-edit-num", "Q" + q.questionNumber));
      head.appendChild(el("span", "q-edit-issues", (q.issues || []).join(", ")));
      box.appendChild(head);
      const stemText = (q.stem || []).join(" ");
      box.appendChild(el("p", "q-edit-stem", stemText.slice(0, 220) + (stemText.length > 220 ? "…" : "")));

      const row = el("div", "q-edit-row");
      const typeLab = el("label", "q-edit-lab", "Type:");
      const typeSel = el("select", "text-input");
      [["mc", "Multiple choice"], ["sata", "Select all that apply"], ["tf", "True/False"], ["open", "Open-ended"]].forEach(function (t) {
        const o = el("option", null, t[1]);
        o.value = t[0];
        if (q.type === t[0]) o.selected = true;
        typeSel.appendChild(o);
      });
      typeSel.addEventListener("change", function () {
        q.type = typeSel.value;
        q.issues = (q.issues || []).filter(function (i) {
          return i.indexOf("type-") !== 0 && i.indexOf("missing-rationale") !== 0;
        });
      });
      typeLab.appendChild(typeSel);
      row.appendChild(typeLab);

      if (q.type !== "open") {
        row.appendChild(el("span", "q-edit-lab", "Correct:"));
        (q.options || []).forEach(function (opt) {
          const lab = el("label", "q-edit-key");
          const cb = el("input");
          cb.type = "checkbox";
          cb.checked = (q.correctAnswers || []).indexOf(opt.label) !== -1;
          cb.addEventListener("change", function () {
            const arr = (q.correctAnswers || []).slice();
            const at = arr.indexOf(opt.label);
            if (cb.checked && at === -1) arr.push(opt.label);
            if (!cb.checked && at !== -1) arr.splice(at, 1);
            q.correctAnswers = arr.sort();
            q.options.forEach(function (o) { o.isCorrect = q.correctAnswers.indexOf(o.label) !== -1; });
            q.issues = (q.issues || []).filter(function (i) {
              return i.indexOf("missing-key") !== 0 && i.indexOf("key-label") !== 0;
            });
          });
          lab.appendChild(cb);
          lab.appendChild(el("span", null, opt.label));
          row.appendChild(lab);
        });
      }
      box.appendChild(row);
      return box;
    }

    function renderReview(host) {
      host.textContent = "";
      if (ImportFlow.step < 2 || !ImportFlow.draft) return;
      const d = ImportFlow.draft;
      host.appendChild(el("h2", "step-title", "Review detected sets"));

      const errCount = d.warnings.filter(function (w) { return w.level === "error"; }).length;
      const warnCount = d.warnings.filter(function (w) { return w.level === "warn"; }).length;
      const summary = el("div", "banner " + (errCount ? "banner-warn" : "banner-ok"));
      summary.appendChild(el("p", null, d.sets.length + " set(s) · " + d.questions.length +
        " question(s) detected · " + errCount + " error-level issue(s) · " + warnCount + " warning(s). " +
        (errCount ? "Review the flagged questions below before publishing." : "No blocking problems found.")));
      host.appendChild(summary);

      const grouped = new Map();
      d.warnings.forEach(function (w) {
        const key = (w.file || "?") + "|" + (w.setId || "?");
        if (!grouped.has(key)) grouped.set(key, []);
        grouped.get(key).push(w);
      });

      d.sets.forEach(function (set) {
        const panel = el("details", "set-panel" + (set.metadata.confidence === "review" ? " set-panel-review" : ""));
        panel.open = set.metadata.confidence === "review";
        const sum = el("summary", "set-panel-sum");
        sum.appendChild(el("span", "badge " + (set.metadata.confidence === "review" ? "badge-review" : "badge-ok"),
          set.metadata.confidence === "review" ? "boundary needs review" : "detected: " + set.metadata.detectedBy.join(", ")));
        const qsInSet = d.questions.filter(function (q) { return q.setId === set.id; });
        const setWarns = (grouped.get((set.sourceFile || "?") + "|" + set.id) || []);
        sum.appendChild(el("span", "set-panel-count", qsInSet.length + " questions · " + setWarns.length + " issue(s)"));
        panel.appendChild(sum);

        const titleLab = el("label", "set-panel-titlelab", "Set title:");
        const titleInput = el("input", "text-input");
        titleInput.value = set.title;
        titleInput.setAttribute("aria-label", "Set title (was: " + set.title + ")");
        titleInput.addEventListener("input", function () { set.title = titleInput.value; });
        titleLab.appendChild(titleInput);
        panel.appendChild(titleLab);
        if (set.metadata.autoTitle) {
          panel.appendChild(el("p", "set-panel-note", "No explicit set heading was found; this title was generated. You can rename it."));
        }

        const flagged = qsInSet.filter(function (q) { return q.issues && q.issues.length; });
        if (flagged.length) {
          panel.appendChild(el("h4", "set-panel-issues", "Flagged questions (" + flagged.length + ")"));
          flagged.forEach(function (q) { panel.appendChild(questionEditor(q)); });
        } else {
          panel.appendChild(el("p", "set-panel-ok", "All questions in this set parsed cleanly."));
        }
        host.appendChild(panel);
      });

      if (d.dupes.length) {
        const dupBox = el("div", "dup-box");
        dupBox.appendChild(el("h3", null, "Possible duplicate sets"));
        dupBox.appendChild(el("p", null, "These detected sets already exist in your active bank. Choose what to do for each:"));
        d.dupes.forEach(function (dup) {
          const row = el("div", "dup-row");
          const set = d.sets.find(function (s) { return s.id === dup.setId; });
          row.appendChild(el("span", "dup-name", set ? set.title : dup.setId));
          ["skip", "replace", "copy"].forEach(function (mode) {
            const lab = el("label", "dup-opt");
            const rb = el("input");
            rb.type = "radio";
            rb.name = "dup-" + dup.setId;
            rb.value = mode;
            rb.checked = dup.mode === mode;
            rb.addEventListener("change", function () { dup.mode = mode; });
            lab.appendChild(rb);
            lab.appendChild(el("span", null, mode === "skip" ? "Skip (keep existing)" : mode === "replace" ? "Replace existing" : "Add as extra copy"));
            row.appendChild(lab);
          });
          dupBox.appendChild(row);
        });
        host.appendChild(dupBox);
      }

      const actions = el("div", "import-actions");
      const accept = el("button", "btn btn-primary import-accept", "Accept import into this browser");
      if (errCount) {
        const ack = el("label", "ack-label");
        const ackBox = el("input");
        ackBox.type = "checkbox";
        ackBox.addEventListener("change", function () { accept.disabled = !ackBox.checked; });
        ack.appendChild(ackBox);
        ack.appendChild(el("span", null, "I understand there are unresolved structural issues and want to import anyway."));
        accept.disabled = true;
        actions.appendChild(ack);
      }
      accept.addEventListener("click", function () { applyImport(); });
      const cancel = el("button", "btn btn-ghost", "Discard draft");
      cancel.addEventListener("click", function () {
        ImportFlow.files = [];
        ImportFlow.draft = null;
        ImportFlow.step = 1;
        render();
      });
      actions.appendChild(accept);
      actions.appendChild(cancel);
      host.appendChild(actions);
    }

    function applyImport() {
      const d = ImportFlow.draft;
      if (!d) return;
      let base;
      if (!S.bank) {
        base = { format: "exam-bank", version: 1, title: "Practice Examination Bank", documents: [], sets: [], questions: [] };
      } else {
        base = JSON.parse(JSON.stringify(S.bank));
      }
      const byId = new Map(base.sets.map(function (s) { return [s.id, s]; }));

      d.dupes.forEach(function (dup) {
        if (dup.mode === "skip") {
          dup.skip = true;
        } else if (dup.mode === "replace") {
          dup.matches.forEach(function (mid) {
            const old = byId.get(mid);
            if (old) {
              const oldIds = new Set(old.questionIds || []);
              base.questions = base.questions.filter(function (q) { return !oldIds.has(q.id); });
              base.sets = base.sets.filter(function (s) { return s.id !== mid; });
              byId.delete(mid);
            }
          });
        } else if (dup.mode === "copy") {
          dup.renumber = true;
        }
      });

      // apply copy-renumbering to draft sets & questions first
      d.sets.forEach(function (set) {
        const dup = d.dupes.find(function (x) { return x.setId === set.id; });
        if (dup && dup.renumber && !dup.suffix) {
          let n = 2;
          while (byId.has(set.id + "--c" + n)) n++;
          dup.suffix = "--c" + n;
        }
      });
      d.questions.forEach(function (q) {
        const dup = d.dupes.find(function (x) { return x.setId === q.setId; });
        if (dup && dup.renumber && dup.suffix) {
          q.setId = q.setId + dup.suffix;
          q.id = q.id + dup.suffix;
        }
      });
      d.sets.forEach(function (set) {
        const dup = d.dupes.find(function (x) { return x.setId === set.id; });
        if (dup && dup.skip) return;
        const newSet = JSON.parse(JSON.stringify(set));
        if (dup && dup.renumber && dup.suffix) {
          newSet.id = newSet.id + dup.suffix;
          newSet.questionIds = newSet.questionIds.map(function (qid) { return qid + dup.suffix; });
        }
        if (byId.has(newSet.id)) {
          const oldIds = new Set((byId.get(newSet.id).questionIds) || []);
          base.questions = base.questions.filter(function (q) { return !oldIds.has(q.id); });
          base.sets = base.sets.filter(function (s) { return s.id !== newSet.id; });
        }
        base.sets.push(newSet);
        byId.set(newSet.id, newSet);
      });
      d.questions.forEach(function (q) {
        const owningSet = base.sets.find(function (s) { return s.id === q.setId; });
        if (!owningSet) return;
        if ((owningSet.questionIds || []).indexOf(q.id) === -1) return;
        base.questions = base.questions.filter(function (x) { return x.id !== q.id; });
        base.questions.push(JSON.parse(JSON.stringify(q)));
      });
      d.documents.forEach(function (doc) {
        const at = base.documents.findIndex(function (x) { return x.file === doc.file; });
        if (at !== -1) base.documents[at] = doc; else base.documents.push(doc);
      });
      base.generatedAt = nowISO();
      base.bankId = "bank-" + window.ExamParser._internals.hashText(
        base.sets.map(function (s) { return s.id + ":" + s.questionIds.length; }).join("|") + "#" +
        base.questions.map(function (q) { return (q.metadata && q.metadata.contentHash) || ""; }).join("|")
      ).slice(0, 12);

      const v = window.ExamParser.validateBank(base);
      if (!v.valid) {
        toast("Import produced an invalid bank: " + v.errors[0], "error", 8000);
        console.error(v.errors);
        return;
      }
      S.bank = base;
      S.bankSource = "local";
      const res = Storage.set(LS_KEYS.bank, base);
      Storage.set(LS_KEYS.bankMeta, { savedAt: nowISO(), source: "import" });
      if (!res.ok) {
        toast("Bank is active in this session but too large to keep in browser storage (" + res.reason + "). Download it below and publish it to keep it.", "warn", 9000);
      } else {
        toast("Import accepted — the bank is now active in this browser.");
      }
      ImportFlow.step = 3;
      ImportFlow.draft = null;
      render();
    }

    function renderPublish(host) {
      host.textContent = "";
      if (ImportFlow.step < 3 || S.bankSource !== "local") return;
      host.appendChild(el("h2", "step-title", "Publish your updated bank"));
      const box = el("div", "publish-box");
      box.appendChild(el("p", null, "Your imported bank is active in THIS browser only. To share it with classmates:"));
      const ol = el("ol", "publish-steps");
      [
        "Download the two files below.",
        "On github.com open your repository, click the file question-bank.json, click the pencil (edit) or trash icon, then upload the new file. Keep a backup copy of the old file on your computer first. Repeat for question-bank.js.",
        "Wait about a minute for GitHub Pages to rebuild, then refresh your published site and confirm the new sets appear.",
      ].forEach(function (t) { ol.appendChild(el("li", null, t)); });
      box.appendChild(ol);
      const btns = el("div", "publish-actions");
      const dl = el("button", "btn btn-primary", " Download question-bank.json");
      dl.prepend(svgIcon("download"));
      dl.addEventListener("click", function () { downloadBank("json"); });
      const dl2 = el("button", "btn btn-primary", " Download question-bank.js");
      dl2.prepend(svgIcon("download"));
      dl2.addEventListener("click", function () { downloadBank("js"); });
      const backLib = el("a", "btn btn-ghost", "Go study now");
      backLib.href = "#/";
      btns.appendChild(dl); btns.appendChild(dl2); btns.appendChild(backLib);
      box.appendChild(btns);
      host.appendChild(box);
    }

    renderReview(reviewHost);
    renderPublish(publishHost);
    mainHost.appendChild(wrap);
    // test/automation hook (harmless in production)
    ImportFlow._api = { startParse: startParse, buildDraft: buildDraft, applyImport: applyImport };
  }

  /* ------------------------------ SETTINGS ------------------------------- */

  function viewSettings() {
    const wrap = el("div", "view settings-view");
    wrap.appendChild(el("h1", "view-title", "Settings & Data"));

    const bankCard = el("section", "settings-card");
    bankCard.appendChild(el("h2", null, "Question bank"));
    if (S.bank) {
      const info = el("ul", "settings-list");
      info.appendChild(el("li", null, "Source: " + (S.bankSource === "local" ? "imported in this browser (not published)" : S.bankSource === "demo" ? "demo fixture" : "published question-bank.json")));
      info.appendChild(el("li", null, "Bank ID: " + (S.bank.bankId || "n/a")));
      info.appendChild(el("li", null, "Sets: " + (S.bank.sets || []).length + " · Questions: " + (S.bank.questions || []).length));
      if (S.bank.generatedAt) info.appendChild(el("li", null, "Generated: " + new Date(S.bank.generatedAt).toLocaleString()));
      bankCard.appendChild(info);
    } else {
      bankCard.appendChild(el("p", null, "No bank loaded."));
    }
    const bankActions = el("div", "settings-actions");
    const expJ = el("button", "btn btn-small btn-ghost", "Download question-bank.json");
    expJ.addEventListener("click", function () { downloadBank("json"); });
    const expJs = el("button", "btn btn-small btn-ghost", "Download question-bank.js");
    expJs.addEventListener("click", function () { downloadBank("js"); });
    bankActions.appendChild(expJ); bankActions.appendChild(expJs);

    const impBank = el("button", "btn btn-small btn-ghost", "Import bank from JSON file…");
    const bankFile = el("input", "sr-only");
    bankFile.type = "file"; bankFile.accept = ".json,application/json";
    bankFile.addEventListener("change", function () {
      const f = bankFile.files && bankFile.files[0];
      if (!f) return;
      const reader = new FileReader();
      reader.onload = async function () {
        let data = null;
        try { data = JSON.parse(String(reader.result)); } catch (e) {
          toast("That file is not valid JSON: " + e.message, "error"); return;
        }
        const v = window.ExamParser.validateBank(data);
        if (!v.valid) {
          toast("File is not a valid question bank: " + v.errors[0], "error", 8000);
          return;
        }
        const yes = await confirmDialog({
          title: "Replace active bank?",
          body: "This makes the imported file the active bank in this browser (and saves it to browser storage). Your study progress is kept where question IDs still match. The published website is not changed.",
          confirmLabel: "Use this bank",
        });
        if (!yes) return;
        S.bank = data;
        S.bankSource = "local";
        Storage.set(LS_KEYS.bank, data);
        Storage.set(LS_KEYS.bankMeta, { savedAt: nowISO(), source: "json-import" });
        render();
        toast("Imported bank is now active.");
      };
      reader.readAsText(f);
      bankFile.value = "";
    });
    bankActions.appendChild(impBank); bankActions.appendChild(bankFile);

    if (S.bankSource === "local") {
      const usePub = el("button", "btn btn-small btn-ghost", "Switch back to published bank");
      usePub.addEventListener("click", async function () {
        const yes = await confirmDialog({
          title: "Use published bank?",
          body: "Your locally imported bank will be removed from this browser (download a backup first if you need it). Study progress stays saved.",
          confirmLabel: "Switch to published",
        });
        if (!yes) return;
        Storage.remove(LS_KEYS.bank); Storage.remove(LS_KEYS.bankMeta);
        await loadBank(true);
        render();
      });
      bankActions.appendChild(usePub);
    }
    bankCard.appendChild(bankActions);
    wrap.appendChild(bankCard);

    const progCard = el("section", "settings-card");
    progCard.appendChild(el("h2", null, "Study progress"));
    const totalAnswered = Object.keys(S.progress).reduce(function (acc, setId) {
      const p = S.progress[setId];
      return acc + Object.keys(p.answers || {}).filter(function (qid) {
        return p.answers[qid].submitted || p.answers[qid].revealed;
      }).length;
    }, 0);
    progCard.appendChild(el("p", null, totalAnswered + " answered/reviewed question(s) stored in this browser."));
    const progActions = el("div", "settings-actions");
    const expProg = el("button", "btn btn-small btn-ghost", "Export progress backup");
    expProg.addEventListener("click", exportProgress);
    const impProg = el("button", "btn btn-small btn-ghost", "Restore progress from backup…");
    const progFile = el("input", "sr-only");
    progFile.type = "file"; progFile.accept = ".json,application/json";
    progFile.addEventListener("change", function () {
      const f = progFile.files && progFile.files[0];
      if (!f) return;
      const reader = new FileReader();
      reader.onload = async function () {
        let data = null;
        try { data = JSON.parse(String(reader.result)); } catch (e) { toast("Invalid JSON file.", "error"); return; }
        if (!data || data.format !== "exam-progress") { toast("Not a progress backup file.", "error"); return; }
        const yes = await confirmDialog({
          title: "Restore progress?",
          body: "This REPLACES your current saved progress in this browser with the backup file.",
          confirmLabel: "Restore",
          danger: true,
        });
        if (!yes) return;
        S.progress = data.progress || {};
        saveProgress();
        render();
        toast("Progress restored from backup.");
      };
      reader.readAsText(f);
      progFile.value = "";
    });
    progActions.appendChild(expProg); progActions.appendChild(impProg); progActions.appendChild(progFile);

    const resetAll = el("button", "btn btn-small btn-danger", "Reset ALL progress");
    resetAll.addEventListener("click", async function () {
      const yes = await confirmDialog({
        title: "Reset all progress?",
        body: "Every saved answer, result and position in every set will be deleted from this browser. Your question bank is NOT affected. This cannot be undone.",
        confirmLabel: "Delete everything",
        danger: true,
        requireText: "RESET",
      });
      if (!yes) return;
      S.progress = {};
      saveProgress();
      render();
      toast("All progress was reset.");
    });
    progActions.appendChild(resetAll);
    progCard.appendChild(progActions);
    wrap.appendChild(progCard);

    const prefCard = el("section", "settings-card");
    prefCard.appendChild(el("h2", null, "Appearance"));
    const themeLab = el("label", "settings-lab", "Theme:");
    const themeSel = el("select", "text-input");
    [["system", "Follow my device"], ["light", "Light"], ["dark", "Dark"]].forEach(function (t) {
      const o = el("option", null, t[1]);
      o.value = t[0];
      if (S.prefs.theme === t[0]) o.selected = true;
      themeSel.appendChild(o);
    });
    themeSel.addEventListener("change", function () {
      S.prefs.theme = themeSel.value;
      savePrefs();
      applyTheme();
    });
    themeLab.appendChild(themeSel);
    prefCard.appendChild(themeLab);
    wrap.appendChild(prefCard);

    const about = el("section", "settings-card");
    about.appendChild(el("h2", null, "Privacy & storage"));
    about.appendChild(el("p", null,
      "All study progress lives in this browser's localStorage only. Nothing is sent to a server, and nothing syncs between devices or browsers. The question bank is a static file; technically anyone can open it and read the answers, so do not use this app for secure exam administration."));
    if (!Storage.available) {
      const warn = el("div", "banner banner-warn");
      warn.appendChild(el("p", null, "localStorage is unavailable in this browser/context (private mode or file restrictions). Your progress will NOT survive a page reload."));
      about.appendChild(warn);
    }
    wrap.appendChild(about);

    mainHost.appendChild(wrap);
  }

  /* --------------------------- export helpers ---------------------------- */

  function download(filename, text, mime) {
    const blob = new Blob([text], { type: mime || "application/json" });
    const url = URL.createObjectURL(blob);
    const a = el("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 400);
  }

  function downloadBank(kind) {
    if (!S.bank) { toast("No bank to export.", "warn"); return; }
    if (kind === "js") {
      download("question-bank.js", "/* Mirror of question-bank.json for file:// fallback */\nwindow.EXAM_BANK = " + JSON.stringify(S.bank) + ";\n", "text/javascript");
    } else {
      download("question-bank.json", JSON.stringify(S.bank, null, 2), "application/json");
    }
    toast("Downloaded. Remember: upload it to GitHub to publish (see README).");
  }

  function exportProgress() {
    const payload = { format: "exam-progress", version: 1, exportedAt: nowISO(), progress: S.progress };
    download("exam-progress-backup.json", JSON.stringify(payload, null, 2));
    toast("Progress backup downloaded.");
  }

  /* ------------------------------ demo bank ------------------------------ */

  function demoBank() {
    return {
      format: "exam-bank", version: 1, bankId: "demo", title: "Demo bank",
      generatedAt: nowISO(),
      documents: [{ file: "demo-fixture.json", title: "Built-in demo fixture", importedAt: nowISO(), paragraphCount: 0, setIds: ["demo--s1"] }],
      sets: [{
        id: "demo--s1", title: "Demo Set — 4 sample questions", sourceFile: "demo-fixture.json",
        description: "A tiny labeled fixture so you can try the interface.",
        questionIds: ["demo--s1--q001", "demo--s1--q002", "demo--s1--q003", "demo--s1--q004"],
        metadata: { setIndex: 1, headingText: "", detectedBy: ["demo"], confidence: "high", autoTitle: false },
      }],
      questions: [
        {
          id: "demo--s1--q001", setId: "demo--s1", sourceFile: "demo-fixture.json", questionNumber: 1, ordinal: 1, type: "mc",
          stem: ["Which number is prime?"], options: [
            { label: "A", text: "9", isCorrect: false, rationale: "9 = 3 × 3." },
            { label: "B", text: "11", isCorrect: true, rationale: "11 has no divisors besides 1 and itself." },
            { label: "C", text: "15", isCorrect: false, rationale: "15 = 3 × 5." },
            { label: "D", text: "21", isCorrect: false, rationale: "21 = 3 × 7." },
          ],
          correctAnswers: ["B"], sourceMapping: "Demo fixture", metadata: { contentHash: "demo1" },
        },
        {
          id: "demo--s1--q002", setId: "demo--s1", sourceFile: "demo-fixture.json", questionNumber: 2, ordinal: 2, type: "sata",
          stem: ["Select all even numbers."], options: [
            { label: "A", text: "2", isCorrect: true, rationale: "2 is even." },
            { label: "B", text: "3", isCorrect: false, rationale: "3 is odd." },
            { label: "C", text: "4", isCorrect: true, rationale: "4 is even." },
            { label: "D", text: "5", isCorrect: false, rationale: "5 is odd." },
          ],
          correctAnswers: ["A", "C"], sourceMapping: "Demo fixture", metadata: { contentHash: "demo2" },
        },
        {
          id: "demo--s1--q003", setId: "demo--s1", sourceFile: "demo-fixture.json", questionNumber: 3, ordinal: 3, type: "mc",
          stem: ["This question has no rationale in the source, to demonstrate the neutral fallback message."],
          options: [
            { label: "A", text: "Option one", isCorrect: true, rationale: "" },
            { label: "B", text: "Option two", isCorrect: false, rationale: "" },
          ],
          correctAnswers: ["A"], metadata: { contentHash: "demo3" },
        },
        {
          id: "demo--s1--q004", setId: "demo--s1", sourceFile: "demo-fixture.json", questionNumber: 4, ordinal: 4, type: "open",
          stem: ["Describe how you would review this project with a study group."],
          answerBlock: [
            { text: "Answer source: demo.", bold: false, bullet: null },
            { text: "Walk through each set and discuss missed items.", bold: false, bullet: 0 },
          ],
          correctAnswers: [], metadata: { contentHash: "demo4" },
        },
      ],
    };
  }

  /* ------------------------------ bank load ------------------------------ */

  function loadBankMirror() {
    // Inject question-bank.js only when needed (file:// mode or missing JSON),
    // so hosted deployments do not download the bank twice.
    return new Promise(function (resolve) {
      if (window.__PEP_SKIP_MIRROR__) { resolve(false); return; } // test hook
      let settled = false;
      const s = document.createElement("script");
      s.src = "question-bank.js";
      s.onload = function () { if (!settled) { settled = true; resolve(true); } };
      s.onerror = function () { if (!settled) { settled = true; resolve(false); } };
      setTimeout(function () { if (!settled) { settled = true; resolve(false); } }, 2500);
      document.head.appendChild(s);
    });
  }

  async function loadBank(force) {
    let bank = null, source = "none", note = "";
    const local = Storage.get(LS_KEYS.bank, null);
    if (local && local.format === "exam-bank" && !force) {
      const v = window.ExamParser.validateBank(local);
      if (v.valid) { S.bank = local; S.bankSource = "local"; S.bankNote = ""; return; }
      note = "Saved local bank failed validation and was ignored.";
    }
    try {
      const res = await fetch("question-bank.json", { cache: "no-cache" });
      if (res.ok) {
        const data = await res.json();
        const v = window.ExamParser.validateBank(data);
        if (v.valid) { bank = data; source = "published"; }
        else note = "question-bank.json exists but is invalid: " + (v.errors[0] || "");
      }
    } catch (e) {
      /* file:// or offline: fall back to the .js mirror below */
    }
    if (!bank && !window.EXAM_BANK) await loadBankMirror();
    if (!bank && window.EXAM_BANK) {
      const v = window.ExamParser.validateBank(window.EXAM_BANK);
      if (v.valid) {
        bank = window.EXAM_BANK; source = "published";
        note = note || "Loaded question-bank.js mirror (fetch was blocked, e.g. file:// mode).";
      }
    }
    S.bank = bank;
    S.bankSource = source;
    S.bankNote = note;
    if (note) toast(note, "warn", 6000);
  }

  /* ------------------------------- theme --------------------------------- */

  function applyTheme() {
    const pref = S.prefs.theme || "system";
    const dark = pref === "dark" || (pref === "system" && window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches);
    document.documentElement.setAttribute("data-theme", dark ? "dark" : "light");
  }

  /* ------------------------------- boot ---------------------------------- */

  function boot() {
    mainHost = document.getElementById("main");
    loadPersisted();
    applyTheme();
    if (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").addEventListener) {
      window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", function () {
        if ((S.prefs.theme || "system") === "system") applyTheme();
      });
    }
    window.addEventListener("hashchange", render);
    document.addEventListener("keydown", function (e) {
      if (S.route.name !== "exam" || !S.exam) return;
      const t = e.target;
      if (t && (/^(INPUT|TEXTAREA|SELECT|BUTTON)$/.test(t.tagName))) return;
      if (e.altKey || e.ctrlKey || e.metaKey) return;
      if (e.key === "ArrowRight") navigateNext(1);
      else if (e.key === "ArrowLeft") navigateNext(-1);
    });
    function navigateNext(d) {
      const qs = setQuestions(S.exam.setId);
      const i = S.exam.index + d;
      if (i < 0 || i >= qs.length) return;
      S.exam.index = i;
      const p = setProgress(S.exam.setId);
      p.last = i; p.updated = Date.now();
      saveProgress();
      render();
    }
    loadBank(false).then(render);
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot);
  else boot();

  window.ExamApp = {
    state: S, storage: Storage, grade: gradeQuestion, setStats: setStats,
    loadBank: loadBank, render: render, navigate: navigate,
    downloadBank: downloadBank, exportProgress: exportProgress,
    confirmDialog: confirmDialog, toast: toast, LS_KEYS: LS_KEYS, demoBank: demoBank,
    setQuestions: setQuestions, answerRec: answerRec,
    importFlow: ImportFlow,
  };
})();
