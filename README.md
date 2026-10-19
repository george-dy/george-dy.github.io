# Practice Examination Platform — your complete instruction manual

A static (no-server) study website for practicing exam questions that live in your
course **Word (.docx) documents**. You upload a small folder of files to GitHub,
turn on **GitHub Pages**, and you get a private-class study website like
`https://YOURNAME.github.io/YOUR-REPO/`. Classmates can practice multiple-choice,
select-all-that-apply, and open-ended case questions, see rationales after each
answer, and their progress is saved in their own browser.

No programming, no server, no database, no accounts, no paid services.

---

## 0. First-time setup checklist

**One-time setup (do once):**

- [ ] 1. Download/copy the project folder (the files listed in section 2) to your computer.
- [ ] 2. Open `index.html` by double-clicking it to preview the app (section 3).
- [ ] 3. Create a free GitHub account (section 4).
- [ ] 4. Create a new repository (section 5).
- [ ] 5. Upload the files listed in section 6.
- [ ] 6. Enable GitHub Pages (section 7).
- [ ] 7. Open your published URL and run the 2-minute test in section 9.

**Adding new questions (do every time you receive new .docx files):**

- [ ] 1. Open your published site (or the local `index.html`).
- [ ] 2. Go to **Import**, choose the new `.docx` files (section 10).
- [ ] 3. Review the detected sets, fix anything flagged.
- [ ] 4. Click **Accept import**, then **Download question-bank.json** and **question-bank.js**.
- [ ] 5. Back up the old bank files on your computer, then upload the new ones to GitHub (section 11).
- [ ] 6. Wait ~1 minute, refresh the site, confirm the new sets appear (section 9).

**Routine maintenance (every few weeks / before big edits):**

- [ ] Download a backup of the question bank (Settings → *Download question-bank.json*).
- [ ] Download a backup of your personal progress (Settings → *Export progress backup*).
- [ ] Keep your original `.docx` files in a safe folder — they are the true source.

> **The most important concept in this project (read once, saves pain later):**
> *Importing* a Word document happens **inside your browser only**. Nothing is
> uploaded anywhere and your classmates see **no change** until *you publish*:
> download the updated `question-bank.json` + `question-bank.js` and replace the
> old copies in your GitHub repository. Workflow diagram:
>
> ```
>  your .docx files
>        │  (1) Import screen: browser reads & parses them locally
>        ▼
>  review & correct detected sets      ← happens on YOUR computer
>        │  (2) "Accept import" → bank active in YOUR browser
>        │  (3) Download question-bank.json + question-bank.js
>        ▼
>  files on your computer
>        │  (4) Upload both files to GitHub (replacing old ones, after backup)
>        ▼
>  GitHub Pages rebuilds (~1 min)
>        │  (5) Everyone refreshes → new sets appear for the whole class
>        ▼
>  the published website
> ```
>
> Your personal study progress is stored **only in your browser** and is never
> part of the published question bank.

---

## 1. What the application does

- **Library view** — lists every examination set (e.g. *Week 1 — Form A*), grouped
  by source document, with search, status filters (Not started / In progress /
  Completed), question counts, how many you answered, how many were correct,
  a completion bar, and a **Start / Continue / Review** button per set.
- **Exam view** — one question at a time with a progress bar, question navigator
  grid (jump anywhere; colour + icon + label show unanswered / correct /
  incorrect / reviewed), Previous/Next, and **Submit answer**.
  After submitting, the result and the **rationale for every option** appear
  (never before). Select-all-that-apply is graded all-or-nothing (partial credit
  is shown separately as a percentage note, never counted as correct).
  Open-ended case questions (from documents without A/B/C options) hide the
  faculty answer behind a **Reveal faculty answer** button and let you grade
  yourself honestly ("I knew this" / "Needs review").
- **Import view** — converts your `.docx` files in the browser: detects set
  boundaries (headings like *Form A - 50 Items*, *Examination 2 of 5: …*,
  *Part II - …*, or numbering restarts), questions, options, answer keys,
  per-option rationales, source mappings, and embedded images; flags anything
  uncertain for your review; then exports a publishable question bank.
- **Settings view** — bank information, export/import of the bank and of your
  progress backup, theme choice, and the (deliberately cautious) reset buttons.

Everything runs in the browser. There is no backend, no login, and no internet
dependency after the page loads (the vendored library `vendor/jszip.min.js` ships
inside the project, so even the import tool works offline).

## 2. What files you will receive and what each one is for

You need these for the website (upload **all** of them, keeping the folders):

| File / folder | Purpose |
|---|---|
| `index.html` | The whole application (interface + styles). Open this to study. |
| `question-bank.json` | **Your questions.** The published examination data. |
| `question-bank.js` | Same data in a form the browser can read when you open `index.html` by double-clicking. Always replace it together with the `.json`. |
| `assets/` (folder) | The images extracted from your Word documents, referenced by the question bank. |
| `js/` (folder: `app.js`, `parser-core.js`, `docx-ir.js`) | The application code and the document-parsing engine. |
| `vendor/jszip.min.js` | A small, well-known library (kept inside the project, not loaded from the internet) that lets the browser open `.docx` files. |
| `.nojekyll` | Tells GitHub Pages not to run its optional Jekyll processing. Keep it. |
| `README.md` | This manual. |
| `SCHEMA.md` | Technical description of the JSON format (for the curious / future maintainers). |
| `docs/example-bank.json` | A tiny valid example bank used by the documentation/tests. |

Optional, **not** needed by the website (developer/maintenance tooling — you may
upload or ignore; it changes nothing):

| File / folder | Purpose |
|---|---|
| `tools/build-bank.mjs`, `tools/compress-assets.py`, `tools/smoke.mjs` | Rebuild `question-bank.json` from the original `.docx` files outside the browser (used to create the first published bank). |
| `tests/` | Automated test suite (`npm test`). |
| `package.json` | Only lists the developer test tools. |
| `assets-raw/`, `question-bank.staging.json` | Intermediate build outputs; delete or ignore. |

## 3. How to preview the application on your computer

1. Put the whole project folder somewhere permanent (e.g. `Documents/exam-practice`).
2. Double-click `index.html`. Your browser opens the app.
   - Because the published `question-bank.json` is next to it, your questions load
     even in double-click mode (that is what `question-bank.js` is for).
3. Click any set → **Start** → answer a question → submit → see the rationales.
4. Optional (closer to the real website): serve the folder like a mini web server.
   - **Windows:** open the Start menu, type `cmd`, press Enter. In the black window
     type `cd /d ` then drag the project folder into the window and press Enter,
     then type `python -m http.server 8000` and press Enter (if Python is
     installed). Open `http://localhost:8000` in your browser.
   - **Mac:** open *Terminal* (Spotlight → type Terminal), type `cd `, drag the
     folder in, Enter, then `python3 -m http.server 8000`, Enter.
     Open `http://localhost:8000`.
   - To stop the server: click the black window and press `Ctrl`+`C`.
   - If you do not have Python, skip this — double-clicking works fine.

## 4. How to create a GitHub account (if you do not have one)

1. Go to <https://github.com> and click **Sign up**.
2. Enter an email you check regularly, create a password, choose a username
   (this username becomes part of your website address).
3. Solve the puzzle/verification and confirm your email via the link GitHub sends.
4. You can stay on the free plan forever; this project needs nothing paid.

## 5. How to create a GitHub repository

A *repository* ("repo") is just a folder of files stored on GitHub.

1. Sign in at <https://github.com>.
2. In the top-right corner click the **+** button, then choose **New repository**.
3. In **Repository name** type a short name without spaces, e.g. `practice-exams`.
   (This name becomes part of your website address.)
4. Choose **Public** (GitHub Pages for free accounts publishes public repos) —
   note that classmates with the link can also see the question bank file, which
   contains the answers; that is inherent to a static study site.
5. Check **Add a README file** — *optional*; if GitHub creates one you can
   overwrite it later with this README.
6. Click the green **Create repository** button.

## 6. Exactly which files and folders to upload

From section 2, upload **these, keeping the folder structure**:

```
index.html
question-bank.json
question-bank.js
.nojekyll
README.md                (optional but recommended)
SCHEMA.md                (optional)
assets/                  (the whole folder, all files inside)
js/                      (the whole folder: app.js, parser-core.js, docx-ir.js)
vendor/                  (the whole folder: jszip.min.js)
docs/                    (optional)
```

Do **not** upload: `node_modules/`, `assets-raw/`, `tests/`, `tools/`,
`package.json`, `question-bank.staging.json` (they are not needed; harmless if
you do, but they clutter the repo).

**How to upload with the website interface (no command line):**

1. In your new repository, click the **uploading an existing file** link
   (or the **Add file** button → **Upload files**).
2. Drag the *files* `index.html`, `question-bank.json`, `question-bank.js`,
   `.nojekyll`, `README.md` into the box.
   - Hidden files: `.nojekyll` starts with a dot. On Windows it may be invisible;
     if you cannot see/drag it, you can create it later with the GitHub editor:
     **Add file → Create new file**, name it exactly `.nojekyll`, leave it empty,
     commit. (The site works without it; it only avoids rare Jekyll quirks.)
3. Scroll down and click the green **Commit changes** button.
4. Now the folders: GitHub uploads folders if you drag them, but the most
   reliable way is one folder at a time:
   - Click **Add file → Upload files** again, then drag the *contents* of `assets/`?
     No — instead: first create the folder by uploading one file into it:
     in the upload screen you can type a path in the file-name box.
   - **Simplest reliable method:** drag the whole `assets` folder from your file
     manager into the upload box — modern browsers keep the folder structure.
     Repeat for `js` and `vendor`.
   - After each drop, verify the preview list shows paths like `assets/xxx.jpg`,
     `js/app.js`, `vendor/jszip.min.js`, then click **Commit changes**.
5. Verify structure: on the repository's main page you should see the entries
   `assets/`, `js/`, `vendor/`, `docs/` (folders) and the single files next to them.

## 7. How to enable GitHub Pages

1. In your repository, click the **Settings** tab (top row of the repo page, far right).
2. In the left sidebar click **Pages**.
3. Under **Build and deployment**, set **Source** to **Deploy from a branch**.
4. Under **Branch**, choose `main` (or `master`, whichever exists) and folder `/ (root)`.
   Click **Save**.
5. Wait 30–120 seconds. Refresh the Pages settings page: a green box appears with
   "Your site is live at …".

## 8. How to find your published website URL

- On the same **Settings → Pages** screen, the banner shows the URL.
- The pattern is always: `https://YOUR-USERNAME.github.io/REPO-NAME/`
  e.g. `https://maria-smith.github.io/practice-exams/`.
- Because the app uses only *relative* links, it works both at a project URL
  like the above and at a custom domain root.

## 9. How to test the website after publishing

Run this 2-minute checklist on the published URL (and repeat after every update):

1. The **Library** lists your sets, grouped per document, with counts.
2. Open one set → the first question shows stem + options; **no rationales visible**.
3. Select an option → it highlights. Press **Submit answer** → result banner and
   per-option rationales appear; the navigator square turns green/red.
4. Press **Next**, answer nothing, press **Previous** → your earlier selection is
   still there; the submitted question still shows its feedback.
5. Reload the page (F5 / Cmd+R) → **Continue** resumes at your last question and
   your statistics are unchanged.
6. Open a second set → its progress is independent (first set untouched).
7. Library → **Reset** on one set → confirm → only that set returns to 0%.
8. If your bank has images: an image displays; clicking it opens the enlarge view.
9. **Import** screen opens and shows the drop zone (you do not need to import here).
10. No red error text anywhere; browser console (F12 → Console) has no errors.

## 10. How to add new Word examination documents

1. Open the site (published URL or local `index.html`) → top menu **Import**.
2. Drag `.docx` files onto the drop zone, or click **Choose files…** (multiple allowed).
   Only `.docx` works (not old `.doc`, not PDF).
3. The app reads each file *on your computer* and shows what it found:
   sets, question counts, detected types, and an issue summary
   (missing answer keys, missing rationales, unresolved images, uncertain set
   boundaries, suspected duplicates).
4. Expand each set panel:
   - Rename the set title if the auto-title is ugly.
   - Panels marked **"boundary needs review"** were created from a numbering
     restart rather than a heading — check the question ranges make sense.
   - Flagged questions can be corrected inline: change the type, or tick the
     correct option labels.
   - Nothing is ever invented: a missing rationale stays missing and the exam
     screen will say "No rationale provided in the source material."
5. If a set already exists in your bank (re-import of the same file), choose per
   set: **Skip**, **Replace**, or **Add as extra copy**.
6. Click **Accept import into this browser**.
   - If error-level issues exist you must first tick the acknowledgement box —
     the app refuses to silently publish broken structure.
7. You are now on step 3, **Publish**: download `question-bank.json` and
   `question-bank.js` (both!) and continue with section 11.

The importer understands the layouts used by your course documents:
`Form A - 50 Items`, `Examination 2 of 5: …`, `Examination 3 of 3 — …`,
`Part II - …` headings; `Question 12: Multiple Choice` / `: Select All That Apply`
/ `: SATA` / `. Card 3 - …` question lines; options `A)`…`F)`; blank-line spoilers;
`Correct Answer: B, D, E`; `Clinical Rationale:` with `• A): …` or
`A) <option text>: <rationale>` styles; `Source Mapping:` lines; `Figure N:`
captions; embedded photos/charts (attached to the right question, stem images
shown before submission, reference figures shown with the rationale); and
open-ended card/picture/scenario cases with faculty answer blocks.

## 11. How to publish updates after adding questions

1. **Back up first:** in your repository, open the old `question-bank.json`,
   copy its content into a text file on your computer named
   `question-bank.BACKUP-YYYY-MM-DD.json` (or simply download it: view the file →
   **Download** button). Same for `question-bank.js` if you like.
2. In the repository file list click `question-bank.json`.
3. Click the **pencil icon (Edit this file)** — or the **trash icon** to delete and
   re-upload. With the editor: select all old content (Ctrl/Cmd+A), delete it,
   then paste nothing — GitHub's editor cannot upload into an open editor, so the
   easier path is: **trash icon → Commit changes**, then **Add file → Upload files**
   and drop your new `question-bank.json`. Repeat for `question-bank.js`.
   (Deleting + re-uploading is safe *because you made a backup in step 1*.)
4. Commit with a message like "Add Week 4 sets".
5. Wait ~1 minute (GitHub Pages rebuild), then refresh your published site and
   check the new sets appear (section 9).
6. Tell classmates to refresh; their saved progress stays (answers are attached to
   stable question ids, so unchanged questions keep their history).

## 12. How to back up and restore your question bank

- **Bank backup:** Settings → *Download question-bank.json* (and `.js`). Store the
  files somewhere safe (cloud drive, USB stick). These files ARE the backup —
  they contain every question, key, rationale and image reference.
- **Bank restore:** Settings → *Import bank from JSON file…* → choose a backup →
  confirm. This affects only your browser; publish it (section 11) if you want it online.
- **Progress backup:** Settings → *Export progress backup* downloads
  `exam-progress-backup.json` (your answers/stats for this browser).
- **Progress restore:** Settings → *Restore progress from backup…* → choose that file.
- **Word originals:** keep the `.docx` files forever in your own folder; they are
  the ultimate source of truth and can always be re-imported.

## 13. How to troubleshoot common problems

| Symptom | What to do |
|---|---|
| Blank page or "No question bank found" | You likely opened `index.html` by double-click **and** `question-bank.js` is missing/older. Put both bank files next to `index.html`, or serve the folder (section 3), or use the published URL. |
| Library empty on the published site | `question-bank.json` was not uploaded, or uploaded into a wrong folder (it must be at the repository root, next to `index.html`). Check the file list. |
| 404 on images | The `assets/` folder was not uploaded, or was uploaded flat (files must live under `assets/`). Re-upload keeping the folder. |
| Site shows an old version | GitHub Pages caches: hard-refresh (Ctrl+Shift+R / Cmd+Shift+R). Also confirm the commit actually replaced the file (file list shows new "updated" time). |
| Import says "not a .docx file" | Old Word `.doc`, or a PDF/Google-Docs export. In Word: *File → Save As → Word Document (.docx)*. In Google Docs: *File → Download → Microsoft Word (.docx)*. |
| Import finds 0 questions | The document uses an unrecognized layout. Open the file and compare with section 10's supported patterns; worst case, retype headings as `Question 1: Multiple Choice` etc. Nothing is lost: the import screen lists exactly what it could not parse. |
| A set is split/merged wrongly | In the import review, panels flagged *boundary needs review* can be renamed; for structural fixes adjust the source document headings (e.g. add `Form B - 50 Items` lines) and re-import. |
| My progress disappeared | Progress lives per browser + device. A different browser, private window, or cleared site data starts fresh. Restore from a progress backup if you made one. |
| "browser storage is full" warning | Your imported bank with many large images exceeds the browser quota. It still works for the session; publish the downloaded bank and, on other machines, load the published one instead. |
| Nothing happens when clicking | Very old browser. Use a current Chrome/Edge/Firefox/Safari. |

## 14. How to reset progress without accidentally deleting your questions

- **One set:** Library card → **Reset** → confirmation dialog explains exactly what
  will be cleared → confirm. Only that set's answers/position are removed.
- **Everything:** Settings → **Reset ALL progress** → you must type `RESET`.
- Resets **never** touch the question bank: your questions, keys, rationales and
  images stay, in the browser and on GitHub.
- There is no silent reset: every destructive action asks first, and "Reset all"
  additionally requires typing a word.

## 15. How to keep your project organized as you add weeks and sets

- Name source documents consistently, e.g. `Week 5.docx`, `Week 6.docx`; the
  library groups sets under that name automatically.
- Inside a document, always give each examination set a heading line
  (`Form A - 50 Items`, `Examination 1 of 3 — Title`, `Part II - …`) — the
  importer then detects boundaries with high confidence.
- Keep one published bank: always replace (never duplicate) `question-bank.json`
  at the repository root; keep dated backups in your own folder, not in the repo.
- Before importing a big new document, download a bank backup (section 12).
- Once a month, run the section 9 checklist on the published site.

---

## Appendix A — JSON schema (short version)

Full detail in `SCHEMA.md`. A question record looks like:

```json
{
  "id": "week-1-docx--s1-form-a--q001",
  "setId": "week-1-docx--s1-form-a",
  "sourceFile": "Week 1.docx",
  "questionNumber": 1,
  "ordinal": 1,
  "type": "mc",
  "stem": ["A 26-month-old insists on putting on shoes independently…"],
  "images": [{ "src": "assets/week-1-img-001.jpg", "alt": "Figure 1…", "position": "stem" }],
  "options": [
    { "label": "A", "text": "Will, via autonomy versus shame and doubt", "isCorrect": true,  "rationale": "The toddler stage…" },
    { "label": "B", "text": "Competence, via industry versus inferiority", "isCorrect": false, "rationale": "School-age…" }
  ],
  "correctAnswers": ["A"],
  "sourceMapping": "Slide 3; Week 1 transcript",
  "metadata": { "contentHash": "0b9328df18db139a" }
}
```

## Appendix B — rebuilding the bank from scratch (maintainers, optional)

Requires Node.js and Python with Pillow, only if you want to regenerate
`question-bank.json` from the original `.docx` files outside the browser:

```
npm install                 # installs jszip, @xmldom/xmldom, jsdom (dev only)
node tools/build-bank.mjs   # parses the .docx uploads -> question-bank.staging.json + assets-raw/
python3 tools/compress-assets.py   # compresses images -> assets/, writes question-bank.json/.js
npm test                    # 48 automated tests (parser + UI)
```

The browser Import tool uses the identical parser, so banks built either way are
equivalent; the in-app route needs no installation at all.

## Appendix C — honest limitations

- The question bank is a public static file: anyone with the URL can read the
  answers by opening `question-bank.json`. Fine for studying, **not** for secure
  exam administration.
- Progress is per browser/device (localStorage). No sync, by design.
- Very large imports (hundreds of MB of images) may exceed browser storage;
  the app warns and still lets you publish via download.
- Word `.doc` (pre-2007) and PDF are not supported; convert to `.docx` first.
- Set boundaries that the source document makes genuinely ambiguous are flagged
  for human review instead of being guessed silently.
