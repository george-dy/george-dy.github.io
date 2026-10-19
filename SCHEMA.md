# Question Bank JSON Schema (format `exam-bank`, version 1)

The question bank is a single static JSON file (`question-bank.json`) that the app
fetches on startup. It is produced either by the **in-app Import tool**
(Import → review → accept → download) or by the maintainer build script
(`tools/build-bank.mjs` + `tools/compress-assets.py`). Both use the exact same
parser code (`js/parser-core.js`), so banks are deterministic.

A mirror file `question-bank.js` (`window.EXAM_BANK = {...};`) may be published
alongside it. It is only used when `fetch("question-bank.json")` is blocked
(e.g. opening `index.html` by double-click). Keep both files in sync.

## Top level

| Field         | Type     | Meaning |
|---------------|----------|---------|
| `format`      | string   | Always `"exam-bank"`. |
| `version`     | number   | Schema version (`1`). |
| `bankId`      | string   | Content-derived stable id (`bank-` + 12 hex chars). Changes only when content changes. |
| `title`       | string   | Human title of the whole bank. |
| `generatedAt` | string   | ISO-8601 timestamp of the export. |
| `generator`   | string   | Which tool built the file. |
| `documents`   | array    | One record per source `.docx` (see below). |
| `sets`        | array    | Examination sets (see below). |
| `questions`   | array    | All questions of all sets (see below). |
| `info`        | object   | Optional provenance notes, incl. `parseWarnings` from import time. |

### `documents[]`

| Field | Type | Meaning |
|-------|------|---------|
| `file` | string | Original source filename, e.g. `"Week 1.docx"`. |
| `title` | string | Document title detected in the file (first heading). |
| `description` | string | Front-matter text (scope/instructions), truncated. |
| `importedAt` | string | ISO timestamp. |
| `paragraphCount` | number | Paragraphs read from the file. |
| `setIds` | string[] | Ids of sets detected in this file, in order. |

### `sets[]`

| Field | Type | Meaning |
|-------|------|---------|
| `id` | string | Stable unique id: `<file-slug>--s<index>[-<heading-slug>]`, e.g. `week-1-docx--s2-form-b`. Never reuses array positions. |
| `title` | string | Original set heading when found (e.g. `Form A - 50 Items`), else a generated label. Editable in the import review screen. |
| `sourceFile` | string | Source `.docx` filename. |
| `documentTitle` | string | Title of the containing document. |
| `description` | string | Scope/subtitle lines under the heading, when present. |
| `questionIds` | string[] | **Ordered** ids of the questions in this set. |
| `metadata.setIndex` | number | 1-based index of the set inside its file. |
| `metadata.headingText` | string | Raw heading text (`""` if none). |
| `metadata.detectedBy` | string[] | Signals used: `heading`, `numbering-restart`, `implicit-first`. |
| `metadata.confidence` | string | `"high"` or `"review"` (review = boundary auto-detected; flagged in the import UI). |
| `metadata.autoTitle` | boolean | True when the title was generated, not from the source. |

### `questions[]`

| Field | Type | Meaning |
|-------|------|---------|
| `id` | string | Stable unique id: `<setId>--q<ordinal padded>`. |
| `setId` | string | Owning set id. |
| `sourceFile` | string | Source `.docx` filename. |
| `questionNumber` | number | Number as printed in the source document. |
| `ordinal` | number | 1-based position inside its set (what the UI navigator shows). |
| `type` | string | `"mc"` (single answer), `"sata"` (select-all, exact match), `"tf"` (true/false), `"open"` (open-ended case, self-graded). |
| `title` | string? | Optional case title (e.g. `Card 3 - Scalp Swelling…`). |
| `stem` | string[] | Question text, one array entry per source paragraph. |
| `images` | array? | See below. |
| `options` | array | `{label, text, isCorrect, rationale}` per option (empty for `open`). |
| `correctAnswers` | string[] | Option labels of the key, e.g. `["B","D"]`. Must agree with `isCorrect` flags (validator enforces). |
| `answerBlock` | array? | Only for `open`: faithful faculty-answer lines `{text, bold, bullet}`. |
| `sourceMapping` | string? | Verbatim "Source Mapping:" line from the source. |
| `rationaleNotes` | string[]? | Rationale-section lines that belong to no option (kept, never dropped). |
| `trailingNotes` | string[]? | Any other trailing source lines preserved for review. |
| `issues` | string[]? | Parser flags, e.g. `missing-rationale:CD`, `missing-key`, `image-extraction-failed:x.png`. |
| `metadata.contentHash` | string | Hash of stem+options+key+images; used for duplicate detection. |

`options[]` entries:

| Field | Type | Meaning |
|-------|------|---------|
| `label` | string | Source option label (`A`…`J`). |
| `text` | string | Verbatim option text. |
| `isCorrect` | boolean | Derived from the answer key. |
| `rationale` | string | Verbatim per-option rationale; `""` when the source has none (the UI then shows “No rationale provided in the source material.”). |

`images[]` entries:

| Field | Type | Meaning |
|-------|------|---------|
| `src` | string | `assets/<file>` relative path (published bank) or `data:` URL (in-browser exports). |
| `alt` | string | Alt text: figure caption or Word image description when available. |
| `caption` | string | Figure caption line from the source, when found. |
| `position` | string | `"stem"` (needed to answer; shown before submission) or `"rationale"` (reference asset; shown after submission). |
| `sourceName` | string | Original media filename inside the `.docx`. |
| `contentType`, `widthPx`, `heightPx` | – | Technical metadata. |
| `error` | string? | Set when extraction failed — the UI shows a visible warning instead of silently omitting the image. |

## Validation rules enforced by `ExamParser.validateBank`

1. `sets`/`questions` arrays exist; all ids unique.
2. Every `question.setId` and every `set.questionIds[]` entry resolves.
3. `type` ∈ {mc, sata, tf, open}; `mc` has exactly one correct option; options exist and labels are unique.
4. `correctAnswers` and `isCorrect` flags **must agree**.
5. Images must have a `src`.
6. Warnings (non-blocking): empty stems, missing keys, `sata` with <2 keys, count mismatches.

## Minimal example

See `docs/example-bank.json` for a complete, valid two-set fixture
(multiple choice, SATA, and an open-ended case, including a data-URL image).

## Progress storage (separate from the bank)

Study progress is **not** part of the bank. It lives in browser localStorage:

* `pep.progress.v1` → `{ [setId]: { answers: { [questionId]: { sel, submitted, correct, partial, revealed, self, note, attempts, submittedAt } }, last, updated } }`
* `pep.bank.v1` → a locally imported bank (overrides the published one until you switch back).
* `pep.prefs.v1` → `{ theme }`.

Because answers are keyed by stable set/question ids, re-publishing an updated
bank keeps your progress for every question whose id did not change.
