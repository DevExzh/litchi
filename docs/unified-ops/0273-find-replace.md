# 0273 — Formatting-aware find/replace (plan step 20a, M4 engine 1 of 3)

Status: implemented on `uo2/s20-find` (cut from `feat/unified-ops-api` at
395d0ffa9f); awaiting coordinator review and merge.

Plan: [0248](0248-coordinator-restart-plan.md) §6 Phase E step 20. Design:
§13 P0 "formatting-aware find/replace" and the naming note (OfficeCLI models
find/replace as `find=`/`replace=` props on `set`). Goal: "Verb names" —
`find=`/`replace=` props on `set` first, a dedicated verb only if usage
demands it. Glossary: [API_GLOSSARY.md](../API_GLOSSARY.md), section
"Formatting-aware find/replace (record 0273)".

## 1. What landed

### Surface

- `File::set`, `transaction::Edit::set`, `File::apply`/`batch` `set` items:
  a property named `find` (ASCII case-insensitive, as OfficeCLI's property
  dictionary) turns the call into a text search over the target's subtree.
  No new verb. Keywords:

  | Key | Value | Meaning |
  |---|---|---|
  | `find` | non-empty text (integers are spelled as text) | the literal to search for |
  | `replace` | text, may be empty | the replacement; empty deletes each match |
  | `matchCase` | bool, default `true` | `false`: per-character simple lowercase comparison |
  | `wholeWord` | bool, default `false` | `true`: only matches not adjacent to a letter, digit or `_` |
  | `regex` | bool | `true` (and OfficeCLI's `r"..."` form) is `unsupported_capability` (record 0010 Q3: no approved regex crate) |
  | any other key | run property | formats the replaced text, or the matched text without `replace` |

- The match count: `Applied::matched() -> Option<usize>` (`None` for every
  other command). It is packed beside `changed` (a private `u32`), so
  `ItemResult` stays four words (`docx_batches` size assertion). CLI: the
  result item carries `"matched": N` in `--json` (OfficeCLI's name) and
  `(N matched)` in text. Python: the result dict of `apply`/`batch` has
  `matched`. `File::set` returns only the file (the count is dropped); a
  caller that needs it uses `File::apply`.
- `capabilities::Verb::Replace` (`"replace"`): supported for DOCX, PPTX,
  XLSX and RTF (refused with `RETAINED_SOURCE` on a handle opened with
  `retain_source(false)`), `unsupported_capability` with a follow-up for DOC,
  PPT, XLS, XLSB and the read-only formats; signed/locked states answer as
  every edit verb.
- The schema export (`schema()`, CLI `schema`, Python) gains a Litchi member
  `setKeywords`: the verb's support and the five keywords in the property
  shape (`type`, `constraint`, `default`, `set`, and `writeGap` for `regex`).
  OfficeCLI documents them as cross-cutting `set` keywords outside its
  element schemas (`schemas/help/docx/revision.json`), so no element table
  lists them.

### The engine (`litchi-ops/src/find.rs`, `find/{matcher,plan}.rs`)

- `Spec::admit` admits the keywords (repeated key `invalid_argument`,
  non-text `invalid_value`, empty pattern `invalid_value`, control characters
  `invalid_value`, > 64 KiB `limit_exceeded`, neither `replace` nor a format
  property `missing_property` naming `replace`, as OfficeCLI refuses a bare
  find). Every other key is kept, in order, for the format's run registry.
- `Matcher`: leftmost, non-overlapping, left to right (OfficeCLI's
  `FindHelpers.FindMatchRanges`). Case-sensitive search uses `str::find`.
  Case-insensitive search folds each character by its one-to-one lowercase
  mapping (final sigma as sigma; a character whose lowercase expands, `İ`,
  compares as itself; `ß` never matches `ss`). Whole-word retries one
  character after a rejected candidate.
- `plan(pieces, matcher, replacement, marks, max_bytes)`: a format adapter
  hands the text of each run as a `Piece` with a group; a match never crosses
  a group. For every changed piece the plan gives its new text and, with
  format properties, the byte ranges to format. A match spanning pieces
  collapses into the piece where it starts (that run's formatting); the
  matched characters of the following pieces are removed from them; pieces
  without a match are never touched. A match starting exactly at a run
  boundary belongs to the following run. `scoped: false` pieces filter
  matches (OfficeCLI's run scope keeps only matches fully inside the run) but
  still give whole-word checks their context.

### Formats

| Format | Scope | Groups (a match never crosses) | Write | Format props |
|---|---|---|---|---|
| DOCX | `/`, `/body`: every main-story paragraph, including table cells and block content controls; a table/row/cell/`sdt`: its paragraphs; a paragraph; a hyperlink or run: only its own text | tabs, breaks, symbols, drawings, field characters, note references; entering or leaving a hyperlink, inline `sdt`, `smartTag`, `customXml`, `ins`, `moveTo`, `dir`, `bdo` | one leaf call for every changed paragraph (`Edit::rewrite_text_nodes`): direct-body paragraphs share one splice and layout derivation; nested paragraphs replace their element | the runs the rewrite isolated go through the ordinary DOCX run setters (`format_run_at`) |
| PPTX | `/`: every slide; a slide: its shapes, inside groups too; a shape/placeholder: its body; a paragraph; a run or line break: only its own text | line breaks and fields; each paragraph is a group | `TextTarget::Run` text edits, one native edit for the whole find | `TextTarget::RunSpan` edits on the replaced/matched bytes |
| XLSX | `/`: every worksheet; a sheet, cell, range, row or column | each stored text value | one native workbook edit, `Worksheet::set` per changed cell | refused (`unsupported_property`), as OfficeCLI's Excel find takes only `replace` |
| RTF | `/`, `/body`: every main-story paragraph; a paragraph; a run: only its own text | each paragraph | one native `replace_body_run_texts` | refused (`unsupported_property`): no native run split yet |

Markers that carry no text join the text around them (DOCX bookmarks,
proofing marks, permission and comment ranges, `w:lastRenderedPageBreak`).
Text that is a cached or deleted result is not searched: DOCX complex-field
results (between `begin` and `end`), `w:del`/`w:moveFrom`, `w:fldSimple`,
math and markup-compatibility choices; XLSX formula cells (OfficeCLI rewrites
pure string-literal formulas and warns about the others; Litchi searches
neither).

Refused rather than skipped, so a broad find never silently leaves text
behind: a PPTX table cell in scope whose text matches
(`unsupported_capability`; the leaf edits cells only whole), an XLSX cell
with rich-text runs whose text would change (a plain value would drop the
runs' formatting), and DOCX isolation (format props) inside a container no
element path addresses (`w:ins`, `w:smartTag`, ...): `Refusal::ComplexRun`.

### Leaf primitives (small, separate hunks)

- `litchi-docx` (`document/transaction/formatting/element/text_nodes.rs`,
  new): `TextNode { text, group, run: Option<ElementPath> }`,
  `TextNodeRewrite::new(node, text).isolating(ranges)`, `ParagraphRewrite`,
  `Edit::{text_nodes, rewrite_text_nodes}` on the snapshot edit, forwarded by
  the package edit (`owned.rs`). The scanner is namespace-aware (inherited
  bindings, default-namespace fallback) and bounded (65,536 nodes or runs per
  paragraph). A rewrite keeps each `w:t` start tag (adding
  `xml:space="preserve"` when the new text needs it); isolation copies the
  run's start tag and `w:rPr` to every part; every rewrite is read back
  semantically. Operations are recorded as `Operation::ReplaceElement`, so
  replay, inverse and durable patches reuse the existing element machinery.
- `litchi-pptx` (`opened/text_format.rs`): `TextTarget::RunSpan { shape,
  paragraph, run, start, end }`, a byte span of one `a:r` split out with a
  copy of its markup.
- `litchi-xlsx` (`workbook/model.rs`): `Worksheet::has_rich_text(at)`, from
  the cell's inline-rich flag or a lazily scanned set of rich shared-string
  items.

### Atomicity and exactness

Every find is one transaction operation and one native commit (it never
joins a coalesced run: `coalesce_parsed` flushes and declines it). Zero
matches commit an empty native edit; the published bytes are identical
(tested on both owners for DOCX, PPTX and XLSX, and for RTF). A replacement
equal to the matched text counts the match and changes nothing. Signed and
locked files refuse as every edit does (`File::set` checks them before
admitting the find); the source fence brackets every adapter.

## 2. Deltas vs OfficeCLI

| Id | OfficeCLI | Litchi |
|---|---|---|
| F1 | regular expressions (`regex=true`, `r"..."`, `(?i)` for case) | refused `unsupported_capability`; `matchCase=false` and `wholeWord=true` are Litchi's literal options instead |
| F2 | tabs and breaks are invisible to the match (a match can span `a<w:tab/>b`), `\v`/`\n` in find/replace mean breaks and paragraph splits | tabs, breaks and other run content separate text; control characters in `find`/`replace` are `invalid_value` |
| F3 | a match crossing a hyperlink edge is refused | the edge separates text, so no such match exists |
| F4 | runs emptied by a cross-run replace are removed | they stay as runs with an empty text node (DOCX `<w:t/>`, PPTX `<a:t></a:t>`) |
| F5 | DOCX `/` also searches headers, footers, footnotes, endnotes; PPTX `/` also notes | main story / slides only; those stories are capability gaps of the read adapters |
| F6 | DOCX complex-field results and `fldSimple` results are searched | not searched (cached results) |
| F7 | XLSX: shared strings rewritten in place (a sheet-scoped find changes every sheet sharing the string); per-run replace inside rich strings; pure string-literal formulas rewritten | each changed cell gets its own value; rich-text cells refused when they would change; formulas not searched; cell, range, row and column scopes added |
| F8 | find + `revision.*` records tracked changes | `revision.*` is not a run property: `unsupported_property` (dropped under `--unsupported warn`) |
| F9 | find + paragraph properties (`align`, ...) apply to matched paragraphs | only run properties are admitted; paragraph keys are `unsupported_property` |
| F10 | `text=` with find is reported unsupported and ignored | `invalid_argument` naming `text` (DOCX, PPTX); `unsupported_property` (XLSX, RTF, which take no format properties) |
| F11 | `--json` envelope member `matched` | per result item `matched`; no zero-match warning (the count says it) |

## 3. Tests

- Engine (`find/tests.rs`, 16): leftmost/non-overlapping, case folding
  (ẞ/ß, Σ/ς, İ), whole-word retry, cross-run collapse, run-boundary
  ownership, groups as barriers, several matches per run, empty pieces,
  unchanged pieces, replacement and format-only marks, scope filtering, the
  staging bound, UTF-8 boundaries, admission and refusals.
- DOCX leaf (`text_nodes/tests.rs`, 7, plus 2 doctests): grouping and
  transparent markers; field results, `w:del`, `fldSimple` excluded;
  byte-exact rewrite with entity escaping and `xml:space`; isolation paths
  in the paragraph and in a hyperlink; refusal inside `w:ins`; body and
  nested paragraphs in one call with an exact inverse; malformed requests.
- PPTX leaf: run spans split, format and replace, ordered edits, refusals.
- XLSX leaf: rich shared-string scan.
- Ops integration: `docx_find_replace.rs` (8, both owners), 
  `pptx_find_replace.rs` (5, both owners), `xlsx_find_replace.rs` (4),
  `rtf_find_replace.rs` (2); capability test
  `replace_is_supported_where_a_find_adapter_exists`.
- CLI: `find_and_replace_report_the_match_count`; the option-spelling test
  now expects `--find` alone to be `missing_property` (`replace`).
- Python: `test_find_and_replace_report_the_match_count`.

## 4. Corpus

Re-recorded (`LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test
officecli_corpus`, `python3 tools/officecli_corpus.py`, `--check`: exit 0).
Totals before and after are identical: 6,431 lines, 3,803 passed, 1,105
warned, 891 gap, 168 delta, 464 cascade, 0 bug, 177 resident. The corpus's
only `--find` calls (`word/revisions.sh` sections 7–8) address paragraphs
by the `@paraId` paths `add` prints, so they stay D1 cascades (no path), and
they combine find with `revision.*` (delta F8). Named gap G3 now cites this
record; it counts 0 lines, as before. The capability delta table
(`tools/officecli_capability_delta.py`) changes only its input hashes: the
tool reads element tables, and the keywords are a schema-level member.

## 5. Performance

`cargo run --release -p litchi-ops --example find_replace_bench --features
docx,pptx` (pinned with `taskset -c 29`; 2 warm-up rounds, 9 ABBA rounds of
four samples, medians; build at 683a2d4b06 plus the scope fast path and
prefilter of the follow-up commit). Evidence:
[results/0273-find-replace-bench.json](results/0273-find-replace-bench.json)
(4.6 KB, SHA-256 `8323d0cb8d46a5094eb517ed7993ea9b12aee86f292426b7da5de4bbeb4f62c1`).
Machine load average was 15–20 (six agents share it).

| Format | Containers | Matches | Leaf ms | Unified ms | Ratio | Unified containers/s |
|---|---:|---:|---:|---:|---:|---:|
| DOCX | 10,000 paragraphs | 1 | 14.93 | 16.83 | 1.128 | 594,197 |
| DOCX | 10,000 paragraphs | 1,000 | 19.61 | 21.51 | 1.097 | 465,006 |
| PPTX | 1,000 slides | 1 | 21.41 | 22.29 | 1.041 | 44,862 |
| PPTX | 1,000 slides | 1,000 | 133.54 | 137.99 | 1.033 | 7,247 |

The leaf legs are lower bounds: they use the same leaf primitives but know
the fixture (no cross-run matching, one run per PPTX body). Both legs'
outputs are checked equal (published bytes) outside the clock.

- PPTX meets the 10% single-op budget. The PPTX adapter skips a shape whose
  scene text cannot hold a match (`Matcher::may_match`) before decoding its
  body; without it the one-match case was 1.145.
- DOCX misses the 10% budget by 2.8 points at one match (1.9 ms over 10,000
  paragraphs). Reading direct body paragraphs from the block index instead
  of node values took it from 1.214 to 1.128, and reusing the planner's
  buffers to 1.124; the rest is the general matcher and planner per
  paragraph against the leaf leg's `str::contains`. Follow-up: a leaf-level
  prefilter over the paragraph bytes.
- A find over 10,000 DOCX paragraphs is one leaf call and one splice, so
  1,000 matches cost 5 ms more than one match.

## 6. Gaps and follow-ups

- DOC, PPT, XLS, XLSB: `Verb::Replace` is `unsupported_capability` with the
  follow-up "a find adapter over the format's run-level text edits" (DOC's
  `TextTarget::BodySpan`, XLS/XLSB cell values; PPT replaces shape text
  whole). ODF and iWork have no write adapter yet (M6).
- RTF format properties need a native run split.
- PPTX table cells (a cell run target in `litchi-pptx`), notes, masters and
  layouts; DOCX headers, footers, notes and comments (read-adapter gaps
  first).
- XLSX rich-text cells: a native per-run shared-string rewrite.
- Regular expressions wait on an owner-approved dependency (record 0010 Q3).
- Tracked-change finds (`revision.*`) and paragraph properties on matched
  paragraphs.
- Consolidation with record 0274 (template merge, on trunk at 3cb803e29c):
  its leaf `litchi_docx::document::substitute` also rewrites text across
  runs preserving formatting, in every story (headers, footers, notes,
  comments), and its PPTX merge module rewrites run by run. This record keeps
  its own engine (literal options, run scope, isolation for format
  properties, one splice for body paragraphs); 0274's follow-up F9 and this
  one should move both onto one leaf primitive, which would also give
  find/replace the header/footer/note stories (delta F5).
