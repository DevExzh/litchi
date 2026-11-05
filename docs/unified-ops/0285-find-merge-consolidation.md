# 0285 — Find/replace fixes and consolidation with template merge

Status: implemented on `uo2/find2` (cut from `feat/unified-ops-api` at
9fe45629ab); awaiting coordinator review and merge.

Inputs: [0273](0273-find-replace.md) (find/replace; its consolidation note
and gaps), [0274](0274-template-merge.md) (template merge; follow-ups F2, F3,
F9), [0279](0279-field-refresh.md) (its finding: a DOCX find over `/body`
rewrites a table of contents whose field spans paragraphs) and
[0281](0281-per-format-write-performance.md) (DOCX find at 1.07–1.09×).
Glossary: [API_GLOSSARY.md](../API_GLOSSARY.md), section "Find/replace and
template merge on one primitive (record 0285)".

## 1. The field bug and the decision

0273's DOCX scanner (`text_nodes`) skipped the text between a complex
field's `begin` and `end` markers, but counted the markers paragraph by
paragraph. A table of contents begins in its first paragraph and ends in
its last, so the first paragraph's entries were skipped and every later
entry was searched: one find rewrote some entries of the same field and not
others. 0274's scanner (`substitute`) had the same per-paragraph field state
and treated result text as text.

OfficeCLI decides what a find reads. `WordHandler.Helpers.FindReplace.cs`
(`BuildRunTexts`) walks `para.Descendants<Run>()` and reads each run's
`Text` elements: field results (complex results, `w:fldSimple` runs) are
text, and instructions (`w:instrText`, a `FieldCode`) are not. Word's own
find also searches displayed field results. So, in the one scanner both
engines now share:

- a field **result** is ordinary text, in every paragraph it spans;
- text inside a field **instruction** (the cached result of a field nested
  in an instruction, `{ IF { MERGEFIELD a } … }`) is fixed content, never
  text, so no find or merge can change what an instruction says;
- field state is followed across paragraphs: the scanner keeps one stack of
  open fields for the story (`Instruction` until `separate`, then
  `Result`), and a scan that starts inside the document (a scoped find)
  starts with the state in effect there;
- markers and instructions are never rewritten; field edges, like hyperlink
  edges, separate text, so no match crosses into or out of a field.

Regression tests: a table of contents whose result spans three paragraphs
(found consistently in every entry, markers and instruction byte-identical,
`TOC` never matched in the instruction, a paragraph-scoped find inside the
result), a simple field (result searched, `w:instr` untouched, the field
edge separating text), complex fields whose result spans runs (cross-run
match collapsing into the result's first run, instruction `PAGE` never
matched, no match across a marker), and an instruction that continues into
the next paragraph (its nested result fixed, in a whole and in a scoped
story).

The refresh test that changed a heading with a find over `/body`
(`tests/refresh.rs`) now addresses the heading paragraph (`/body/p[2]`): a
body find also rewrites the table's entry, which is a field result, so the
table would already read current.

## 2. One leaf primitive

`litchi_docx::document::substitute` is the one DOCX primitive behind both
find/replace and template merge. 0273's `Edit::{text_nodes,
rewrite_text_nodes}` and `TextNode`, `TextNodeRewrite`, `ParagraphRewrite`
are removed and banned (`tools/api_glossary.json`, both DOCX profiles).

| Piece | Behaviour |
|---|---|
| Scanner (`substitute/scan.rs`) | One pass over a story part, or over a fragment of the main document with the namespace bindings and field state in effect where it starts. Lists each paragraph's text (`w:t` text, `\t`/`\n` for tabs and breaks, `OBJECT` for other run content, field markers and instruction text), its fields (0274's `Field`, now also a field that spans paragraphs, listed where it begins and never replaceable), its `Segment`s (one per `w:t`, with 0273's grouping: tabs, breaks, field characters, drawings and container edges separate text; bookmarks, proofing marks, comment ranges and rendered page breaks do not), its runs, and, in a part, the root child it belongs to (`Paragraph::item`: the comment, footnote or endnote). Plain `w:t` content is taken as is; only content with a reference, markup or a carriage return goes through the full decoder. |
| Requests | `Replacement::new(paragraph, range, text)` (0274) replaces a range: the new text takes the first character's run and formatting; `Replacement::keep(paragraph, range)` keeps the text; `.formatted()` splits the new (or kept) text into runs of their own, each a copy of the run it came from. A range equal to a replaceable field's range replaces the field (merge). |
| Writer (`substitute/write.rs`) | Rewrites only the `w:t` elements the requests change (keeping each start tag; `xml:space="preserve"` added when needed; tabs and line feeds become `w:tab`/`w:br`); splits a run only where a formatted request isolates text; applies the caller's run-property closure (`substitute_formatted`) to each split-out run through the ordinary `properties::{parse_run, splice_run}`, with the run's namespace scope, in any story. |
| Commit (`Edit::substitute_formatted`) | One `Operation::ReplaceElement` per changed body block, so patches replay and invert exactly. Direct-body paragraph changes that are a small part of the document (spliced bytes under half of it) derive the layout incrementally (`with_spliced_paragraphs_under`, 0273's path); anything else rescans once (0274's path). Semantic readback of every changed paragraph (its own fragment, or one document rescan when more than 32 and a quarter of the story's paragraphs changed) and exact readback of every changed block. Part stories (`substitute_part_formatted`) are staged as exact part replacements and read back whole. |
| Scoped reads | `Edit::story_within(&[ElementPath])`: a paragraph lists itself, a table, row, cell or block content control every paragraph inside it (text boxes included), a run, hyperlink or inline control its paragraph with only its own text in scope (`Segment::is_scoped`). Scopes inside direct-body paragraphs with no field character before them are scanned paragraph by paragraph; anything else scans from the document start to the end of the last scoped block, so the field state is right. |
| Story parts | `Edit<Package>::story_parts` lists `StoryPart`s (`name`, `kind`: `StoryKind::{Header, Footer, Footnotes, Endnotes, Comments}`); `source_backed::Package::story_parts` reads only the main part's relationships. |

The PPTX side shares one run rewrite: `find::rewrite_ranges` (the planner's
collapse of a range into the run where it starts) plans both a find's
matches and a merge's placeholders, and PPTX find and merge walk the same
text bodies (`find_scopes` / `presentation_scopes`).

What each caller keeps:

- **Merge**: placeholders spanning runs, merge fields, tab and line-feed
  values, every story, exact no-ops and reversible patches are unchanged
  (all 0274 tests pass unchanged). Differences: a placeholder inside a field
  instruction is no longer replaced (instruction text is fixed); a complex
  `MERGEFIELD` spanning paragraphs is now listed and counted in
  `skipped_fields` (0274 claimed this; the old scanner dropped the field);
  values starting or ending with non-ASCII white space get
  `xml:space="preserve"` too. Values are now checked once per key, when it is
  first used, instead of once per placeholder.
- **Find**: grouping, run scope, whole-word context, the cross-run collapse,
  format-only finds and byte-exactness are unchanged (all 0273 tests pass
  unchanged). Format properties are applied by the leaf instead of one
  `format_run_at` per isolated run (which rescanned the document per run),
  so formatting now also works inside tracked insertions, smart tags,
  simple-field results and every story: 0273's `Refusal::ComplexRun` for
  those is gone. Format values are coerced once before anything is read, so
  an invalid value fails even when nothing matches.

## 3. Coverage

| Format | Scope | Before (0273) | Now |
|---|---|---|---|
| DOCX | `/` | main story | main story, every header, footer, footnote and endnote (OfficeCLI's root fan-out); comments are not swept, as OfficeCLI's root find excludes them |
| DOCX | `/body` | main story | main story (unchanged; OfficeCLI also sweeps headers here, delta F5b) |
| DOCX | `/comments`, `/comments/comment[N]` | refused | the comments, or one comment (OfficeCLI's `/comment[N]` paths) |
| DOCX | tables, rows, cells, controls | their paragraphs | their paragraphs and the text boxes in them (OfficeCLI's `Descendants<Paragraph>`) |
| DOCX | field results | not searched (and searched inconsistently across paragraphs) | searched in every paragraph; instructions never |
| PPTX | `/` | slides | slides, their tables' cells, and their notes pages (OfficeCLI's root fan-out) |
| PPTX | `/slide[N]` | shapes; a matching table cell refused | shapes and table cells |
| PPTX | table, row, cell | refused when matching | searched run by run |
| PPTX | `/slide[N]/notes` | refused | the notes page's shapes, run by run |
| XLSX | rich-text cells | refused when changing | still refused (gap below) |

PPTX leaf additions: `TextTarget::{CellRun, CellRunSpan}` edit one run (or a
byte span of one) of a table cell's text body as `Run` and `RunSpan` do for
a shape; `table::Cell::text_body` reads the cell's body with the positions
they address; `opened::Edit::set_notes_text_properties` applies the same
run edits to a notes page's shapes. A source-backed presentation is
materialized for a find or merge that changes a notes page, as a notes text
write already is.

PPTX merge walks the same bodies, so it now fills table cells and speaker
notes run by run (record 0274 follow-ups F2 and F3 closed).

A find over the DOCX root of a source-backed document reaches stories that
publish only through the owning package, so it promotes the handle (as
merge does) — but only when the main part relates a header, footer or
notes part; otherwise it keeps the source editor.

## 4. Deltas vs OfficeCLI (updates to 0273 §2)

| Id | OfficeCLI | Litchi now |
|---|---|---|
| F5 | DOCX `/` also searches headers, footers, footnotes, endnotes; PPTX `/` also notes | **closed**: the same; comments through their own paths, as OfficeCLI |
| F5b | DOCX `/body` is the same sweep as `/` | `/body` is the main story only (Litchi's `/body` node is the body) |
| F6 | DOCX complex-field and `fldSimple` results are searched | **closed**: the same; instructions are not searched (OfficeCLI's `instrText` is not `Text` either) |
| F12 | a match may cross a field edge into or out of a result (runs are concatenated) | field edges separate text, as tabs and hyperlink edges do (F2, F3) |
| F13 | PPTX table cells are searched (`table[M]/tr[R]/tc[C]` paths) | the same |

The other 0273 deltas stand. The OfficeCLI corpus has no `--find` call
that reaches these paths (its only finds address `@paraId` paths and stay
D1 cascades); re-recorded with no outcome change (§7).

## 5. Tests

- Leaf `litchi-docx` (`substitute/tests.rs`, 22): the 0274 cases, plus field
  results across paragraphs (whole and scoped), a spanning result rewritten
  with markers byte-identical and an exact inverse, instruction text fixed
  across paragraphs, segment groups (markers, tabs, hyperlinks, simple
  fields, deletions, insertions), scoped stories (hyperlink scope, table and
  repeated paragraph scopes, a missing path), formatted and kept
  replacements (split runs, formatting in a hyperlink and a tracked
  insertion), nested and body paragraphs with an exact inverse, part items
  and part formatting, malformed requests. Doctests for `story_within` and
  `substitute_formatted`.
- Leaf `litchi-pptx` (`text_format_tests.rs`): cell runs and spans (text,
  split and format, refusals), notes-page run edits (span, inverse, a slide
  without notes refused).
- Ops: `tests/docx_find_stories.rs` (7, both owners, save and reopen): the
  table of contents, simple fields, complex fields spanning runs, root
  stories, formatting in every story, comments by path, text boxes and
  paragraph/cell scopes; `tests/pptx_find_replace.rs` gains table cells and
  speaker notes (both owners); `tests/merge.rs` gains PPTX table cells and
  notes (both owners). All 0273/0274/0279 integration tests pass.

## 6. Performance

Release build, `taskset -c 21`, each harness's ABBA order, medians; machine
shared with other agents (load average noted per run). Before = trunk at
9fe45629ab, after = this branch. Evidence: [results/0285/](results/0285/)
(every sample; binary SHA-256s in `binaries.sha256`). The before and after binaries ran
back to back in the order before, after, after, before; load average 12–17
(`load.txt`). Ranges are the two runs of each build.

Find (`find_replace_bench`; leaf leg = the best hand-written use of the
same leaf API, which changes with this record: 0273's `text_nodes` before,
`story`/`story_within` + `substitute` after):

| Scenario | Unified before | Unified after | Ratio before | Ratio after |
|---|---:|---:|---:|---:|
| DOCX 10,000 paragraphs, 1 match, `/` | 16.74–17.13 ms | 12.62–13.27 ms | 1.063–1.073 | 0.981–1.033 |
| DOCX 10,000 paragraphs, 1,000 matches, `/` | 21.90–22.23 ms | 14.97–15.57 ms | 1.049–1.061 | 0.990–1.030 |
| DOCX one paragraph of 10,000 (`/body/p[5001]`, new case) | — | 3.79–3.86 ms | — | 1.014–1.015 |
| PPTX 1,000 slides, 1 match | 25.17–26.67 ms | 22.86–23.13 ms | 1.047–1.051 | 1.066–1.074 |
| PPTX 1,000 slides, 1,000 matches | 145.5–147.5 ms | 140.0–143.5 ms | 1.018–1.026 | 1.033–1.034 |

Merge (`merge_bench 1000 11`; leaf leg = `story` + a plain `{{key}}` scan +
`substitute`):

| Scenario | Unified before | Unified after | Ratio before | Ratio after |
|---|---:|---:|---:|---:|
| one record, 2,000 paragraphs, 10,000 placeholders | 9.59–9.84 ms | 9.07–9.09 ms | 1.045–1.052 | 1.030–1.031 |
| 1,000 records, source-backed template | 38.62–38.97 ms | 39.28–39.55 ms | 1.033–1.035 | 1.037–1.044 |
| 1,000 records, eager template | 38.40–38.41 ms | 39.18–39.20 ms | 1.031–1.034 | 1.035–1.040 |

- DOCX find is 23–31% faster in absolute time and within its budget; its
  ratio is no longer meaningful against 0281's 1.07–1.09 because the leaf
  leg now uses the faster primitive too (one story scan instead of one
  scanner per paragraph; plain `w:t` content taken without the decoder).
- Merge stays within 1.05 on every scenario; the 10,000-placeholder case got
  faster (its leaf part spends less in the scanner; values are checked once
  per key). The 1,000-record cases are dominated by serialization and moved
  by under 2%.
- What the first consolidated build cost, and what fixed it (callgrind,
  `placeholders-10k`): the per-paragraph readback and a block lookup per
  change made merge 2× slower (19.8 ms); one document rescan when many
  paragraphs change, a hashed block lookup, no cloned splices, the plain-text
  fast path and a full rescan instead of proving many incremental splices
  brought it to 8.8 ms.
- PPTX find: absolute time is lower in this run, but the one-match ratio
  rose by about two points: the presentation now also reads every slide's
  notes relationship (1,000 slides). Within the 10% budget; follow-up: skip
  the notes lookup when the package has no notes slides.
- The paragraph-scoped DOCX find spends almost all its time in
  `Package::edit` (a full layout scan of the document, both legs); the find
  reads one paragraph.

## 7. Gates

All with `CARGO_TARGET_DIR=<worktree>/target` and the build environment of
the rules; exit 0 unless noted.

- `cargo fmt --all --check`
- `cargo clippy -p litchi-docx -p litchi-pptx -p litchi-ops -p litchi -p
  litchi-py --all-targets --all-features -- -D warnings`
- `cargo clippy -p litchi-ops --no-default-features [--features docx | pptx
  | xlsx | rtf | docx,pptx,xlsx] --all-targets -- -D warnings` (the first
  runs of `docx`, `pptx`, `xlsx`, `rtf` failed on ungated plan items and
  were fixed in 8adeb1c697, then passed)
- `cargo test -p litchi-ops --all-features` (2,183 tests),
  `cargo test -p litchi-docx --all-features`, `cargo test -p litchi-pptx
  --all-features`, `cargo test -p litchi --all-features`
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-docx
  -p litchi-pptx -p litchi-ops` (first run failed on two intra-doc links,
  fixed in 8adeb1c697)
- `python3 tools/check_crate_boundaries.py`; `python3
  tools/non_iwork_gate.py verify`
- `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir
  target/api-glossary` (the removed names banned in both DOCX profiles)
- `crates/pyo3-litchi/tests/run_tests.sh`
- `cd tools/perf-baseline && cargo check --locked --all-targets`
- capability and registry goldens: `cargo test -p litchi-ops
  --all-features --test capabilities --test registry` and the same with
  every feature but `encryption`: unchanged, not re-blessed (no capability
  changed); `python3 tools/officecli_capability_delta.py --officecli …`: no
  diff
- `LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test
  officecli_corpus`, `python3 tools/officecli_corpus.py`, `--check`:
  outcomes unchanged; only gap G3's text now cites this record

## 8. Gaps and follow-ups

- XLSX rich-text cells stay refused when a find would change them: the leaf
  has no formatting-preserving rich-string write (its shared-string run
  model keeps only a few run properties). Follow-up: a leaf per-run rewrite
  of one `si` item (or inline `is`) that keeps each run's `rPr` bytes and
  gives the changed cell its own item.
- DOCX `/body` vs OfficeCLI's `/body` sweep (F5b): kept as Litchi's tree
  reads; revisit if usage wants the OfficeCLI spelling.
- Document transactions refuse markup-compatibility content (a text box
  written as `mc:AlternateContent`), so find and merge cannot change such a
  document; the scanner reads both branches.
- PPTX masters, layouts and comments; DOCX tracked-change finds
  (`revision.*`) and paragraph properties on matched paragraphs (0273 F8,
  F9); regular expressions (record 0010 Q3).
- A paragraph-scoped DOCX find costs what `Package::edit` costs (a full
  document layout scan, about 3.7 ms of 3.8 ms on 10,000 paragraphs); the
  find itself reads only that paragraph.
