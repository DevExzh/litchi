# 0335 — DOCX P0 residue: picture, comment and SmartArt writes; chart series and axes (audit WP4)

- Status: implemented on `uo2/docx-p0`, cut from `feat/unified-ops-api` at
  `3bec513e6a`; trunk merged at `61de724e33` (records 0329 and 0332).
- Date: 2026-10-17
- Input: the definition-of-done audit of 2026-10-17, §2 WP4: in the P0
  table of design §13, DOCX "Comments", "Pictures" and "Typed SmartArt
  CRUD" were only partly met, and the DOCX `chart-series`/`chart-axis`
  elements were missing (XLSX and PPTX have them, `sgq`).
- Owned: `litchi-docx` (the eager `document::Edit` verbs below, the
  SmartArt inventory's drawing extent, a notes fix); the DOCX adapter of
  `litchi-ops` (`format/docx/*`); the DOCX rows of the capability,
  registry, CRUD-checklist and fidelity goldens; the OfficeCLI delta table
  and dispositions. Small shared hunks: `capabilities.rs` (the DOCX claim
  lists), `path/vocabulary.rs` (the WORD `node` rule), the DOCX probe of
  `tests/corpus_fidelity.rs` and a DOCX recipe in `tests/crud_matrix.rs`.
  `litchi-drawingml` is unchanged.
- `performance_claim`: every new write is within the single-op budget on a
  warm handle; the first write on a fresh handle pays one-time handle costs
  (§6).

## 1. Result

| element | before | now |
|---|---|---|
| `picture` (`/body/p[N]/picture[M]`) | `ag` (add, get; query unreachable) | `asgqr`: set `width`, `height` (one side keeps the other), `alt`, `name`, `src` (a new image part); remove; query (pictures are paragraph children for the query walker) |
| `comment` (`/comments/comment[N]`) | `agqr` | `asgqr`: set `author`, `initials`, `date`, `text`; `parentId` and `done` are typed property gaps, refused on add and set |
| `smartart` (`/smartart[N]`) | `agq` | `asgqr`: set `layout` and `text` (the whole outline); remove with its five parts and their relationships |
| `node` (`/smartart[N]/node[K]`, nested `/node[K]/node[J]`) | — | new, `asgqr`: `level` read, `text` written by add and set (the PPTX node of record 0284) |
| `series` (`/body/p[N]/chart[M]/series[K]`, `/chart[N]/series[K]`) | missing | `sgq`: `name`, `values`, `color` set; `categories` read |
| `axis` (`.../chart[M]/axis[@role=ROLE]`, `.../axis[N]`) | missing | `sgq`: `title`, `min`, `max` set; `role` read and an identity |

`edit` and `batch` reach every new target (the formatting run treats them
as barriers, as it treats charts).

OfficeCLI capability delta, DOCX (before the trunk merge, the base's
columns): elements supported **31 → 33** (`chart-axis`, `chart-series`),
missing 9 → 7, litchi-only 6 → 7 (`node`); property reads supported
**289 → 296**, gap 2 → 4 (`comment.parentId`, `comment.done`), missing
373 → 364; property writes supported **249 → 260**, partial 39 → 34, gap
508 → 502.

## 2. Leaf surface (`litchi-docx`)

Every verb is on the eager package-bound `document::Edit` and travels in
the main-document `Patch` (element replacements and graph changes), so
replay and the exact inverse apply; tests:
`crates/litchi-docx/tests/element_set_remove.rs`.

| verb | what it stages |
|---|---|
| `set_picture_at(paragraph, position, &PictureChange)` | the drawing's `wp:extent` and the shape's own `a:ext`, `wp:docPr` `name`/`descr`, and for a new image a new `media/imageN.<ext>` part with its relationship and the `a:blip` embed rewritten; the previous image part stays, unreferenced (the removed-chart policy). Positions are `Paragraph::images` order, found by `picture_spans` (a scanner kept state-for-state with the reader). Identical values are an exact no-op |
| `remove_picture_at(paragraph, position)` | the drawing, and its run when the run holds nothing else |
| `set_comment(id, &CommentChange)` | the comment's `author`, `initials`, `date` attributes, and its text as a note's text is replaced: the first run that is not the `w:annotationRef` mark holds the text with its properties, other runs go; a comment holding a field is refused (`ComplexContent`); paragraphs and their identifiers (which `commentsExtended` keys on) stay |
| `smartart_parts(data_rel)`, `set_smartart(data_rel, SmartArtParts)` | the five parts as the PPTX `opened::SmartArtParts`; the cached drawing is replaced, created (`diagrams/drawingN.xml` with its main relationship) or removed (part and relationship, a backward graph change), and the data model's drawing reference rewritten to match |
| `remove_smartart(data_rel)` | the graphic's drawing in the main story (with its run when sole), its data, layout, quick style, colors and drawing parts, and the main relationships naming them (a backward graph change); a part another main relationship names stays |

Builders follow `styles::catalog::StyleChange`: `PictureChange` and
`CommentChange` (`new`, `with_*`, `is_empty`). `smartart::Diagram` gains
`extent` (the drawing's `wp:extent`), which a regenerated cached drawing
takes.

**Bug fixed.** `set_note_text` scanned the notes part one level short
(`part_scan::scan(.., 3)`): a run's children were not indexed, so the
footnote or endnote reference mark run was taken for content and replaced,
and the first run's properties were dropped. A footnote text set therefore
published a note without its `w:footnoteRef`. It now scans to depth 4
(test `a_note_text_set_keeps_its_reference_mark`); the comment set uses the
same depth.

## 3. Adapter (`litchi-ops`)

- `format/docx/picture.rs`: `try_set_picture` (one change for every
  selected picture; a one-sided size keeps the picture's other side) and
  `remove_pictures` (later pictures first). The paragraph of a picture is
  re-resolved from its canonical path.
- `format/docx/nested.rs`: `try_set_comment`; `refuse_reply` refuses
  OfficeCLI's `parentId`, `done` and `resolved` with
  `unsupported_capability` on add and set, so a reply is never silently
  written as a plain comment.
- `format/docx/smartart.rs`: the `node` element (`DiagramNodeValue` over
  the data model's `Outline`), `try_set_smartart`, `add_smartart_node`,
  `remove_smartarts`, and the change application of record 0284 §3 (an
  authored graphic is regenerated whole; a foreign graphic takes node text
  in place, with its drawing patched or removed; node add or remove on a
  foreign graphic is refused, and `layout`/`text` rebuild it).
- `format/docx/chart.rs`: `SERIES_WRITE_TABLE`, `AXIS_WRITE_TABLE`, the
  series and axis nodes (`Value::ChartChild`), `pick_chart_child`
  (position, or `@role`), and `chart_child_edit` feeding the existing
  `ChartEdit`; chart `childCount` is its series and axes, as on PPTX and
  XLSX.
- Per-handle caches (`parts::Cache`): the SmartArt inventory (each
  graphic's outline read on first use) and the chart projections by
  relationship, as the handle already keeps its styles and numbering
  catalogs.
- Views of the root and the body are unchanged; a chart or a SmartArt
  node lists its series and axes or its nodes when its children are asked
  for.

## 4. Deltas from OfficeCLI

| # | OfficeCLI | Litchi |
|---|---|---|
| 1 | a picture is a run (`/body/p[N]/r[M]`) | a picture is `/body/p[N]/picture[M]` (record 0251d) |
| 2 | `src` set deletes the previous image part | the previous part stays, unreferenced, as for a removed picture or chart (follow-up F5) |
| 3 | comment `parentId` threads a reply through `commentsExtended`; `done` resolves | refused (`unsupported_capability`, follow-up F2) |
| 4 | comment text set replaces the comment body | the note rule: first content run takes the text, fields refused, paragraphs kept |
| 5 | no SmartArt writes (`diagram` is a Mermaid renderer) | litchi-only, the PPTX vocabulary of record 0284 |
| 6 | `chart-series` add and remove; formatting keys (`lineDash`, `marker`, `smooth`, ...) | set of `name`, `values`, `color` only (follow-up F4) |
| 7 | `chart-axis` formatting keys (`format`, `labelRotation`, `logBase`, ...) | set of `title`, `min`, `max` only (follow-up F4) |

## 5. Goldens, checklist, fidelity, corpus

- Capability and registry goldens re-blessed in both feature sets
  (`--all-features`, and every feature but `encryption`): `docx.json` and
  `docx-encryption.json` of each. Every hunk is a §1 row: `picture`,
  `comment`, `smartart` operation and property flags, `paragraph` children
  gain `picture`, `chart` children `series` and `axis`, `smartart`
  children `node`, the new `series`, `axis` and `node` elements, two
  comment property gaps. The 0330 consistency test stays green.
- CRUD checklist (`results/0289/crud-checklist.{md,tsv}`): 59 DOCX rows
  change or are new, and every claimed cell is `ok` (`series` and `axis`
  add and remove are `n/a`; `move` and `swap`, format verbs, are refused
  typed); a `node` recipe adds under the graphic the SmartArt cell added. Its first run found the new targets refused by `edit`/`batch`
  (the formatting run did not yet route them); fixed before blessing.
- Fidelity (`results/0289/fidelity.{md,tsv}`): a DOCX structural probe —
  a picture and a comment added to the first body paragraph, each set
  once, both removed — reads back the paragraph text on 56 files
  (value-exact), is refused typed on 13 (protected documents, paragraphs
  with comment or wrapped runs, invalid core properties) and one-way on 6
  (the second edit of a promoted owner); no mismatch. Every open → empty
  edit → save stays byte-exact.
- OfficeCLI corpus re-recorded (`LITCHI_BLESS=1 ... officecli_corpus`,
  `tools/officecli_corpus.py`, `--check`): no outcome changed;
  `officecli-corpus-deltas.md` reclassifies the six `word/pictures.sh`
  lines blocked by floating-picture keys from WP4 to the 0335 follow-ups.
- Delta table regenerated; `tools/officecli_dispositions.json`: the closed
  `delta:docx:chart-axis` and `delta:docx:chart-series` keys are removed
  (they had become stale), and the remaining DOCX `picture.*`,
  `comment.*`, `chart-series.*`, `chart-axis.*` gap rows are planned under
  `0335` (follow-ups F1, F2, F4). `--check` passes.

## 6. Measurements

`examples/docx_p0_residue_bench.rs`: release build, `taskset -c 11` on a
shared 32-core host (load average 7.6–7.9), 9 ABBA rounds after 2
warm-ups, 64 iterations a sample, medians, three trials
(`results/0335/trial-{1,2,3}.json`, all samples). Binary SHA-256 in
`results/0335/binaries.sha256` (built from `c8e434dd56` with `cargo build
--release -p litchi-ops --features docx --example docx_p0_residue_bench`).
Fixture: `create()`, 200 paragraphs, a comment, a picture, a SmartArt
graphic and a column chart. The leaf leg stages the `document::Edit` verb
(with the chart or diagram editor work the unified call also does) and
commits; the unified leg runs `File::set`/`remove` on the canonical path.
Both legs publish the same parts (checked outside the clock). The warm
cases reuse one opened handle; the cold cases open the bytes in both legs.

| case | leaf µs | unified µs | unified/leaf |
|---|---:|---:|---:|
| `picture_set_size` | 198.9–204.4 | 203.1–205.8 | 1.007–1.021 |
| `picture_set_alt` | 197.4–202.5 | 201.1–205.6 | 1.016–1.020 |
| `picture_remove` | 195.7–197.0 | 198.6–200.1 | 1.012–1.020 |
| `comment_set_text` | 112.0–115.7 | 116.9–119.4 | 1.024–1.044 |
| `smartart_node_text` | 716.1–723.6 | 715.7–720.8 | 0.996–1.000 |
| `chart_series_color` | 257.2–259.8 | 263.4–263.9 | 1.016–1.024 |
| `chart_axis_title` | 282.7–286.5 | 263.5–266.3 | 0.928–0.932 |
| `open_control` (open only) | 107.1–107.3 | 295.9–297.9 | **2.761–2.776** |
| `smartart_node_text_cold` | 1,032.8–1,038.8 | 1,527.2–1,538.4 | **1.479–1.483** |
| `chart_series_color_cold` | 511.8–516.2 | 945.3–946.1 | **1.832–1.847** |

- Every warm write is within the 10% single-op budget. `chart_axis_title`
  is faster unified because the leaf leg parses the chart model to find
  the value axis's id, while the handle answers it from its cached
  projection.
- **Cold-handle misses, explained, not fixed.** The first write on a fresh
  handle pays one-time handle costs the leaf leg does not: the eager
  unified open itself (`open_control`, +190 µs on this fixture, a
  pre-existing cost outside this record), the body block index the path
  resolver builds, and the SmartArt inventory or the chart projection the
  resolver reads to pick a node. Before the per-handle caches the warm
  cases measured the same way: `smartart_node_text` 1.41 and
  `chart_series_color` 1.14 (one trial, an uncommitted build of the
  previous commit). Follow-up F6.

**Existing paths, after against before.** `examples/docx_view_regression.rs`
(record 0319; public `File` API only) built on the base (`3bec513e6a`,
extracted with `git archive`) and the branch, alternated B A A B, 15 rounds
each (`results/0335/view-{strict,heading123}.txt`; binaries in
`binaries.sha256`):

| fixture | case | before ns | after ns | after/before |
|---|---|---:|---:|---:|
| `strict.docx` (a chart, a SmartArt, pictures) | paragraph format, source / eager | 13,280–13,555 / 13,048–13,334 | 13,456–13,578 / 13,068–13,153 | 1.007 / 0.994 |
| | query paragraphs + format | 31,295–31,379 / 30,930–31,505 | 31,194–31,673 / 30,975–31,114 | 1.003 / 0.994 |
| | view text | 10,522–10,527 / 9,937–10,012 | 10,816–10,867 / 10,110–10,263 | 1.030 / 1.021 |
| | root children (warm handle) | 286,301–286,858 / 287,422–288,508 | 587–588 / 375 | 0.002 / 0.001 |
| `heading123.docx` | paragraph format | 3,434–3,454 / 2,993–3,002 | 3,392–3,419 / 3,025–3,062 | 0.989 / 1.015 |
| | query paragraphs + format | 46,421–46,500 / 46,626–49,467 | 46,406–47,053 / 46,196–47,352 | 1.006 / 0.974 |
| | view text | 21,087–21,290 / 20,569–20,917 | 21,428–21,590 / 20,650–20,820 | 1.015 / 1.000 |
| | root children (warm handle) | 77,094–77,129 / 76,685–76,955 | 458–460 / 252–255 | 0.006 / 0.003 |

Paragraph reads, queries and the text view are unchanged within noise
(the text view's 2–3% on `strict.docx` does not touch the changed code).
Root listings on a warm handle no longer re-read the SmartArt inventory;
the first listing on a handle costs what it did.

## 7. Gaps and follow-ups

| # | gap | follow-up |
|---|---|---|
| F1 | floating pictures: `anchor`, `wrap`, `behindText`, `hPosition`/`vPosition`, `hRelative`/`vRelative`, `hAlign`/`vAlign`; `crop`, `decorative`, `link`, `effectExtent`, `relativeHeight`; reads of `id`, `contentType`, `fileSize` | a typed anchor model in the leaf (`wp:anchor` positions and wrap elements), then the OfficeCLI keys |
| F2 | comment replies and resolution (`parentId`, `done`), `anchoredTo`, `runStart`, `range` | author and read `commentsExtended` (`w15:paraIdParent`, `w15:done`), keeping `commentsIds` in step |
| F3 | node add or remove on a foreign SmartArt graphic (refused); node formatting | as record 0284 |
| F4 | `chart-series` add/remove and formatting keys; `chart-axis` formatting keys | the chart editor's formatting verbs (record 0270 gap 3) |
| F5 | a replaced or removed picture's image part, and a removed chart's part, stay unreferenced | a dependency sweep on publication |
| F6 | cold-handle cost of the first write (§6): the eager open (+190 µs), the body index, the SmartArt inventory or chart projection | a resolver that reads one graphic's anchor or one chart's series indexes without the full inventory or projection; the open overhead is a separate open-route follow-up |
| F7 | the SmartArt change application and the series/axis tables are duplicated from the PPTX adapter (the PPTX files belong to the parallel record 0333) | move them to `format/chart.rs` and a shared `diagram` helper in one change once both adapters are merged |
| F8 | the DOCX table-column, row-write, style and list rows the 0332 dispositions also file under WP4 | not in this brief; still `planned WP4` |

## 8. Gates

Run on the merged tree; cargo through the memory-capped shim,
`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`.

Every command exited 0.

- `cargo fmt --all --check`.
- `cargo clippy -p litchi-docx -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings`.
- `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`, and with `--features docx`.
- `cargo test -p litchi-docx --all-features` (2,447 passed), `cargo test -p litchi --all-features` (230).
- `cargo test -p litchi-ops --all-features`: `--lib` (929), `--doc` (208 + 2), and the 153 integration test files in eleven groups of up to 15 (1,382 tests), each group's executables deleted before the next.
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-docx -p litchi-ops`.
- `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify`; `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (40 profiles, 2,313 bans, 253 families); `python3 -m unittest tools.test_officecli_corpus tools.test_no_3rdparty_references tools.test_officecli_capability_delta` (65); `python3 tools/unified_bench_matrix.py --check`.
- `crates/pyo3-litchi/tests/run_tests.sh` (285 tests).
- `cd tools/perf-baseline && cargo check --locked --all-targets`.
- Goldens: `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test capabilities --test registry` and the same with every litchi-ops feature but `encryption`; after the trunk merge both sets pass unblessed. `python3 tools/officecli_capability_delta.py --officecli <OfficeCLI>` and `--check`; `LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test officecli_corpus`, `python3 tools/officecli_corpus.py` and `--check`; the CRUD checklist and the fidelity report pass unblessed on the merged tree.
- Target directory 17 GB after the gates, test executables and docs pruned.
