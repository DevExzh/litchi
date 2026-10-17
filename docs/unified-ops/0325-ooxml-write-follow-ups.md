# 0325 — OOXML write follow-ups of records 0319–0322 (tier 1)

- Status: implemented on `uo2/ooxml-followups`, cut from `feat/unified-ops-api`
  at `fde847fe90`; trunk merged at `fc274041da` (records 0315, 0323, 0324).
- Date: 2026-10-17
- Input: record 0319 follow-ups 1–3 (DOCX bookmark, field and TOC writes;
  story paragraphs; `levelN.*`), record 0321 F1 (PPTX `link`/`tooltip` as
  properties), record 0322 follow-ups 9 and 4 (XLSX note fonts and rich
  runs; the shared spreadsheet drawing code).
- Owned: the write paths of `litchi-docx`, `litchi-pptx`, `litchi-xlsx` and
  `litchi-xlsb`; `litchi-spreadsheet-drawing` (new shared modules); the DOCX,
  PPTX and XLSX adapters of `litchi-ops` (small hunks in `capabilities.rs`,
  `Cargo.toml`, DOCX `nav.rs`); their checklist and fidelity rows;
  `tools/officecli_capability_delta.py` (numbered-key matching).
- `performance_claim`: none beyond §6.

## 1. Result

| # | follow-up | get | set | add | remove |
|---|---|---|---|---|---|
| 1 | `/header[N]/p[M]`, `/footer[N]/p[M]`, `/footnote[N]/p[M]`, `/endnote[N]/p[M]` and their runs | ✅ every paragraph and run key (the story's block index) | ✅ text and every element-path setter (style, alignment, run formatting) | ✅ paragraphs under a story (start, end, position, before/after a sibling), runs under a story paragraph; bookmarks and fields in a story paragraph | ✅ paragraphs and runs; a story keeps one paragraph |
| 2a | `/bookmark[N]` | ✅ (0319) | ✅ `name` (duplicates refused) | ✅ under a paragraph (before a run, around optional `text`); under the body with `text` | ✅ markers only, content kept |
| 2b | `/field[N]` | ✅ (0319) | ✅ `instruction` (marks the field `w:dirty`), result `text` | ✅ complex field from `fieldType` (OfficeCLI's 31 types, `add --type page`… shortcuts, `if`, REF `\h`) or a raw `instruction` | — (OfficeCLI has none) |
| 2c | `/toc[N]` | ✅ (0319) | ✅ `levels`, `hyperlinks`, `pageNumbers` | ✅ `TOC \o "1-3" \h \u` with the "Update field to see table of contents" placeholder (the `refresh` engine computes it) and an optional `TOCHeading` title | ✅ its paragraphs and title, or Word's whole TOC content control |
| 3 | `level0.*`–`level8.*` on `add abstractNum` and `add num` | — | — | ✅ `format`, `text`, `start`, `indent`, `hanging`, `justification`, `suff`, `font`, `size`, `color`, `bold`, `italic`; `add num` without `abstractNumId` creates the definition ("Mode A") | — |
| 4 | PPTX `link`, `tooltip` on shape (and placeholder), run, picture, group | ✅ | ✅ | ✅ shape, picture (same native edit) | `link=none` |
| 5 | XLSX note `font.bold/italic/underline/size/color/name`, `runs` | ✅ one run: non-default `font.*`; two or more: `runs` JSON | ✅ `font.*` rebuilds one run, `runs` | ✅ | — |
| 6 | shared spreadsheet drawing | `litchi_spreadsheet_drawing::{objects, vml}` hold the drawing model, scanner, markup builders, in-place change/detach and the VML note shapes; XLSX and XLSB keep only staging. Outputs byte-identical (§5). | | | |

Delta table (`officecli-capability-delta.md`), base `fde847fe90` → this
record, on the merged tree:

| format | property reads supported | litchi-only reads | property writes supported | partial | gap |
|---|---|---|---|---|---|
| docx | 291 → 289 | 165 → 269 | 227 → 250 | 30 → 38 | 539 → 508 |
| pptx | 226 → 230 | 262 → 274 | 109 → 113 | 34 → 34 | 418 → 414 |
| xlsx | 333 → 333 | 552 → 559 | 212 → 212 | 60 → 60 | 272 → 272 |

- Elements are unchanged (31/33/30/33 supported). The new XLSX note keys and
  the PPTX run, group and placeholder `link`/`tooltip` are `litchi-only`:
  OfficeCLI's help schemas do not list them, though its handlers accept them.
- The two DOCX reads fewer are `num`'s `text` and `type`: they were counted
  as supported through the node's text; they now match their write-only
  registry entries.
- The generator now matches OfficeCLI's `<N>` keys (`level<N>.format`,
  `startOverride.<N>`) to the numbered registry spellings.

OfficeCLI corpus (`LITCHI_BLESS=1 cargo test -p litchi-cli --test
officecli_corpus`, `tools/officecli_corpus.py`, `--check`):

| family | passed | warned | gap | delta | cascade | bug |
|---|---:|---:|---:|---:|---:|---:|
| word, base | 645 | 187 | 221 | 25 | 55 | 0 |
| word, this record | **661** | 185 | **211** | 36 | 40 | 0 |
| excel | 1552 → 1552 | 26 | 51 → 51 | 4 | 0 | 0 |
| ppt | 129 → 129 | 37 | 141 → 141 | 2 | 1 | 0 |

- Word: gap −10, passed +16, cascade −15. `numbering.sh` "Mode A" adds now
  apply, so more of its later `numId=` calls meet delta D1b (positional
  `num` paths): D1b 17 → 28. `fields.sh` bookmark, field, TOC and footer
  field adds apply.
- Excel: no corpus script writes notes, so no row changes.
- PPT: `pictures-basic.sh` links fail earlier, on the harness's missing
  image files (`file_not_found`), so no row changes.

## 2. Leaf surface

| crate | addition |
|---|---|
| `litchi-spreadsheet-drawing` | `objects::{Color, Element, ElementKind, Placement, Marker, Image, Fill, Line, Font, Align, VerticalAlign, PictureSpec, ShapeSpec, Change, LimitKind, ChartFrame, Detached}` and `objects::{read, chart_frames, drawing_of, count, next_id, append, check_picture, picture_fragment, shape_fragment, change, detach}`; `vml::{Shape, shapes, next_shape_id, note_shape, new_drawing, SHAPE_TYPE}`; `Error::Opc` |
| `litchi-xlsx`, `litchi-xlsb` | `drawing::edit` re-exports the shared model; `Fill`, `Line`, `Font`, `Element` and `ShapeSpec` are aliases over the format's color (`Rgb`, `u32`) |
| `litchi-xlsx` | `workbook::comments::{TextRun, RunFont, RunColor}`, `Comment::runs` (rich runs read by the codec), `comments::edit::Edit::{set_runs, add_runs}` |
| `litchi-docx` | `document::Edit::edit_story(StoryTarget, ..)` and `StoryTarget::{Part, Note}`; `insert_bookmark_at`, `insert_field_at` (snapshot and package edits); `fields::Update::{Instruction, RenameBookmark, RemoveBookmark}`; `Fields::body_block`; `NewLevel::{start, paragraph}` |
| `litchi-pptx` | `set_shape_hyperlink` writes a `p:pic`'s or `p:grpSp`'s `p:cNvPr` click; `hyperlinks::mentions_click` |

`edit_story` wraps a header's, footer's or note's content in a synthetic
`w:document`/`w:body` (keeping the part root's namespace declarations),
runs the main document's element-path verbs on a snapshot of it, and
splices the edited content back between the story's own tags; every other
byte of the part is kept.

## 3. Adapter (`litchi-ops`)

- DOCX stories (`format/docx/story.rs`, `element::story_path`,
  `parts::{story_of, story_blocks}`): story nodes resolve their block
  children through the story's block index, cached per handle. Story
  writes run on the owning package editor.
- DOCX marks (`format/docx/marks/write.rs`): the field index
  (`document::fields`) plans renames, removals, instruction and result
  sets in one `refresh_fields` call. The read side's bookmarks and fields
  are cached per handle.
- DOCX numbering (`parts.rs` `seeded_registry!`): 108 `level<N>.*` entries
  for each of `abstractNum` and `num`, plus `num`'s template keys.
- PPTX (`links.rs` `try_set_click`, `ClickInput`): `link` and `tooltip` are
  split from a set and written last; the shape and picture adds write them
  in the same native edit. `LINKED_FRAME_PROPS` (pictures, groups) and
  `GROUP_WRITE_TABLE` are new tables.
- XLSX (`features.rs`, `feature_write.rs`): `NoteFont` and `note_runs`
  admit `font.*` and OfficeCLI's `runs` JSON.

## 4. Deltas from OfficeCLI

| # | OfficeCLI | Litchi |
|---|---|---|
| 1 | header and footer nodes list their paragraphs; unscoped `p` queries include them | story paragraphs are addressed by path only; views and unscoped queries stay in the main story, as in record 0319 |
| 2 | removing a story's last paragraph leaves a schema-invalid story | refused (`invalid_argument`) |
| 3 | a body bookmark without text is a bare start/end pair between paragraphs | refused; add it under a paragraph |
| 4 | bookmark `remove` drops the start marker only | both markers go; the content stays |
| 5 | date fields cache the current time | no cached result (no clock); `refresh` computes it |
| 6 | TOC "no page numbers" is `\z` | Word's `\n`, which the read side reports |
| 7 | `add field` returns the result run of a paragraph | the same; a body-level field's new paragraph and its result run |
| 8 | XLSX note text without `font.*` gets Tahoma 9 run properties | plain text, byte-identical to record 0320; `font.*` or `runs` write OfficeCLI's run properties |
| 9 | picture `tooltip` alone is ignored | changes the existing click's tooltip, as for groups and runs |
| 10 | `level<N>.*` aliases (`.fmt`, `.numFmt`, `.lvlText`, `.jc`) | canonical spellings only |

## 5. Byte identity

The drawing refactor was checked with a probe of public leaf APIs
(`.coord/ooxml-followups/probe/`, not committed), built on the base tree
(`git archive fde847fe90`) and on this branch. It covers, on XLSX and XLSB:
three notes added, one removed, one replaced; two shapes (solid and scheme
fills, line with dash, font, margins, rotation) and a picture added;
placement, name, description, rotation and flip changed; a removal; and an
error message. The workbook bytes, the element reads and the messages are
identical; the probe was rerun on the final tree. The plain-text note path
is unchanged code. The capability-free goldens (query corpus digests,
fidelity byte-exact rows for unchanged writes) did not move.

The fidelity report changed in one way: two POI media decks
(`EmbeddedAudio.pptx`, `EmbeddedVideo.pptx`) now pick the picture `link` for
their set-then-revert round trip. The revert keeps the relationship the
temporary link added, so the result is value-exact; the leaf keeps unused
click relationships by design (record 0321).

## 6. Measurements

`examples/ooxml_followups_bench.rs`: each scenario times the leaf call and
the unified call on the same eager owner, ABBA, 20 samples after two
warm-up round pairs, and checks outside the clock that both legs publish
the same bytes (they do, for every scenario). Release build (`cargo build
--release -p litchi-ops --example ooxml_followups_bench --features
docx,pptx,xlsx`, binary SHA-256 `abb3712e…e284`, not committed),
`taskset -c 7`, three processes, load 5. Raw lines:
[results/0325/followups-bench.jsonl](results/0325/followups-bench.jsonl).

| scenario | leaf µs | unified µs | unified/leaf (3 runs) |
|---|---:|---:|---|
| `story_paragraph_set` (`align` of `/header[1]/p[2]`) | 58.2–60.2 | 59.3–61.8 | 1.018, 1.018, 1.026 |
| `story_text_set` | 55.1–55.7 | 55.8–57.6 | 1.033, 1.014, 1.035 |
| `bookmark_rename` | 130.6–131.5 | 131.5–133.1 | 1.007, 1.007, 1.012 |
| `field_instruction_set` | 130.5–133.4 | 131.5–134.5 | 0.996, 1.008, 1.016 |
| `field_add` (`page` into `/body/p[30]`) | 92.6–95.4 | 98.7–101.0 | 1.054, 1.065, 1.065 |
| `levels_add` (three `level<N>` levels) | 108.1–111.2 | 109.8–111.0 | 1.014, 1.019, 0.998 |
| `shape_link_set` | 671.1–674.6 | 677.9–681.0 | 1.009, 1.010, 1.010 |
| `note_runs_add` | 256.0–264.4 | 257.9–266.6 | 1.006, 1.008, 1.009 |
| `note_font_set` | 179.9–181.0 | 266.2–267.9 | **1.471, 1.485, 1.486** |
| `note_text_set` (record 0320's plain set) | 177.7–178.4 | 262.6–265.4 | **1.472, 1.488, 1.493** |
| `rich_comment_get` | 76.5–77.6 | 85.0–85.7 | **1.119, 1.095, 1.108** |

- The first runs missed on six DOCX paths and on one PPTX read. These were
  fixed before this table:
  - `field_add` was 3.13×: it read the document's fields before and after
    the edit. It now returns the result-run path, as OfficeCLI does.
  - `bookmark_rename` was 1.36× and `field_instruction_set` 1.39×. They
    rescanned the read side; that scan is now cached per handle, and the
    field index proves the node directly.
  - `levels_add` was 1.51×: it reread the committed numbering part.
  - The story sets were 1.13–1.15×: they built the story block index
    twice; it is now cached per handle.
- **Miss, not fixed: note sets are 1.47–1.49×.** Resolving
  `/Sheet1/comment[19]` parses the comments part once more than the leaf
  leg, which names the cell. The plain-text set of record 0320 has the same
  cost, so this record did not introduce it. Follow-up 4.
- **Miss:** `rich_comment_get` is +10–12% (+8 µs). The node clones the
  note's runs. This is record 0320's `comment_get` (1.10) with rich text
  added.

**Existing reads, before against after.** The probe
[results/0325/read_regression_probe.rs.txt](results/0325/read_regression_probe.rs.txt)
uses only the public `File` API. It was built in both trees
(`--features docx,pptx`; binaries base `596e0a9c…210c`, after
`5a54e7dc…b962`) and run B A A B B A, pinned. Medians are in ns; the full
output is in [results/0325/read-regression.txt](results/0325/read-regression.txt).

| scenario (source / eager) | before | after | after/before |
|---|---:|---:|---|
| pptx `shape[1]` format | 4,394–4,506 / 5,516–5,546 | 4,497–4,623 / 5,588–5,693 | 1.028 / 1.020 |
| pptx query `shape` + every format | 324,828–330,476 / 488,729–493,236 | 338,772–342,433 / 499,505–505,016 | 1.046 / 1.024 |
| pptx run format | 3,392–3,450 / 4,443–4,488 | 3,356–3,406 / 4,501–4,529 | 0.990 / 1.007 |
| docx `/body/p[3]` format | 3,583–3,639 / 3,070–3,115 | 3,547–3,581 / 3,038–3,137 | 0.982 / 1.002 |
| docx query `paragraph` + every format | 46,327–47,580 / 45,924–47,312 | 46,441–47,407 / 46,029–46,449 | 0.987 / 0.993 |
| docx `/header[1]` text | 750–767 / 739–767 | 778–791 / 752–753 | 1.033 / 1.009 |
| docx `/field[1]` format (warm handle) | 11,928–12,078 / 12,418–12,769 | 288–301 / 279–293 | 0.024 / 0.023 |

- Shape and run formats now answer `link` and `tooltip`. A first build
  made a shape's format 1.65–1.83× slower, because it parsed the shape
  standalone to look for a click. `Click::of_shape` and the adapter now
  skip shapes whose markup has no `hlinkClick`, and the cost is +2–5%.
- Field and bookmark reads on a warm handle no longer rescan the document.

## 7. Gaps and follow-ups

1. Story paragraphs in views and queries (delta 1); story tables and
   content controls are readable but not written; list styles and
   hyperlink targets in stories (relationships of the story part).
2. Bookmark `text` set; field `fieldType`/`name`/`format` set (OfficeCLI
   rebuilds the instruction); field `remove`; TOC `title` set and read;
   `customStyles`/`bookmark` TOC switches; `noSeparator`, `fldLock`,
   `switches` on field add; static `IF` evaluation for the cached result.
3. `level<N>.*` aliases; `startOverride.<N>` with Mode A already works.
4. **Note set performance**: a per-handle cache of a sheet's notes would let
   `comment[N]` resolve without a second parse (§6).
5. XLSB note fonts and runs (BIFF12 rich text in `BrtCommentText`).
6. PPTX run `link` on add (runs are not added), `hyperlink` element and
   property views unified; picture `link` on media pictures keeps the
   relationship a temporary link adds (fidelity, §5).

## 8. Gates (exit 0)

Run on the merged tree (`8f91bb2cbb` and the record commit); cargo through
the memory-capped shim, `CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`. Logs:
`.coord/ooxml-followups/gate-*.out`.

| command | exit | notes |
|---|---:|---|
| `cargo fmt --all --check` | 0 | |
| `cargo clippy -p litchi-docx -p litchi-pptx -p litchi-xlsx -p litchi-xlsb -p litchi-spreadsheet-drawing -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 | |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`, and with `--features` `docx`, `pptx`, `xlsx`, `xlsb` | 0 ×5 | |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-docx -p litchi-pptx -p litchi-xlsx -p litchi-xlsb -p litchi-spreadsheet-drawing -p litchi-ops` | 0 | |
| `cargo test --all-features` of `litchi-spreadsheet-drawing` (31), `litchi-docx` (2,432), `litchi-pptx` (1,818), `litchi-xlsx` (2,464), `litchi-xlsb` (1,295), `litchi` (230) | 0 | |
| `cargo test -p litchi-ops --all-features`: `--lib` (927), `--doc` (210), 146 integration targets in 8 chunks of 20 (1,331 tests), each chunk's executables deleted | 0 | the first `--lib` run failed in `capabilities_match_live_paths_queries_and_views` (the new story path templates resolved on a fixture without stories); fixed in the record commit and rerun |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test capabilities --test registry`, and with every feature but `encryption` | 0, 0 | DOCX, PPTX, XLSX goldens |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test crud_matrix --test corpus_fidelity` | 0 | |
| `python3 tools/officecli_capability_delta.py --officecli …/3rdparty/OfficeCLI` and `--check` | 0, 0 | |
| `LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test officecli_corpus`; `tools/officecli_corpus.py`; `--check` | 0, 0, 0 | |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0, 0 | no new crate edges |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 | 40 profiles, 2,313 bans, 252 families |
| `tools/test_no_3rdparty_references.py`; `tools/unified_bench_matrix.py --check`; `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta` | 0, 0, 0 | |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 | 283 tests (new `test_record_0325_follow_ups`) |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 | |
| `cargo clippy -p litchi-ods --lib --all-features -- -D warnings` | 0 | lib only |

Target directory 19 GB after pruning.
