# 0319 — DOCX P0 and P1 elements for OfficeCLI parity (tier 1)

- Status: implemented on `uo2/docx-p0p1`, cut from `feat/unified-ops-api`
  at `f94b388742`; trunk merged at `edac6302e7` (records 0318 and 0320).
- Date: 2026-10-16
- Input: the OfficeCLI element gaps L-DOCX-R6 (P0: styles and numbering)
  and L-DOCX-R5 (P1: headers, footers, footnotes and endnotes), then the
  P1 marks: bookmarks, fields and tables of contents, breaks, tab stops,
  text boxes and the watermark.
- Owned: `litchi-docx` (located part catalogs and their edit verbs), the
  DOCX adapter of `litchi-ops` (`format/docx/{parts,marks}*` and small
  hunks in `spec`, `nav`, `children`, `node`, `props`, `write`,
  `structure`, `nested`, `capabilities`); the DOCX rows of the capability,
  registry, CRUD-checklist and fidelity goldens; the OfficeCLI delta table.
- `performance_claim`: none; §6 reports every new and touched path.

## 1. Result

| # | element | get / query / view | set | add | remove |
|---|---|---|---|---|---|
| P0 | `/styles`, `/styles/<id or name>` | ✅ both owners; every OfficeCLI style key (`id`, `name`, `type`, `default`, `basedOn`, `next`, `linked`, flags, `uiPriority`, `customStyle`) and the style's paragraph and run properties | ✅ | ✅ | ✅ |
| P0 | `/numbering`, `abstractNum`, `num`, `level` | ✅ definition `type`, `name`, `styleLink`, `numStyleLink`; instance `abstractNumId`; level `start`, `format`, `lvlText`, `justification`, `suff`, `lvlRestart`, `isLgl`, `indent`, `hanging` | ✅ | ✅ | ✅ (removing an `abstractNum` removes its `num`s) |
| P0 | paragraph `styleName`, `numFmt`, `listStyle`, `start` | ✅ through the parts | `style`/`styleName` by id or display name (M1-Δ4) | — | — |
| P1 | `/header[N]`, `/footer[N]`, `[@type=T]` (T = `default`, `first`, `even`) | ✅ text and type, per section reference | `text`, `type` and story formatting | ✅ | ✅ |
| P1 | `/footnotes`, `/footnote[N]`, `/endnotes`, `/endnote[N]` | ✅ | `text` | ✅ anchored at a run of a paragraph | ✅ (with its reference) |
| P1 | `/body/p[N]/tab[M]` (paragraph tab stops) | ✅ `pos` (twips), `val`, `leader` | ✅ | ✅ | ✅ |
| P1 | `/body/p[N]/break[M]` (`pagebreak`, `columnbreak`, `br` aliases on add) | ✅ `breakType` | — | ✅ `type`, placement before a direct run | ✅ when its run holds only breaks |
| P1 | `/bookmark[N]`, `/bookmark[@name=X]` | ✅ `name`, `id` | — | — | — |
| P1 | `/field[N]`, `/toc[N]` | ✅ `fieldType`, `instruction`, `fldLock`, `evaluated`, result text; TOC `levels`, `hyperlinks`, `pageNumbers` | — | — | — |
| P1 | `/watermark` | ✅ `text`, `color`, `font`, `rotation`, `opacity`, `size` of the first header's text watermark | — | — | — |
| P1 | `/body/textbox[N]` | ✅ `name`, `geometry` and its story text | — | — | — |

OfficeCLI delta table, DOCX elements: supported **14 → 31**, gap
**21 → 4** (`formfield`, `ole`, `revision`, `shape`, all R4x).
DOCX property reads: supported **174 → 291**, element gap **283 → 58**;
property writes: supported **129 → 227**, gap **637 → 539** (base
`f94b388742`).

## 2. Leaf surface (`litchi-docx`)

Owned, source-located catalogs read a part once and keep each entry's byte
range: a read answers the typed direct-property keys, an edit splices the
raw part, so every byte the edit does not name is kept. Both owners and the
eager `Edit` return the same catalog type (glossary section "DOCX located
parts", record 0319).

| family | read | edit (on `document::Edit`) |
|---|---|---|
| `styles::catalog` | `Catalog::{get, resolve, entries, paragraph_properties, run_properties, is_strict}`; the scan keeps only `w:style` children, so latent styles cost nothing | `add_style`, `set_style` (`StyleChange`), `remove_style` |
| `numbering::catalog` | `abstract_nums`, `nums`, `level_of` (follows `numStyleLink`), `format_of` | `set_level`, `set_abstract_num`, `set_num_definition`, `add_num`, `add_abstract_num` (nine template levels), `add_level`, `remove_num`, `remove_abstract_num` (removes its instances), `remove_level` |
| `header_footer::catalog` | `entries_of` (natural `rId` order), `references_of` (role and section reference) | `add_header_footer`, `remove_header_footer`, `set_header_footer_kind`, `set_story_text`, `format_story` |
| notes | `Note::{footnotes_from_xml, endnotes_from_xml}` | `add_note_at` (reference run at a run position), `set_note_text`, `remove_note` (with its reference) |
| marks | `source_backed::Document::{bookmarks, text_boxes}` | `insert_break_at` |

`ElementStep::SectionProperties` addresses the body-final `w:sectPr`, which
header and footer references need. Source-backed owners read the parts
through their effective graph (`main_document_part().rels()`), so a handle
keeps answering after commits; `comments()` and `has_comments` moved to the
same route.

## 3. Adapter (`litchi-ops`)

- `format/docx/parts.rs` and `parts/write.rs`: the part elements, a
  per-handle `Cache` (one `OnceLock` per part), the registries
  `STYLE/STYLES/NUMBERING/ABSTRACT_NUM/NUM/LEVEL/STORY/NOTE_PROPS`, and the
  set, add and remove routes. A style's path step borrows its id from the
  cache (`PathStep<'f>`), so `/styles/Heading1` allocates no id.
- `format/docx/marks.rs` and `marks/write.rs`: the marks. Tab stops and
  breaks are paragraph children; bookmarks, fields, tables of contents and
  the watermark are document children; text boxes are body children. Views
  of the root and the body list only their structural children, so every
  existing view and query output is unchanged (the query corpus digests
  match the base on every file the base had).
- Paragraph `styleId`/`styleName` are writable (`Setter::ParagraphStyle`):
  a value that is a display name writes its style id; `numFmt` is read,
  `listStyle` and `start` are read and set. Coalesced sets resolve style
  names once per batch.
- Native identities (`abstractNum`/`num` `id`, level `ilvl`, note ids) are
  addressed by raw paths only (`litchi_ops::path::raw`) and hidden from
  formats and capabilities, as paragraph `paraId` already is; `name`,
  `role` and `type` stay semantic.

## 4. Deltas from OfficeCLI

| # | OfficeCLI | Litchi |
|---|---|---|
| 1 | header/footer order follows the main part's relationship list | natural `rId` order (relationships are a map in both owners, so their order was not stable) |
| 2 | `level[N]` addresses `w:ilvl=N` | `level[N]` is the Nth level; `ilvl` is a native identity, raw paths only |
| 3 | `add --type num` prints `/numbering/num[@id=N]` | prints the positional `/numbering/num[K]` (corpus delta D1b) |
| 4 | re-adding a built-in style id upserts it | refused, `duplicate_name` (corpus delta D11); `set /styles/<id>` changes it |
| 5 | a header, footer or note text set replaces whatever the story holds | record 0304: a set that would drop a field or text outside the first paragraph's direct runs is refused (`ComplexContent`); setting the same text is an exact no-op |
| 6 | removing a `num` also removes every paragraph `w:numPr` that names it (body, headers, footers, notes); removing an `abstractNum` removes only the definition | removing a `num` keeps paragraph references (Word reads a dangling `numId` as no numbering); removing an `abstractNum` also removes its `num`s (follow-up 4) |
| 7 | header and footer nodes have paragraph children | story nodes have no children yet (follow-up 2) |
| 8 | tab `pos` takes units | `pos` is an integer in twips |

## 5. Corpus

`LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test
officecli_corpus`, then `tools/officecli_corpus.py` (and `--check`), on
the merged tree:

| word scripts | passed | warned | gap | delta | cascade | bug |
|---|---:|---:|---:|---:|---:|---:|
| base `f94b388742` | 629 | 183 | 248 | 6 | 67 | 0 |
| this record | 645 | 187 | **221** | 25 | 55 | 0 |

Word gaps **248 → 221 (−27)**; DOCX-format gaps 217 → 190. By script:
`fields.sh` gap 20 → 11, `numbering.sh` 19 → 5 (cascade 27 → 11),
`sections.sh` 9 → 5. The first re-record showed 19 bugs (calls the registry
claims, refused): two built-in style re-adds and 17 `numId=` values taken
from an empty `@id=` capture. Both are behaviour this record chose (deltas 3
and 4 above), so `tools/officecli_corpus.py` names them D1b and D11 instead
of leaving them as bugs; no other bug remains.

`test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/textWatermark.docx`
(17,870 bytes, SHA-256 `9beceea6…a6c345`, MPL-2.0) is vendored with its row
in `test-data/libreoffice-core/PROVENANCE.md`; it adds one file to the
fidelity report (byte-exact) and the query corpus (DOCX files 75 → 76).
Without it the query corpus digests are those of the base.

## 6. Measurements

Release builds, `taskset -c 9` on a shared 32-core host (load 3–5).
Evidence: [results/0319/](results/0319/).

**New paths, unified against leaf.** `examples/docx_parts_bench.rs`: each
write case times one edit against the same opened eager owner (leaf: the
`document::Edit` verb and commit; unified: the `File` verb with path
resolution, admission and the returned path; both legs must publish the
same parts, checked outside the clock); reads are warm. 9 rounds after 2
warm-ups, 64 iterations a sample, ABBA, medians; three trials
(`parts-trial-{1,2,3}.json`, binary SHA-256 `ea3e763a…f65612`, built by
`cargo build --release -p litchi-ops --features docx --example
docx_parts_bench` at `1109b213a4`).

| case | leaf µs | unified µs | unified/leaf |
|---|---:|---:|---:|
| `style_set_bold` (`set /styles/Heading1 bold`) | 267.1–268.0 | 177.2–178.3 | 0.663–0.665 |
| `style_add` | 161.8–162.7 | 162.6–163.8 | 1.005–1.010 |
| `level_set_format` | 129.5–130.0 | 130.0–131.3 | 1.004–1.011 |
| `header_text` | 74.5–80.0 | 75.1–76.2 | 0.939–1.022 |
| `footnote_add` | 150.4–151.2 | 162.9–164.4 | 1.081–1.093 |
| `paragraph_style_by_name` (`style=heading 1`) | 77.9–78.8 | 79.1–79.6 | 1.007–1.018 |
| `tab_set` | 79.8–80.0 | 81.8–82.4 | 1.022–1.032 |
| `break_add` | 126.4–126.8 | 131.9–133.8 | 1.040–1.058 |
| `style_get_format` (read) | 4.7–4.8 | 5.8 | **1.211–1.228** |
| `paragraph_list_format` (read) | 2.0 | 2.2–2.3 | **1.114–1.138** |
| `tab_get_format` (read) | 2.0 | 1.7 | 0.831–0.842 |

- Every write is within the 10% single-op budget. `style_set_bold` is
  faster unified because its leaf leg builds a styles catalog from the
  package on every call (`Package::style_catalog`) to read the current run
  properties, while the handle reads its catalog once; both then stage the
  same `StyleChange`.
  `footnote_add` (+8–9%) is the one write close to the budget: beyond the leaf
  verb, the unified leg resolves the anchor paragraph's path and run
  locations to place the reference (not profiled further).
- **Two read misses, not fixed.** `style_get_format` +21–23% and
  `paragraph_list_format` +11–14% exceed 10%. The leaf legs return typed
  property structs; the unified legs resolve the path and project every key
  into `NodeValue`s (a style has ~40 keys; a list paragraph adds
  `styleName`, `numFmt`, `listStyle` and `start`, each a catalog lookup). In
  absolute terms the costs are 1.1 µs and 0.25 µs a call. Follow-up 6.

**Existing paths, after against before.** `examples/docx_view_regression.rs`
uses only the public `File` API, so the same source builds on the base
(`f94b388742`, extracted with `git archive`) and the branch; 15 rounds
after 3 warm-ups, binaries alternated B A A B (B A A B B A for the second
run). Results: `view-regression-p3.txt`, `view-regression-p2.txt`;
binaries (`cargo build --release -p litchi-ops --features docx --example
docx_view_regression` in each tree) SHA-256 before `c345fd60…0582`, after
`ece44291…896b` (the p3 run used an earlier build of the same example
without the path argument: `4ff178ce…397f`, `c6c28d97…6746`).

| scenario | before ns | after ns | after/before |
|---|---:|---:|---:|
| `heading123.docx` `/body/p[2]` format (no style, no list), source | 14,213–14,855 | 14,154–14,305 | 0.98 |
| same, eager | 13,800–14,637 | 13,334–13,874 | 0.97 |
| `heading123.docx` `/body/p[3]` format (styled), source | 3,228–3,238 | 3,406–3,407 | 1.053 |
| same, eager | 2,752–2,754 | 2,865–2,967 | 1.06 |
| `ComplexNumberedLists.docx` `/body/p[3]` format (styled, listed), source | 4,795–4,796 | 4,886–4,931 | 1.024 |
| `ComplexNumberedLists.docx` query `paragraph` + every format, source | 66,360–66,973 | 69,032–70,415 | 1.04–1.05 |
| `heading123.docx` query `paragraph` + every format, source | 45,439–45,596 | 45,524–45,799 | 1.00 |
| root `children()` (both fixtures, both owners) | 58,697–75,610 | 59,244–75,768 | 0.99–1.01 |
| `view text` (both fixtures, both owners) | 20,155–21,317 | 20,361–23,269 | 0.98–1.10 (noisy, no trend) |

- A paragraph without a style or list costs the same as before. A styled
  or listed paragraph's `format` is +2–6% because it now answers
  `styleName` (and the list keys), which the base reported as gaps: new
  output, not a slower path. Root and body views, queries without format,
  and the text view are unchanged within noise.

## 7. Gaps and follow-ups

1. **Bookmark, field and TOC writes** (OfficeCLI `add --type
   bookmark|field|toc`, `set` of a field): read-only here; 10 of the 11
   remaining `fields.sh` gaps.
2. **Story children**: `/header[N]/p[M]`, `/footer[N]/p[M]` and note
   paragraphs (`fields.sh:149–160`, `sections.sh`) need the paragraph and
   run adapter over a story part.
3. **`levelN.*` compound keys** on `add abstractNum`/`num` (OfficeCLI's
   "Mode A" that creates a definition and its levels in one call):
   `numbering.sh` lines 46, 160, 200, 222 are warned or refused; the levels
   are set one by one today.
4. **Numbering reference clean-up**: removing a `num` could clear the
   paragraph `w:numPr` that name it, as OfficeCLI does (delta 6); an owner
   decision, since it edits paragraphs the caller did not name.
5. **Watermark and text box writes**, `formfield`, `ole`, `revision`,
   `shape` (R4x), and `settings` writes such as `updateFields`
   (`fields.sh:50`).
6. **Read misses**: `style_get_format` and `paragraph_list_format` (§6).
   A projection that formats only the keys a caller asks for, or a
   per-style key cache, would close them.

## 8. Gates (exit 0)

Run on the merged tree (`edac6302e7` and the record commit); cargo through
the memory-capped shim, `CARGO_BUILD_JOBS=6`.

- `cargo fmt --all --check`.
- `cargo clippy -p litchi-docx -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings`.
- `cargo clippy -p litchi-ops --no-default-features [--features docx] --all-targets -- -D warnings`.
- `cargo test --all-features` of `litchi-docx` (2,429) and `litchi` (230);
  `litchi-ops` `--lib` (922), `--doc` (208 + 2) and its 142 integration
  tests in ten groups of 15 (1,283 tests), each group's executables
  deleted before the next.
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-docx -p litchi-ops`.
- `tools/check_crate_boundaries.py`; `tools/non_iwork_gate.py verify`;
  `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py` (40 profiles, 2,313
  bans, 238 families); `tools/test_no_3rdparty_references.py`;
  `tools/unified_bench_matrix.py --check`;
  `python3 -m unittest tools.test_officecli_corpus`.
- `crates/pyo3-litchi/tests/run_tests.sh` (280 tests; the story-gap test
  now checks that headers and footnotes are read).
- `cd tools/perf-baseline && cargo check --locked --all-targets`.
- Goldens re-blessed after review: capabilities and registry (DOCX; both
  feature sets), `results/0289/crud-checklist.{md,tsv}` (DOCX parts and
  marks rows), `results/0289/fidelity.{md,tsv}` (one new byte-exact file),
  `tests/golden/query_corpus/digests.json` (the new file only; the base
  digests match without it), `officecli-capability-delta.md`, the OfficeCLI
  corpus (`outcomes.tsv`, `officecli-corpus-deltas.md`, `--check`).
- Target directory 16 GB after the gates.
