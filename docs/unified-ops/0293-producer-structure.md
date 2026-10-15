# 0293 — Structural edits on files from other producers

- Status: implemented on `uo2/producer-structure`, cut from
  `feat/unified-ops-api` at `99cc3138eb`. Follow-up of record
  [0289](0289-phase-g-acceptance-evidence.md) §1.3 and §6 item 5.
- Date: 2026-10-15
- Owned: `litchi-xls` (row/column insertion and deletion, the complete
  reader), `litchi-pptx` (slide removal and slide-order admission),
  `litchi-rtf` (paragraph removal, move and swap), `litchi-docx` and
  `litchi-doc` (editor admission), the XLS edit-error message in
  `litchi-ops`, and the corpus fidelity test.
- `performance_claim`: none. §6 shows no touched path got slower.

## 1. Problem

Record 0289 found that structural edits worked mainly on documents Litchi
created:

- XLS row and column add/remove refused every corpus workbook;
- PPTX slide removal succeeded on no corpus deck;
- RTF paragraph remove, move and swap refused producer RTF;
- three files whose editor refused even an empty edit that the reader
  accepts.

## 2. Method: census, then leaf fixes

`crates/litchi-ops/tests/structure_census.rs` is an ignored diagnostic. Over
every XLS, PPTX and RTF corpus file it runs these edits and tallies the full
error chain (leaf sources included):

- XLS: insert and remove a row and a column on up to three sheets, and
  remove the first existing row and column;
- PPTX: remove the first and the last slide, and move the first slide
  after the second and back;
- RTF: remove the first paragraph, add and remove a paragraph, and move or
  swap the first two paragraphs and back.

Round trips count as `ok` only when they are byte-exact.

```text
cargo test -p litchi-ops --all-features --test structure_census -- --ignored --nocapture
CENSUS_FORMAT=xls CENSUS_FILES=1 …                         # one format, with file lists
CENSUS_PROBE=<test-data path> … empty_edit_probe           # empty-edit error chain
```

### 2.1 Tally before and after

The counts are per attempt: one per sheet for XLS, one per file for PPTX and
RTF. "Before" ran on the base build, and its `ok` means "admitted", without
the byte check.

| attempt | before ok | after ok (byte-exact) | after refused |
|---|---|---|---|
| XLS row insert + remove | 48 of 305 | 176 | 129 |
| XLS column insert + remove | 48 of 305 | 165 | 140 |
| XLS remove an existing row / column | 0 | 5 / 0 | 300 / 305 |
| PPTX remove first / last slide (80 decks, 69 with one slide) | 0 / 0 | 1 / 3 | 79 / 77 |
| PPTX slide move + back | 6 of 11 | 8 | 3 |
| RTF remove first paragraph (147 files) | 0 | 60 | 87 |
| RTF paragraph move + back / swap + back (74 files) | 0 / 0 | 15 / 15 | 59 / 59 |

### 2.2 Causes found, and what each became

**XLS** (`litchi-xls`, commit `48ea079bb3`):

| cause (census, before) | count | now |
|---|---|---|
| Zero padding after the Workbook stream's final EOF was read as records ("Invalid length: expected 4") | 31 files | Raw scans stop where only zero bytes remain. The padding is kept as untouched bytes. |
| The complete reader stopped a worksheet at an embedded chart's EOF, or rejected the chart's HEADER as a duplicate ("not published by the complete XLS reader") | 24 files | Embedded substreams (BOF…EOF) are skipped in the worksheet loop. This is also a read fix: see §5. |
| Defined names (`Lbl`) refused workbook-wide | 12 files | 3-D references that name the shifted sheet are rewritten in place, at fixed width. A context-free 3-D locator was added to the formula decoder. Refused: multi-sheet spans, and relative references in names (`RgceLocRel`). |
| Chart series links (`BRAI`) and 3-D cell formulas refused | 12 | Same rewrite as defined names. |
| `ExternName` refused | 6 | Admitted when it belongs to another workbook's `SupBook`. |
| `ColInfo`, `Note`, page breaks refused | 10+ | Shifted. Ranges that reach the grid edge (`A:A`, `B:IV`) keep that edge, so insert followed by delete is an exact inverse. |
| Producer `Row` extents wider than the row's cells were normalized | 19 files | The stored extents are kept, and moved by column shifts, while they still cover the row's cells. |

The unified XLS edit error now carries the leaf's message (for example
"Unsafe edit refused: row/column movement cannot close workbook dependency
record 0x01B0"). Before, it carried the code summary "the format cannot
perform the requested operation".

**PPTX** (`litchi-pptx`, commit `12822981c4`):

| cause | count | now |
|---|---|---|
| A one-slide deck (final-slide rule) | 69 | Unchanged. This refusal is correct. |
| The removed slide's own XML was held to the copy surface (`p14:creationId`, tables, MCE) | 3 decks | Only its document element is checked. The Part leaves whole, and the inverse restores its exact bytes. |
| Any edge besides the layout edge (hyperlinks, slide jumps) | 2 | External edges and slide-jump edges are admitted. The source topology restores them as typed edges. |
| A bare `xmlns:mc` on the presentation root (LibreOffice) | 2 | Admitted, on both the eager and source paths. MCE elements and attributes are still refused. |
| Default text style vocabulary; p15 guide extensions | 4 | The default text style subtree, and presentation extensions that hold no slide position, are admitted. These are slide/notes guides, image defaults and chart tracking, admitted for removal and for slide order. |

**RTF** (`litchi-rtf`, commits `3d4cf1bdbc`, `5d31bbb189`, `c797a1de31`):

- **Before:** the lifecycle projection and the ordinary paragraph map accept only plain ASCII bodies with default formatting. Producer RTF therefore failed with "unknown RTF syntax" (63), "tables" (42), "default formatting" (15), "mixed formatting" (14) and similar causes.
- **The new seam:** `edit/producer_paragraphs.rs` cuts the body into source segments, one per root-level `\par`, and moves whole segments. It handles:
  - a `\par` that closes a formatting group;
  - the delimiter and line breaks after a `\par`;
  - an empty first paragraph.

  Destination groups are skipped.
- **Verification:** the candidate is reparsed, and it must read back as the expected paragraphs, each with its text, paragraph properties and runs. Every anchored collection must be unchanged. For a move or swap, the inverse must be plannable.
- **Fallback order:** the established paths run first, so plain bodies keep their behaviour and cost. The splice is the fallback.
- **What it still refuses:**
  - a paragraph boundary in the middle of a group (25);
  - paragraphs that inherit properties from their neighbours (15);
  - paragraphs that anchor fields, pictures or bookmarks (28);
  - tables and section breaks (9);
  - first paragraphs whose start no `\pard` or text proves (about 15).

**Editor-versus-reader admission** (commit `0144aa3091`):
- **DOCX `alt-chunk-header.docx`:** has `w:altChunk` after the body-final `sectPr`. Both transaction scanners now refuse only body content there; the tail stays outside the editable region. Now exact.
- **DOC `poi/…/test.doc`:** its PAPX bodies are padded to whole words, which leaves one zero byte after the SPRMs. That pad byte is admitted. Now exact.
- **DOCX `do-not-capture-draw-objs-on-page-draw-wrap-none.docx`:** has run-level `mc:AlternateContent` and stays refused. A run-scoped admission was tried and reverted. The Choice and Fallback branches hold different text-box text, so the raw transaction view cannot match the visible paragraph text. The `body_block_index` invariant (admitted ⇒ positions proven) caught it.

## 3. Oracle: corpus fidelity with structural round trips

`tests/corpus_fidelity.rs` now adds a structural round trip on XLS, PPTX and
RTF:
- **XLS:** insert a row at the top of sheet 1, then remove it.
- **PPTX and RTF:** move slide or paragraph 1 after 2, then move it back.

The reverted file must read back with the same node texts. A mismatch or an
untyped error fails the test.

The report is re-blessed in place:
[`results/0289/fidelity.md`](results/0289/fidelity.md) and
[`.tsv`](results/0289/fidelity.tsv).

| format | structure byte-exact | value-exact | one-way | refused (typed) | none |
|---|---|---|---|---|---|
| xls | 58 | 2 | 0 | 59 | 0 |
| pptx | 8 | 0 | 0 | 3 | 69 |
| rtf | 15 | 10 | 0 | 49 | 88 |

**Other changes in the report:**
- **Empty edits refused:** 21 → 19. DOCX went from 2 to 1 and DOC from 1 to 0. The other 18 are flat ODF files.
- **XLS set-then-revert:**
  - byte-exact: 59 → 100;
  - refused: 57 → 16.

  Sheet renames were refused by the same padding scan.

**What value-exact means here:**
- **XLS (2):** LibreOffice's `pivottable_bool_field_filter.xls` and POI's `55341_CellStyleBorder.xls`. Their producers write `DIMENSIONS` and `DBCELL` values that the row-block writer recomputes.
- **RTF (10):** plain bodies that the lifecycle projection rewrites.

The CRUD checklist (`crud_matrix`) is unchanged. Its remove/move cells run on documents Litchi made first, so its golden needed no re-bless.

`tests/producer_structure.rs` pins one construct per case:
- XLS: padding, defined names, embedded chart, column formats, and the leaf refusal text;
- PPTX: extension metadata, hyperlinks, the LibreOffice root, the notes refusal, and guide-extension moves;
- RTF: move, swap and remove;
- DOCX altChunk and DOC padding.

Leaf unit and integration tests cover each fix:
- **litchi-xls:** `structural::tests`, and `workbook::tests::worksheet_skips_embedded_chart_substreams` (fails without the fix);
- **litchi-pptx:** `tests/producer_slide_removal.rs`, which also checks that the inverse restores the bytes;
- **litchi-rtf:** `tests/producer_paragraphs.rs`;
- **litchi-docx:** `scan_differential_tests::an_alt_chunk_after_the_final_section_properties_is_an_untouched_tail`;
- **litchi-doc:** `papx_padding_tests`.

## 4. Gaps left (follow-ups)

1. **PPTX dependent parts.** A slide that owns a notes slide, picture, diagram, VML drawing, comments or OLE object is refused, and the refusal names the relationship type. Most multi-slide producer decks hit this. Removing the dependents needs an orphan proof, and a change model for the source topology that is not slide-only: its durable wire records slide images only.
2. **PPTX sections.** `p14:sectionLst` is still refused for removal and moves. Removal would need multi-span root splicing.
3. **XLS:**
   - workbook-wide CONDFMT/CF12 (`0x01B0`, `0x0879`), DV (`0x01BE`/`0x01B2`) and shared formulas (`0x04BC`);
   - split OfficeArt records;
   - column moves across MULRK/MULBLANK;
   - hyperlinks with internal locations;
   - sheets without an INDEX record or with duplicate DIMENSIONS;
   - complete-reader strictness on some producer files: zoom 1 in WINDOW2, shared-formula metadata, and `0x0006` reserved flags.

   Removing an existing row stays limited to empty rows whose deletion erases no dependency coordinate. This is the leaf's rule, unchanged.
4. **RTF.** Paragraph boundaries in the middle of a group, inherited properties, anchored fields and pictures, tables, and `\sect`. Paragraph *add* on producer RTF still uses the plain lifecycle path (7 of 147).
5. **DOCX** run-level `mc:AlternateContent` (§2.2).

## 5. Read behaviour change (XLS)

Skipping embedded substreams changes what the complete reader publishes:
- **Before:** a worksheet that held an embedded chart lost every cell after the chart, or was dropped whole.
- **After:** 89 XLS digests in `tests/golden/query_corpus/digests.json` change (re-blessed in `ccae71e0e5`), all XLS. For example, the `cell` query now matches 118,668 cells across the corpus instead of 105,418.
- **Checked:** the OfficeCLI corpus (`litchi-cli` `officecli_corpus`, `tools/officecli_corpus.py --check`) is unchanged. The capability and registry goldens are unchanged in both sets, and the delta table is unchanged.

## 6. Measurements

`examples/producer_structure_bench.rs` compared two release builds:
- **before:** base `99cc3138eb`;
- **after:** this branch at `5d31bbb189`.

The scenarios:
- the complete XLS reader open (20k-cell workbook and `FormulaEvalTestData.xls`);
- XLS row insert on a 10k-cell Litchi workbook, native and unified;
- RTF 1k-paragraph move, native and unified.

Runs were pinned to core 7, ABBA by binary, with 15 rounds after 3 warm-ups
each, done twice. Evidence:
[`results/0293/producer-structure-bench-abba.jsonl`](results/0293/producer-structure-bench-abba.jsonl).

after/before (median of medians):

| scenario | run 1 | run 2 |
|---|---|---|
| XLS complete open, 20k cells | 0.928 | 0.934 |
| XLS complete open, FormulaEvalTestData | 0.968 | 0.964 |
| XLS row insert native / unified | 0.987 / 0.992 | 0.988 / 0.991 |
| RTF move native / unified | 0.991 / 0.994 | 1.000 / 0.999 |

**Reading the results:**
- **XLS reader:** it does strictly more work per record (one comparison) and decodes no fewer records on these inputs. The ~5–7% gain is therefore not claimed: it is likely code layout.
- **Regressions:** none on any touched path.
- **Unified vs native:** the row-insert and move ratios stay at 0.99–1.00.

The binaries were deleted at the coordinator's request (disk). Their SHA-256:
before `b4b624d12826bd6e650a0094bcfb18e782bfd139b879320c3d03d70e5acaaa39`,
after `b5079c07bec110d92a926ec9ad9445e6f9f3bbf22386595bd4bd38e108cdb1cc`.
Build command: `cargo build --release -p litchi-ops --features xls,rtf --example producer_structure_bench`.

## 7. Gates (all exit 0)

- `cargo fmt --all --check`.
- `cargo clippy -p litchi-xls -p litchi-pptx -p litchi-rtf -p litchi-docx -p litchi-doc -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings`.
- `cargo clippy -p litchi-ops --no-default-features [--features xls|pptx|rtf|docx|doc] --all-targets -- -D warnings`.
- `cargo test` for `litchi-xls`, `litchi-pptx`, `litchi-rtf`, `litchi-docx` and `litchi-doc` (all features); `cargo test -p litchi-ops --all-features`; `cargo test -p litchi --all-features`.
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features` for the six changed crates.
- `tools/check_crate_boundaries.py`; `tools/non_iwork_gate.py verify`; `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py`.
- `tools/test_no_3rdparty_references.py`; `tools/unified_bench_matrix.py --check`.
- `crates/pyo3-litchi/tests/run_tests.sh` (267 tests).
- `cd tools/perf-baseline && cargo check --locked --all-targets`.
- Goldens:
  - capability and registry goldens re-blessed in both sets, with no diff;
  - `officecli_capability_delta.py` regenerated, with no diff;
  - OfficeCLI corpus test and `--check` passed;
  - fidelity and query digests re-blessed (§3, §5).

## 8. Not claimed

- The census and the fidelity round trip each test one edit per file (or per sheet). Other positions, or several edits, may meet other constructs.
- The RTF splice proves its result by readback of paragraphs, runs and collection counts. It does not compare document-level formatting outside the body, which it never touches.
