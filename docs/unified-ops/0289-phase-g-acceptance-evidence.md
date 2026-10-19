# 0289 — Phase G acceptance evidence: CRUD checklist, corpus fidelity, benchmark matrix

- Status: implemented on `uo2/phase-g`, cut from `feat/unified-ops-api` at
  `9fe45629ab`. Plan [0248](0248-coordinator-restart-plan.md) §6 Phase G,
  steps 25–27 (M5, design §14 layers 2 and 3). Step 28 (delete
  `litchi::iwork::Document`) waits for iWork M6.
- Date: 2026-10-14
- Owned: new tests `crates/litchi-ops/tests/{crud_matrix,corpus_fidelity,phase_g_regressions}.rs`,
  the example `crates/litchi-ops/examples/phase_g_bench.rs`, the tool
  `tools/unified_bench_matrix.py`, seven vendored fixtures, and four small
  bug fixes in separate commits (§4).
- `performance_claim`: none for the API. The fixes are measured
  before/after only (§5): no change beyond noise.

## 1. Step 25 — CRUD checklist per format

`tests/crud_matrix.rs` builds the checklist from `Capabilities::of(format)`
and the registry. It crosses every element kind of the 18 non-iWork formats
with the 13 verbs get, find, query, view, set, add, remove, move, swap, edit,
batch, create and save, then exercises every claimed-supported cell.

**How each cell is classified:**
- **supported:** the format verb is supported and the element declares the operation.
- **gap:** the format verb, the element, or a property write gap names a follow-up. The follow-up is in the TSV.
- **n/a:** the operation is not offered and nothing declares a follow-up.
- **format-verb:** move and swap, which the report does not declare per element.

**How each claimed cell is exercised:**
- Reads `get`, `find` and the node view on up to eight instances. Query uses the element type selector.
- `set` tries every settable property:
  - Values are generated from the registry type and constraint.
  - Changed values are tried first. A no-op is used only if no changed value is accepted; the TSV marks it.
- `add` tries the declared parents, or a recipe for adds that need data (charts, pictures, tables).
- Each successful `add` is repeated once, so `move` and `swap` get a pair of siblings that Litchi made.
- `remove` targets instances on documents Litchi made first, then the corpus.
- `edit` and `batch` replay the successful `set` through `File::edit` and through an atomic `File::batch`.
- `create` and `save` run per format and reopen the result.
- Every write is published with `to_bytes` and reopened.

**Inputs:**
- `create()`;
- a small seed per format, built with the unified verbs;
- 4–13 pinned corpus fixtures per format. The ignored `discover_fixtures` test chose them as the smallest files that hold each element kind.

**Output and runtime:**
- The checklist is golden-tested: [`results/0289/crud-checklist.md`](results/0289/crud-checklist.md) (tables) and [`crud-checklist.tsv`](results/0289/crud-checklist.tsv) (every cell with its evidence).
- `LITCHI_BLESS=1` rewrites it. A run takes about 11 s in a debug build.

Summary (claimed cells are the per-element read/write/edit/batch claims plus
create/save; move/swap are counted separately):

| format | elements | claimed | ok | FAIL | unreached | gap | n/a | move/swap ok | move/swap refused |
|---|---|---|---|---|---|---|---|---|---|
| docx | 40 | 120 | 120 | 0 | 0 | 325 | 45 | 8 | 8 |
| xlsx | 26 | 92 | 92 | 0 | 0 | 195 | 29 | 2 | 14 |
| pptx | 37 | 189 | 178 | 3 | 8 | 156 | 86 | 2 | 28 |
| xlsb | 26 | 58 | 54 | 0 | 4 | 234 | 30 | 2 | 6 |
| doc | 13 | 86 | 86 | 0 | 0 | 0 | 57 | 2 | 6 |
| xls | 26 | 59 | 59 | 0 | 0 | 247 | 18 | 2 | 6 |
| ppt | 36 | 76 | 65 | 1 | 8 | 338 | 34 | 2 | 8 |
| rtf | 21 | 154 | 153 | 1 | 0 | 0 | 77 | 2 | 6 |
| odt | 22 | 205 | 205 | 0 | 0 | 0 | 37 | 29 | 5 |
| ods | 26 | 78 | 78 | 0 | 0 | 221 | 21 | 2 | 8 |
| odp | 30 | 170 | 165 | 1 | 4 | 143 | 39 | 4 | 10 |
| odg | 4 | 32 | 28 | 0 | 4 | 8 | 12 | 0 | 0 |
| odc | 4 | 28 | 22 | 2 | 4 | 12 | 12 | 0 | 0 |
| odi | 3 | 23 | 23 | 0 | 0 | 12 | 4 | 0 | 0 |
| odb | 4 | 36 | 36 | 0 | 0 | 8 | 8 | 0 | 0 |
| oth | 22 | 205 | 177 | 1 | 21 | 0 | 37 | 26 | 0 |
| odm | 22 | 205 | 184 | 0 | 17 | 0 | 37 | 26 | 2 |
| odf | 2 | 18 | 18 | 0 | 0 | 6 | 2 | 0 | 0 |

### 1.1 Claimed cells that failed

Sixteen claimed cells failed on the first runs. Four fixes (§4) closed seven
of them; the other nine are recorded bugs, listed in the test's
`KNOWN_FAILURES`. The
test fails on any new failure, and also when a listed one starts passing.

| # | cell | cause | follow-up |
|---|---|---|---|
| F1 | pptx and odp `placeholder` query | The query compiles `placeholder` to its own element kind. Placeholder nodes are shapes with a placeholder binding (PPTX) or presentation class (ODP), so the type selector matches nothing; `shape[type=placeholder]` works. | Compile `placeholder` as `shape` with the placeholder test. |
| F2 | pptx `chart` view | litchi-drawingml rejects the negative `c:axId` values some producers write (ST_UnsignedInt) in `ooxml/pptx/radar-chart.pptx`. `get` works; the node view's chart decode fails with `corrupt_file`. | Leaf: lenient signed axis IDs, or a typed per-chart refusal. |
| F3 | pptx `ole` remove | Every corpus OLE object sits in markup compatibility, which the source-backed removal refuses. `add` cannot make one. | Removal through MCE branch selection. |
| F4 | ppt `slide` set (`hidden`) | The leaf rewrites an existing SSSlideInfoAtom. Producer slides and `create()` slides have none, and the leaf does not insert one. Also fails on `create()` + add slide. | Leaf: insert the atom. |
| F5 | odc `axis`/`series` remove | The only ODC fixtures (odfdom) hold content outside the leaf's lossless edit surface. `create()` has no chart, and ODC has no `add`. | ODC chart content in `create`, or a wider edit surface. |
| F6 | oth `metadata` set | `create()` packages from the OTH and ODM leaf builders have no meta.xml, which the ODT metadata editor requires. ODM passes here only on corpus files. | Leaf builders write meta.xml, or the editor creates it. |
| F7 | rtf `picture` remove | The only producer pictures (`n192129.rtf`) are refused by the exact source editor ("replacement bytes do not retain the declared picture format"). RTF has no picture `add`. | Leaf picture-format check. |

### 1.2 Unreached claimed cells (70)

No fixture in test-data, and no small file in the LibreOffice, POI or Tika
trees, holds these elements. Each cell is listed in the TSV.

- **pptx:** `modernComment`, `reply`.
- **xlsb:** `table`.
- **ppt:** `table`, `connector`.
- **odp:** `comment`.
- **odg:** `layer`. LibreOffice writes the document layer set in styles.xml; the adapter projects page-level layer sets only.
- **odc:** `legend`.
- **oth and odm:** `frame`, `picture`, `textbox`, and oth `style`. The ODT cells cover the same adapter code.

### 1.3 Findings beyond the claimed cells

These are typed refusals, so they are not failures, but the owner should
know about them. The TSV has the evidence.

- **Structure edits work mainly on documents Litchi made:**
  - XLS row/column add and remove refuse every corpus workbook, with "no worksheet grid" or "cannot perform". `SimpleWithAutofilter.xls` even reports `corrupt_file`, although it reads fine.
  - PPTX slide remove refuses every corpus deck under 200 KB.
  - RTF paragraph remove/move/swap refuse producer RTF, which is outside the lifecycle editor's policy.
- **Moves and swaps:** most element kinds outside slides, sheets and paragraphs refuse them with `unsupported_type`.

## 2. Step 26 — full-corpus fidelity regression

`tests/corpus_fidelity.rs` runs every file under test-data/ that opens as a
non-iWork format: 813 of 1,926 files. It takes about 35 s in a debug build,
so it always runs the whole corpus and needs no subset or shard.

Each file goes through `open` (file-backed) → `edit` with no mutation →
`save_as`, and the saved bytes are compared with the input:
- **Signed files** go through `open` → `save_as`.
- **Locked files** are only counted.
- **Typed edit refusals** are reported with their cause and the `open` → `save_as` result.

Where it is cheap, a set-then-revert round trip follows:
1. Find the first of up to six nodes with a settable text property.
2. Set the property to a new value, then back to the original.
3. Reopen: the value must be the original.
4. Report whether the bytes also equal the input.

The report is golden-tested:
[`results/0289/fidelity.md`](results/0289/fidelity.md) and
[`fidelity.tsv`](results/0289/fidelity.tsv).

| format | exact | differs | refused | signed-exact | locked | revert byte-exact | revert value-exact | one-way | refused | none |
|---|---|---|---|---|---|---|---|---|---|---|
| doc | 43 | 0 | 1 | 0 | 3 | 0 | 36 | 0 | 4 | 3 |
| docx | 68 | 0 | 2 | 3 | 3 | 36 | 0 | 0 | 5 | 27 |
| odb | 6 | 0 | 0 | 0 | 0 | 0 | 5 | 0 | 0 | 1 |
| odc | 2 | 0 | 0 | 0 | 0 | 0 | 2 | 0 | 0 | 0 |
| odf | 4 | 0 | 0 | 0 | 0 | 0 | 4 | 0 | 0 | 0 |
| odg | 9 | 0 | 0 | 0 | 0 | 0 | 9 | 0 | 0 | 0 |
| odm | 3 | 0 | 0 | 0 | 0 | 0 | 3 | 0 | 0 | 0 |
| odp | 15 | 0 | 5 | 0 | 0 | 0 | 13 | 2 | 0 | 0 |
| ods | 13 | 0 | 13 | 0 | 0 | 0 | 3 | 0 | 10 | 0 |
| odt | 23 | 0 | 0 | 0 | 0 | 0 | 20 | 0 | 3 | 0 |
| ppt | 27 | 0 | 0 | 0 | 4 | 0 | 0 | 0 | 0 | 27 |
| pptx | 80 | 0 | 0 | 2 | 0 | 0 | 4 | 0 | 50 | 26 |
| rtf | 162 | 0 | 0 | 0 | 0 | 14 | 49 | 1 | 61 | 37 |
| xls | 119 | 0 | 0 | 0 | 3 | 59 | 3 | 0 | 57 | 0 |
| xlsb | 18 | 0 | 0 | 0 | 0 | 0 | 18 | 0 | 0 | 0 |
| xlsx | 179 | 0 | 0 | 2 | 1 | 0 | 173 | 1 | 5 | 0 |
| **all** | **771** | **0** | **21** | **7** | **14** | **109** | **342** | **4** | **195** | **121** |

**No-op saves are always exact.** Every no-op save is byte-identical: 771
empty edits and 7 signed saves. The 21 typed refusals also save exactly
through `open` → `save_as`. No format needs a documented normalization
(`NORMALIZING` is empty).

**Root causes of the 21 refusals:**
- **18 flat ODF files** (`.fods`, `.fodp`): read-only in the unified API (the M6 flat-ODF write follow-up).
- **`do-not-capture-draw-objs-on-page-draw-wrap-none.docx`:** markup compatibility in the body. The source-backed DOCX transaction refuses even an empty edit.
- **`alt-chunk-header.docx`:** the editor finds a body-final `sectPr` that is not the last child (`corrupt_file`), where the reader does not.
- **`poi/test-data/document/test.doc`:** the DOC editor reports `corrupt_file` where the reader does not.

The last two are reader/editor admission mismatches. They are worth a leaf
look, but they are typed.

**Not opened:** 46 `.fodt` and 5 `.fodg` files are refused with
`unsupported_capability`, because flat ODT and ODG are not read through
`File`. The rest are deliberately corrupt fixtures (`corrupt_file`) or
legacy versions outside scope (`unsupported_format`).

**Revert:**
- **byte-exact:** where both writes are exact splices, such as DOCX text and XLS cells.
- **value-exact:** everywhere else. The leaves re-serialize the parts they touch: ZIP members, XML attribute order and similar.
- **refused:** the first set was a typed refusal, for example master shape text, sheet renames refused by dependent members, or ODS LibreOffice-file rules.
- **one-way:** in four files the forward set was admitted but the set back was refused:
  - **RTF `tdf107480.rtf`:** the original paragraph text does not survive RTF validation.
  - **ODP `tdf169979` (×2):** the original slide name is referenced by links or custom shows. The forward rename of a referenced name was admitted, so the leaf checks only the target name. Is a referenced page renamed without updating its links? Worth a leaf look.
  - **XLSX `49609.xlsx`:** after one mutation, "owned OPC source preservation is unavailable", so a second edit of the first edit's output is refused.

### 2.1 Determinism bug found and fixed

Two runs of the report gave different byte counts for ODF set-then-revert
writes. The cause was in the ODF package writer:
- It re-declared the source manifest's directory entries (LibreOffice's
  `Configurations2/…`) by iterating the manifest model, a `HashMap`.
- So two identical writes of the same package published different
  manifests, against ADR 0006 (serialization is deterministic).

The fix (commit 4, §4) makes the three copy paths iterate the entries in
path order. `phase_g_regressions::odf::identical_writes_publish_identical_bytes`
fails without it and passes with it. The fidelity and checklist goldens are
now stable from run to run (checked twice).

## 3. Step 27 — benchmark matrix

`tools/unified_bench_matrix.py` (stdlib only) merges two sources:
- the machine-readable evidence under `docs/unified-ops/results/` for records 0249–0284 (975 rows);
- a curated JSON of the ratios that exist only in record prose (41 rows, each quoting its record and section).

It writes one table — format, operation, scenario/size, leaf, unified,
ratio, budget, status, record, source and explanation — to
`results/0289/bench-matrix.tsv` and
[`.md`](results/0289/bench-matrix.md) (record 0348 split the TSV per format:
`results/0289/bench-matrix/<format>.tsv`).

**Budgets** come from the goal's PERFORMANCE BAR:
- open: 1.05 or +20 µs;
- single-op write: 1.10;
- 1k-op batch: 2.0;
- full queries: the 1.10 review trigger of record 0262.

**`--check`** regenerates the table in memory. It fails on stale output, or
when a miss has no explanation or follow-up. Records 0285–0288 do not exist
on this base.

Result: 1,016 rows. 508 meet their budget, 247 have no budget, 170 are
misses superseded by a later record, and 91 are open misses. Every miss
cites its explanation or follow-up. The open misses by area:

| area | rows | worst | record / follow-up |
|---|---|---|---|
| full queries: XLSX/XLSB/PPTX/DOCX/ODT/ODS | 50 | ODS 41×, XLSX up to 180× memory scan | 0280 G1–G4: the owned node cost and the per-match `fstat` the owner kept. ODS eager A1 labels (G3). DOCX node size (0262 F4). |
| XLS get/query/positional writes | 20 | 5.3× | 0281 F3 → 0280; 0259 positional rows 1.11–1.13 |
| RTF single-target writes | 16 | 1.50× | 0280 G5: plan cache, lighter one-command edit |
| DOCX single paragraph set | 1 | 1.19× | 0281 F1 / 0280 G6 |
| PPTX open + first metadata set | 1 | 1.31× | 0281 F2 (promotion re-inflates members) |
| XLSX chart set at 10 and 100 charts | 2 | 1.35× | 0283: memoize the drawing inventory |
| XLSX 1k-row memory scan | 1 | +6% | 0280 (code layout) |

The tool found five inconsistencies; the `.md` cross-check section has all of them:
1. **0281:** its §3 quotes ratio ranges, but its committed evidence holds one run per harness. Three single values fall below the quoted range minimum (cell `align`, cell `text`, DOCX find).
2. **0275:** an unreported RTF run-add run at 1.124, at 10k dense runs. Probably host noise: the other run is 1.019.
3. **0280:** query rows that swing 3–10× between runs (`xlsx_50000_bytes_all_cell`, the ODS pairing).
4. **0262:** the XLSB source metadata ratio, 1.392, predates the 0261 fix.
5. **0261:** a set+save pair at 1.102 that its record table omits; 0281 superseded it.

## 4. Fixes (separate commits)

1. **DOCX edit/batch, hyperlink targets and charts** (`format/docx/write/coalesce.rs`).
   - The formatting run took hyperlink `url` sets, which need the owning package editor, so a source-backed owner failed with `internal_error`.
   - It also refused chart targets, which `File::set` writes.
   - Both are now barriers that fall through to the ordinary setter.
2. **DOC run text in edit/batch** (`format/doc/write.rs`). The paragraph-text plan now returns `None` for non-paragraph targets, as its contract says, instead of refusing them.
3. **RTF `body` add claim** (`format/rtf/spec.rs`).
   - `ADDABLE` listed `body`, so the report claimed an `add` the writer refuses.
   - The capability unit test, the registry golden and the delta table are updated. RTF also joins the capability goldens.
4. **ODF manifest determinism** (`litchi-odf-common/src/core/writer.rs`), §2.1.

`tests/phase_g_regressions.rs` covers all four. Each test fails without its
fix.

## 5. Measurements

The example `phase_g_bench.rs` was run on two builds:
- **before:** base `9fe45629ab` with fixes 1, 2 and 4 reverted;
- **after:** this branch.

Both are release builds, pinned to core 13 of a shared host, ABBA × 2 per
run, 15 timed rounds after 3 warm-ups, run twice. Medians of the four
per-leg medians (µs):

| scenario | run 1 before → after | run 2 before → after | after/before (run 1, run 2) |
|---|---|---|---|
| `docx_edit_bold_1k` (1,000 run sets in one edit) | 22,545 → 22,330 | 22,822 → 22,563 | 0.991, 0.989 |
| `doc_edit_text_200` | 94,248 → 94,308 | 95,623 → 94,934 | 1.001, 0.993 |
| `odt_metadata_set_publish` (LibreOffice ODT) | 763 → 772 | 768 → 764 | 1.012, 0.995 |

**ODT run 1 (+1.2%) is not added work.** callgrind of that scenario alone
counts 285,035,928 instructions before and 285,116,617 after: +0.03%, about
4.5 k instructions per write for the sort. That points to code layout; run 2
is at 0.995.

**Evidence:**
- `results/0289/phase-g-bench-abba-{1,2}.jsonl`.
- Run 2's binaries are not committed (AGENTS.md). Their SHA-256: before `1a6e57e37b05519ce02d13380fbe95e36768bd19a5aeeb343200a82488968246`, after `efcd35f7e0d906c4f48474762c7343a20eb4704dcdc5096ad5d07123dfc3c340`.
- Build command: `cargo build --release -p litchi-ops --example phase_g_bench --features docx,doc,odt`.

## 6. What remains for the Definition of Done (iWork M6 excluded)

1. **Verb-level gaps in the ODF object formats.** These are whole verbs, not element gaps (record 0282 follow-up), so they need leaf work or owner acceptance:
   - move/swap are refused for ODG, ODC, ODI, ODB and ODF formula;
   - `add` is refused for ODC, ODI and ODF formula;
   - `remove` is refused for ODI.
2. **The nine claimed-but-failing cells, F1–F7** (§1.1), and the 70 unreached cells that need fixtures (§1.2).
3. **The 91 open benchmark misses** (§3). The largest are full queries (0280 G1–G4) and XLS reads. Each needs engineering or an owner-recorded explanation (D5).
4. **Flat ODF:**
   - `.fodt` and `.fodg` do not open through `File`;
   - `.fods` and `.fodp` are read-only.
5. **Edit admission is stricter than the reader:**
   - empty edits are refused on DOCX files with markup compatibility in the body, and on two files the editor calls corrupt;
   - one-way reverts (§2), including XLSX chained edits after a promotion.
6. **Python:** §14 layer 3's pytest mirror of the CRUD and fidelity scenarios is not written. This record covers Rust only.
7. **Owner decisions** pending in PROGRESS "Awaiting owner confirmation": D1–D5 and the older provisional defaults.
8. **Step 28** (`litchi::iwork::Document`) and all iWork work wait for D1/D2.

## 7. Not claimed

- The checklist exercises at least one instance per claimed cell. It does
  not prove every property, every placement or every producer.
- Generated values are chosen to be accepted; a cell passes when one
  property (or one add recipe) works.
- The benchmark matrix collects earlier measurements; nothing was
  re-measured except §5.

## 8. Gates (all exit 0)

- `cargo fmt --all --check`.
- `cargo clippy -p litchi-odf-common -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings`.
- `cargo clippy -p litchi-ops --no-default-features [--features docx|doc|rtf|odt|ods|odp|odg|oth] --all-targets -- -D warnings`.
- `cargo test -p litchi-odf-common --all-features`; `cargo test -p litchi-ops --all-features`; `cargo test -p litchi --all-features`.
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ops -p litchi-odf-common`.
- `tools/check_crate_boundaries.py`; `tools/non_iwork_gate.py verify`; `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py`.
- `crates/pyo3-litchi/tests/run_tests.sh`.
- `cd tools/perf-baseline && cargo check --locked --all-targets`.
- `tools/unified_bench_matrix.py --check`; `tools/test_no_3rdparty_references.py`.
- Goldens re-blessed in both sets (`--all-features` and every feature except `encryption`); `tools/officecli_capability_delta.py` regenerated (only the RTF hash changed). OfficeCLI corpus behaviour is unchanged by these fixes, so it was not re-blessed.
