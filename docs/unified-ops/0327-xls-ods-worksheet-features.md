# 0327 — XLS and ODS worksheet features (the nodes of records 0320 and 0322)

- Status: implemented on `uo2/xls-ods-features`, cut from
  `feat/unified-ops-api` at `fc274041da`.
- Date: 2026-10-17
- Input: records 0320 (the nine XLSX worksheet feature nodes) and 0322 (the
  same nodes on XLSB). XLS is tier 2, ODS tier 3.
- Owned: `litchi-xls` (`sheet_features`, small hunks in `package`,
  `writer`, `formula`, `conditional_format`); `litchi-ods`
  (`advanced::sheet_features` and eleven `document::Edit` verbs); the XLS
  and ODS adapters of `litchi-ops` (`format/{xls,ods}/{features,
  feature_write}.rs` and small hunks in `node`, `nav`, `children`, `spec`,
  `write`); the XLS and ODS rows of the registry, capability, checklist and
  delta tables; one integration test per format and two Python tests.
- `performance_claim`: none beyond §5, which reports paired measurements of
  the new read path.

## 1. Result

| OfficeCLI element | XLS before → after | ODS before → after |
|---|---|---|
| `comment` | gap → get, query, add, set, remove (`NOTE` + `TXO` + `Obj` + note shape) | supported → adds `anchoredTo` |
| `hyperlink` | gap → get, query, add, set, remove (`HLink`, `HLinkTooltip`) | gap → get, query, add, set, remove (`text:a` over the cell text) |
| `picture` | gap → all verbs (`MsoDrawing` + `Obj`, BStore blips, PNG/JPEG) | gap → all verbs (`draw:frame` in `table:shapes`, PNG/JPEG/GIF resources) |
| `shape` | gap → all verbs (rect, roundRect, ellipse, line, textbox) | gap → all verbs (`draw:custom-shape`, twelve presets, solid fill and stroke) |
| `cf` (all rule spellings) | gap → all verbs (`CondFmt`/`CF`; `CondFmt12`/`CF12` templates) | gap → all verbs (`calcext:conditional-format`; named cell style in `styles.xml`) |
| `dataValidation` | gap → all verbs (`DVal`/`Dv`, drop-down objects) | gap → all verbs (`table:content-validation` bound by `table:content-validation-name`) |
| `autofilter` | gap → all verbs (`AutoFilterInfo`, `_FilterDatabase`, drop-downs) | gap → all verbs (`table:database-range` with filter buttons) |
| `rowbreak`, `colbreak`, `pagebreak` | gap → all verbs (`HorizontalPageBreaks`, `VerticalPageBreaks`) | gap → all verbs (`fo:break-before="page"` on derived row/column styles) |

Both formats use the XLSX property names and registry entries, so a script
written for XLSX runs unchanged where the format has the concept.

Delta table (`officecli-capability-delta.md`): XLS elements supported
**7 → 29**, ODS **8 → 29**. CRUD checklist claimed cells: XLS **59 → 158**,
ODS **78 → 166**, all `ok`. The OfficeCLI corpus has no XLS or ODS calls
for these elements; behaviour visible to the corpus did not change.

## 2. XLS leaf (`litchi_xls::sheet_features`)

The worksheet substream is split into raw records (an embedded chart's
BOF…EOF is one unit) and edited at the record level. New records go to the
MS-XLS 2.1.7.20.5 position from a rank table. `stream::splice` repairs
`BoundSheet8.lbPlyPos`, the `Index` offsets of this and later sheets and the
`ExtSST` offsets, so every record an edit does not touch keeps its bytes.
The drawing model keeps the `DgContainer`, patriarch and per-object record
groups canonical and updates the DGG clusters, `spidMax`, `cspSaved` and
`cdgSaved`, and the BStore (blips deduplicated by UID, `cRef` counted). The
Workbook stream is exact; the CFB container is re-rendered on publication.
The CF/DV formula compiler (`compile_formula`) admits every built-in
worksheet function except `CALL`, `REGISTER.ID` and `EXEC`; the cell-formula
tokenizer is unchanged.

## 3. ODS leaf (`litchi_ods::advanced::sheet_features`)

`read(content_xml, sheet)` projects drawings, formats, validations (with
their bound rectangles), the auto-filter and breaks from one span scan,
without decoding a cell. Edits are provenance splices through the existing
`XmlSplicePublication`: a start-tag edit where a whole run is touched,
otherwise the repeated row/cell run is split around the touched indexes and
the rest of the run is copied verbatim. Rows, cells or columns past the
sheet's end are appended for a binding or a break. Styles the editor writes
(`Litchi_CF_*` in `styles.xml`, `Litchi_Break_*` and `Litchi_Shape_*` in the
automatic styles) are reused when equal and dropped when nothing references
them, so add-then-remove restores `content.xml` and `styles.xml` exactly
(tested for drawings and every CF kind). CF and DV formulas are converted
between A1 and OpenFormula references (`to_odff`, `from_odff`).

## 4. Gaps and follow-ups (typed refusals)

XLS (`Code::UnsupportedProperty` with follow-ups in the registry):

- shape text formatting (font table) — "record 0327 follow-up: XLS shape
  text formatting";
- drawing effects, crops and gradient fills;
- picture `title` — wontfix (format);
- data-bar `negativeColor`, `axisColor`, `axisPosition`, `direction`
  (Office 2010 extensions);
- auto-filter `criteria0`;
- CF formatting colors must be workbook-palette entries (`invalid_value`
  otherwise); drawing geometry is in cells (lengths are `invalid_value`).

ODS:

- drawing `rotation`, `flip*`, effects, crops and shape text styling —
  "record 0327 follow-up: ODS drawing transforms, effects and text styling";
- `anchor` (cell-anchored add) — "record 0327 follow-up: ODS cell-anchored
  drawings"; drawings are added on the sheet in lengths, and x/y of an
  existing cell-anchored drawing are refused;
- CF `stopIfTrue`, `stdDev`, icon `reverse`, bar `direction` — wontfix
  (format); double underline refused;
- a styled CF on a package without `styles.xml` (litchi-created ODS) is
  `unsupported_capability` — follow-up: author `styles.xml`;
- a multi-rule `calcext:conditional-format` is read (first rule) but `set`
  is refused;
- auto-filter `criteria0` — follow-up: ODS auto-filter column criteria;
- break `min`/`max` — wontfix (format);
- hyperlinks over a range are refused (one `text:a` per cell);
- a later cell write into a validated cell is refused by the worksheet
  family, which will not drop `table:content-validation-name` — follow-up:
  model the binding on `Cell`.

## 5. Measurements

Read hot path: list every feature kind of one sheet (hyperlink, picture,
shape, two CFs, a validation, the auto-filter, a row break), ten times per
run. Binary: `examples/sheet_features_bench.rs` at the 0327 commit,
`cargo build --release -p litchi-ops --features xls,ods --example
sheet_features_bench`, SHA-256 `cf8b78ae0d76cbfa…06ee1aa`. Pinned to CPU 3,
callgrind (perf is blocked), order setup, leaf, unified, unified, leaf
(ABBA). Instructions per run; the loop cost is `mode − setup`.

| Format | setup | leaf (A / A') | unified (B / B') | leaf loop | unified loop | unified / leaf |
|---|---|---|---|---|---|---|
| XLS | 4,774,622 | 7,623,216 / 7,623,178 | 9,040,242 / 9,040,283 | 2.85 M | 4.27 M | **1.50×** |
| ODS | 387,676,743 | 442,043,032 / 442,009,161 | 501,402,303 / 501,699,975 | 54.3 M | 113.9 M | **2.10×** |

The leaf mode reads all kinds in one call (`Package::sheet_features`,
`sheet_features::read`); the unified mode reopens a handle, runs six
queries and projects every node's properties, so it is not a like-for-like
fast path: both ratios miss a 1.25× bar and are reported as misses. Most of
the ODS gap is the handle's open (content index) and the per-kind query
walks; the feature read itself is cached per sheet and snapshot.

The write path is dominated by package re-parsing: each ODS feature verb
re-opens the package archive two to four times (`content`, `part_xml`,
`splice_part`), about 47 M instructions per add on this 6 KiB fixture.
Follow-up: thread one parsed `Archive` through a verb's helpers.
Existing outputs are unchanged: every pre-existing golden other than the
XLS and ODS rows is byte-identical, and the crud checklist changed only in
the XLS and ODS rows. Evidence (callgrind outputs, logs) stays in the
coordinator scratch directory `(deleted-coord-artifact)`.

## 6. Gates

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-ods --lib --all-features -- -D warnings` (its test suite needs a purged fixture) | 0 |
| `cargo clippy -p litchi-xls -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features xls / ods / xls,ods] --all-targets -- -D warnings` | 0 (×4) |
| `cargo test -p litchi-xls` (1,295 lib tests and integration tests) | 0 |
| `cargo test -p litchi-ods --lib` (470) | 0 |
| `cargo test -p litchi-ops --all-features` (lib and 148 test targets in six chunks) | 0 |
| `cargo test -p litchi --all-features` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-xls` / `-p litchi-ods` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (40 profiles) | 0 |
| `python3 tools/check_crate_boundaries.py`, `python3 tools/non_iwork_gate.py verify` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| goldens blessed in both feature sets; `officecli_capability_delta.py` regenerated | 0 |
