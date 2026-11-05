# 0292 — ODF object-format verbs and flat ODF

- Status: implemented on `uo2/odf-verbs`, cut from `feat/unified-ops-api`
  at `99cc3138eb`. Closes record 0289 §6 items 1 and 4 for ODF.
- Date: 2026-10-15
- Owned: the ODF leaf crates (`litchi-odf-common`, `litchi-odt`,
  `litchi-odg`, `litchi-odb`), `crates/litchi-ops/src/format/{odf_tree*,
  odt*, ods*, odp*}`, and small separated hunks in shared files (§6).
- `performance_claim`: paired leaf/unified medians in §4 (release, one
  pinned core of a shared host, ABBA × 12). No frozen-harness
  qualification.

## 1. Object-format verbs are element-level

Record 0289 §6 listed whole verbs the object formats refused. Every edit
verb now dispatches to the leaf for ODG, ODC, ODI, ODB and the ODF formula;
what a leaf cannot do is an element-level operation gap with a precise
reason (`Writes::gaps`, reported by `capabilities()`). The verb-level
refusal `0282 ODF object-format structure` is deleted.

| format | new | leaf work | element-level gaps left |
|---|---|---|---|
| ODG | shape `move` into another page (`Into`, `Before`/`After` an anchor on another page); cross-page shape `swap` | `edit::Edit::move_shape_to_page(page, from, to_page, to)`: the shape's bytes move unchanged as checked source splices (an empty destination page element is expanded); `StructureChange::ShapeTransferred` with its inverse | group members move with their group; a shape whose `xml:id`/`draw:id` a connector or link names, or a connector glued to shapes, stays on its page (`0292 ODG cross-page reference rewriting`); a swap of two groups on different pages (`0292 ODG cross-page group swap`) |
| ODC | `move`/`swap` of axes and of series; legend `add` (`position`), `set` `position`, `remove` | none (typed `Definition` reorder, `replace_chart`) | producer charts outside the lossless typed surface (`0288 ODC lossless producer definitions`); the single legend has no siblings |
| ODI | `remove /metadata` (clears title, author, subject, description, keywords through the metadata family); `move`/`swap` of a singleton onto itself (exact no-op) | none | the one frame cannot be added or removed (`0282 ODI single frame`); any other move position is `out_of_range` |
| ODB | `move`/`swap` of tables, of columns within a table, of queries | `Edit::{move_table, move_column, move_query}`: name-selected, final sibling position, the element's bytes relocated within its collection; self moves are exact no-ops | a column into another table (`0292 ODB cross-table column transfer`); a catalog mixing table representations and definitions (`0292 ODB mixed table collections`) |
| ODF formula | `add` of one MathML element (`tag`, `text` for tokens, `mathvariant`, any placement); `move` within or across parents; `swap` | none (`Edit::{insert_child, remove, replace}`; the leaf validates the content model) | subtrees and other attributes (`0292 formula subtree authoring`); layout elements hold no `text` |

`write::{siblings, destination, swap_pair}` resolve placements once for
every flavor (same-kind siblings, `Before`/`After` anchors, `Into` another
parent). Typed refusals: a mixed-kind swap is `unsupported_type`, a move
into its own subtree `invalid_argument`.

ODB also defers its object tree after a commit like the other object
formats: the ODB commit already parses the published catalog, so the build
cannot fail (closes the 0288 ODB single-op follow-up, §4).

Tests: `crates/litchi-odb/tests/element_moves.rs` (compact and
LibreOffice queries, tables, columns, inverse patches, bounds),
`crates/litchi-odg/tests/producer_structure.rs` (moves into and back from
another LibreOffice page, identity refusals), the five object modules of
`crates/litchi-ops/tests/odf_objects_unified.rs`, Python
`FlatOdfTests.test_flat_drawing_and_moves_across_object_formats`.

The CRUD checklist (`results/0289/crud-checklist.*`) now classifies an
object-format cell the report lists as an operation gap as `gap` with its
follow-up, and gains recipes for the ODC legend and a MathML token. ODC,
ODB and formula move/swap cells pass; ODB query and ODI cells are
`unreached` where the fixtures have one sibling.

## 2. Flat ODF

### 2.1 Read: `.fodt` and `.fodg` open (record 0277 F1)

`litchi_odf_common::flat::Parts::split` partitions a flat document's root
children into `content.xml` (scripts, font faces, automatic styles, body),
`styles.xml` (font faces, styles, automatic styles, master styles),
`meta.xml` and `settings.xml`, each wrapped in its package root with the
flat root's namespace declarations; every child's bytes are kept.
The raw helpers take a required `flat::Limits` carrier. It combines the
caller-selected physical `SourceLimits` with semantic `SemanticLimits`; flat
owners derive it from the policy retained at open. `Parts::to_bytes` publishes
exact-source members as an in-memory package under a separate bounded ZIP
scratch profile, so ZIP framing does not consume the published flat-document
output ceiling.

- `.fodt` opens as `File::Odt` over `litchi_odt::flat::Document::to_package`
  and the packaged ODT reader (body index, views, queries, Markdown). All
  46 corpus `.fodt` open (before: 0). Note bodies are excluded from the
  plain text, as for packaged ODT.
- `.fodg` opens as `File::Odg` over the packaged ODG reader: 3 of the 5
  corpus `.fodg` open; `hatch-angles.fodg` and `dashed-line.fodg` have an
  empty `office:body` and are `corrupt_file`.
- Publication of an unedited flat file is its flat bytes, byte for byte.
- Query digests for `fodt` (46 files) and `fodg` (3 files) were recorded
  by the read build, before any flat write landed
  (`tests/golden/query_corpus/digests.json`, additions only).
- `.fods` and `.fodp` keep their own flat readers; their digests are
  unchanged.

### 2.2 Write: every flat family edits its projection

`Parts::merge` maps an edit of the parts back: every top-level child an
edited part changed replaces that child's bytes in the flat document, and
every other byte is kept; a part the edit creates (`meta.xml` of a flat
document without `office:meta`) is inserted first. An edit that changes a
part's root start tag, adds or removes top-level children, changes a shared
child (font faces, automatic styles) differently in two parts, or adds a
package member is refused with a typed refusal — a flat document has no form
for it. `Parts::merge_package` is the raw external-package route and admits
the package under the supplied source policy. Owners passing an already
admitted packaged projection use `Parts::merge_package_projection`, which
uses a separate bounded synthetic-package policy instead of charging the
original flat-source ceiling a second time.
`litchi_odt::flat::Document::republish` applies it for ODT.

The unified flat handles edit their packaged projection with the families'
ordinary editors and adopt the republished flat bytes:

| family | projection | adopted after a commit |
|---|---|---|
| ODT | built at open (the read owner) | the edited package stays the projection |
| ODG | built at open (the read owner) | the edited package stays the projection; tree on first read |
| ODS | built per edit from the flat reader's bytes | `FlatSpreadsheet` reparse (the read owner) |
| ODP | built once, cached | `FlatPresentation` reparse; the edited package seeds the cache |

So `set`/`add`/`remove`/`move`/`swap`, `/metadata`, `edit` and `batch`
(with each family's coalescing) work on flat files, with the packaged
editors' checked source splices: a paragraph text set on a pretty-printed
LibreOffice `.fodt` changes exactly the paragraph's text bytes
(`flat_odf_unified.rs` asserts the full output). The verb-level flat
refusals (`ODP_FLAT`, the ODS `flat_write_refusal`, the ODT/ODG read-only
gates) are deleted; flat capability gaps are element-level
(`flat-odf-input`, `flat-odf`).

The leaves' own compact-only flat transactions (`odt::flat::Edit`,
`ods::FlatEdit`, `odp` flat `Transaction`) are not used: they refuse every
formatted producer file, and the ODP one replaces a slide's whole drawing
payload. They are left as they are.

A projection the packaged reader or editor refuses although the flat
reader admitted the file is `unsupported_capability` (follow-up `0292 flat
ODF projection admission`), not `corrupt_file`; example:
`draw-image-link.fodp` (empty `draw:master-page-name`).

Fidelity (`results/0289/fidelity.*`, open → empty edit → save, then a
set-then-revert): flat files went from 71 typed refusals (46 `.fodt` and
5 `.fodg` unopened, 18 `.fods`/`.fodp` read-only, plus 2) to none; every
flat file saves exactly. Of the reverts, 27 `.fodt` are byte-exact; the
remaining refusals are the same leaf refusals packaged files get (ODS sheet
renames on formatted XML or with dependent members, a missing ODT style).

Tests: `crates/litchi-odf-common/src/flat.rs` (split, merge, refusals,
created meta, package merge), `crates/litchi-odt/tests/flat_projection.rs`,
`crates/litchi-ops/tests/flat_odf_unified.rs` (every family: open, read,
byte-exact publication, byte-local edits, chained edits, transactions,
metadata), `ods_unified_write.rs`, `odp_unified_write.rs`,
`src/tests/{document,workbook}.rs`; Python `test_errors.py` (flat rows: a
flat formula fixture keeps the unsupported case; `.fodt` opens as a
`Document`), `test_write_api.py::FlatOdfTests`.
New fixture: `test-data/odf/flat/minimal-formula.fodf` (authored,
provenance next to it).

## 3. G-ODP-1 (ODP character formatting)

Not attempted; the follow-up of records 0278/0288 is unchanged.

## 4. Measurements

`crates/litchi-ops/examples/odf_verbs_bench.rs`, release, `taskset -c 9`
on a shared 32-core host (load 6–9 from other agents), ABBA × 12 with two
warm-ups, medians of per-call means; every write scenario asserts outside
the clock that the unified output equals the leaf output byte for byte.
Two runs (`results/0292/odf_verbs_bench-run{1,2}.txt`; binary SHA-256
`dc77d1a899ede15d250c855216c909501ca85245f992559f781a0d42aa96da51`, not
committed; build: `cargo build --release -p litchi-ops --features
odt,odg,odc,odb,odf-formula --example odf_verbs_bench`).

| format | scenario | leaf µs (run 1 / 2) | unified µs (run 1 / 2) | ratio (run 1 / 2) |
|---|---|---:|---:|---:|
| ODG (2 pages, 100 shapes) | cross-page shape move | 1,502.5 / 1,509.3 | 1,502.9 / 1,511.9 | 1.000 / 1.002 |
| ODB (100 queries) | query move | 239.9 / 239.7 | 241.3 / 240.6 | 1.006 / 1.004 |
| ODC (typed, 10 series) | series swap | 164.0 / 164.3 | 165.9 / 166.5 | 1.012 / 1.013 |
| ODF formula (50 tokens) | element move | 164.6 / 163.9 | 165.5 / 164.9 | 1.006 / 1.006 |
| flat ODT (54 KB) | open | 524.7 / 526.5 | 529.4 / 531.0 | 1.009 / 1.008 |
| flat ODT (54 KB) | paragraph text set | 968.4 / 963.0 | 968.2 / 965.1 | 1.000 / 1.002 |

The `open` leaf leg is the flat parse, the packaged projection and the
document projection; the `paragraph_set` leaf leg is the projection's
`Edit::set_text`, commit and `flat::Document::republish`.

Against the PERFORMANCE BAR: single-op writes ≤ 10 % (met, ≤ 1.3 %); open
≤ 5 % or 20 µs (met, +5 µs). The first ODB run, before the tree deferral,
was 1.263; with it, record 0288's own ODB scenario re-measured on this
branch (`odf_gaps_bench`, `results/0292/odf_gaps_bench-rerun.txt`, binary
SHA-256 `10852fc788d762b63dd7bc7bad601a4de00047cf5fa942174a2795f1c59e27f3`)
gives ODB `batch_1` 0.999 (0288: 1.347) and `batch_10` 1.001 (0288:
1.082); the other 0288 scenarios stay within 1.01. The packaged write paths
gained one flat check per publication or adoption; no packaged scenario
was re-measured beyond that rerun.

## 5. Gaps and follow-ups

Listed in `capabilities()` unless noted.

- **G-ODP-1** character formatting (§3).
- **ODG:** cross-page moves of group members and of shapes with identity
  references (`0292 ODG cross-page reference rewriting`); cross-page swap
  of two groups (`0292 ODG cross-page group swap`); 0288's ODG gaps.
- **ODB:** cross-table column moves; mixed representation/definition
  catalogs; 0282 typed table/column creation and query rename.
- **Formula:** subtree and attribute authoring (`0292 formula subtree
  authoring`).
- **ODI:** one frame per image (unchanged).
- **Flat ODF:** edits that add package members (pictures, objects) or
  change a part root; projections the packaged reader refuses
  (`0292 flat ODF projection admission`); flat ODC/ODI/ODB/formula and
  flat OTH/ODM do not open (no producer evidence; `unsupported_capability`).
  ODS flat edits rebuild the projection per edit (not a capability gap;
  measured only through the fidelity run).
- **Not measured:** flat ODS/ODP/ODG write performance, and flat batches.

## 6. Shared files

Separated hunks only: `capabilities.rs` (object verbs, flat refusals
deleted, `retains_source` rows), `open/bytes.rs` (the flat `.fodt`/`.fodg`
arms), `capabilities/tests.rs`. No iWork crate and no `PROGRESS.md`
change. `litchi-odf-common/src/core/writer.rs`: `write_exact_source_file`
became `pub(crate)` for the flat parts.

## 7. Gates (all exit 0 on the final tree)

| Command | Result |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-{odf-common,odt,odg,odb} -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ods --lib --all-features -- -D warnings` | 0 (the ODS test suite needs a purged fixture) |
| `cargo clippy -p litchi-ops --no-default-features [--features f] --all-targets -- -D warnings`, f ∈ none, odt, ods, odp, odg, odc, odi, odb, odf-formula, oth, odm | 0 each |
| `cargo test --no-fail-fast --all-features -p litchi-{odf-common,odt,odg,odb}` | 0 |
| `cargo test --no-fail-fast -p litchi-ops --all-features` | 0 |
| `cargo test --no-fail-fast -p litchi --all-features`; `-p litchi-cli --all-features` | 0, 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-{odf-common,odt,odg,odb} -p litchi-ops` | 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --test capabilities --test registry`, `--all-features` and every feature but `encryption` | identical; only ODF goldens changed |
| `python3 tools/officecli_capability_delta.py --officecli <OfficeCLI>` then `--check` | 0, 0 |
| `python3 tools/unified_bench_matrix.py --check` | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0, 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 (270 tests) |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references` | 0 |

The OfficeCLI corpus (DOCX/XLSX/PPTX) is not affected, so it was not
re-blessed.
