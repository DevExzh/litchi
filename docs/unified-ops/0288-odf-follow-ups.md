# 0288 — ODF follow-ups: shared metadata family, producer drawings, chart and image writes, object-format coalescing

- Status: implemented on `uo2/odf-gaps`, cut from `feat/unified-ops-api`
  at `9fe45629ab`. Closes the top follow-ups of records
  [0276](0276-odt-unified-adapter.md)–[0278](0278-odp-unified-adapter.md)
  and [0282](0282-remaining-odf-formats.md).
- Date: 2026-10-14
- Owned: the ODF leaf crates (`litchi-odf-common`, `litchi-odt`,
  `litchi-ods`, `litchi-odp`, `litchi-oth`, `litchi-odm`, `litchi-odg`,
  `litchi-odc`, `litchi-odi`, `litchi-odb`, `litchi-odf-formula`),
  `crates/litchi-ops/src/format/{odf_metadata,odf_odp,odf_tree*,ods*,odp*}`,
  and small separated hunks in shared files (§8).
- `performance_claim`: paired leaf/unified medians in §6 (release, one
  pinned core of a shared host, ABBA × 12). No frozen-harness
  qualification.

## 1. Metadata family for every ODF leaf (ADR 0034 §7; 0282 G7, 0277 M1)

`litchi_odf_common::metadata` is the shared engine:

- `Field` (title, subject, description, author = `meta:initial-creator`,
  lastModifiedBy = `dc:creator`, language, keywords), `Assignments` (staged
  values; `publication(&OwnedPackage)` returns the `meta.xml` splice), and
  the generic `Edit<P>`, `Commit<P>`, `Patch<P>`, `Diagnostics` bound to a
  leaf through the `Owner` trait (`metadata_source`, `adopt_metadata`,
  `metadata_bytes`).
- A commit replaces only the touched `office:meta` children with balanced
  element splices (LibreOffice's prolog newline and every other byte and
  member kept); a package without `meta.xml` gains one declaring the
  content's `office:version`. Dates, generator and statistics are never
  written (no clock). Exact no-ops return the source; a changed write of a
  signed or encrypted package is refused (`Error::Core(Unsupported)`).
  `Patch::{is_noop, inverse}`; `Patch::apply` checks the exact source bytes.
- ODS, ODP, OTH, ODM, ODG, ODC, ODI, ODB and the formula crate expose
  `Package::{edit_metadata, apply_metadata}` and `litchi_<crate>::metadata::
  {Edit, Patch, Commit, Diagnostics, Error, LimitKind, Field}` (type aliases
  of the generic types). ODT's family now runs on the same engine (its
  `Field` is the shared one). The formula crate gains `Package::metadata()`.
- ODS keeps the primary `document::Edit::set_metadata` (owner decision,
  0277 M1); its unrelated contextual `metadata::Commit` is renamed
  `metadata::XmlCommit` so the family's `Commit` takes the name.
- Glossary: the nine families are registered in their profiles
  (`tools/api_glossary.json`) and the metadata-mutation row flips.
- Unified: `set("/metadata", ..)` for ODS, ODP, ODG, ODC, ODI, ODB and the
  formula goes through the family (`format::odf_metadata`, shared by all
  of them); the seven keys above take text, `null` removes, other
  root-metadata keys are `unsupported_capability`, unknown keys
  `unsupported_property`. Create-time metadata works for every object
  format. ODS and ODP now also write `lastModifiedBy` and `language`.
- Read gap kept: the shared `meta.xml` → `litchi_core::Metadata`
  conversion omits `dc:language`, so `File::metadata()` does not report a
  written language for ODS, ODP and the object formats. Mapping it changes
  the facade Markdown output, which `crates/litchi/tests/golden/
  markdown_corpus.tsv` pins as fixed reference data (12 differences), so
  it is left for an owner decision.

Tests: `crates/litchi-<crate>/tests/metadata_family.rs` in nine crates
(fresh packages without `meta.xml`, producer packages, second edits,
no-ops, inverse/stale patches, removal, refused text), unit tests of the
splice planner in `litchi-odf-common`.

## 2. ODG on LibreOffice drawings (0282 G1)

Structural edits used to fall back to a whole-part rewrite that the
compact-XML audit refuses on every producer `content.xml`. They are now
checked source splices on the primary `Edit`:

- `insert_page`/`add_page` (without a generated transition style),
  `remove_page`, `insert_shape`/`add_shape`, `remove_shape`, and the new
  `move_page(from, to)` and `move_shape(page, from, to)` stage balanced
  markup splices; inserting into an empty owner expands the whole
  `<owner …/>` element. Markup splices that touch (two pages appended at
  one point, a removal next to an insertion) join into one publication
  edit; a splice that cannot map onto untouched source bytes falls back to
  the whole-part rewrite (still compact-only). The publication classifies
  each splice (deletion, markup, escaped text).
- `StructureChange::{PageMoved, ShapeMoved}` (non-exhaustive enum).
- Fix: the page-removal dependency scan refused any document with a
  character reference in text (LibreOffice writes `&apos;`); character and
  predefined references are now accepted, custom entities still refused.
- Unified: `add` honours placements for pages and shapes (`Start`,
  `Position`, `Before`, `After`); `remove` removes pages (never the last
  one); `move` and `swap` work on pages and on shapes within a page
  (`capabilities()` now reports move/swap supported for ODG). A move into
  another page is `unsupported_capability`.

Tests: `crates/litchi-odg/tests/producer_structure.rs` (3page/FontWork
fixtures: inserts, removal, moves both ways, joined splices, inverse move
restores `content.xml`), the updated
`package_semantics::dtd_is_refused_and_noncompact_source_requires_checked_splices`
(page insertion now succeeds on the non-compact source; a layer insertion
still needs the compact rewrite), and the ODG unified tests.

## 3. ODC writes (0282 G3)

- `set` on `/`: `title`, `subtitle`, `chartType` (`bar` or `chart:bar`;
  the twelve ODF classes), `width`, `height`; on `/axis[N]`: `name`,
  `title`, `categories`; on `/series[N]`: `values`, `label`, `chartType`,
  `axis`.
- `add` under `/`: `series` (`values` required; `label`, `chartType`,
  `axis`) and `axis` (`dimension` x/y/z required; `name`, `title`,
  `categories`), positioned by index.
- Route: a chart the leaf projects losslessly (litchi-authored) is
  rewritten through its typed `Definition` (`replace_chart`), so it stays
  projectable; a producer chart keeps attribute writes (class, size,
  series ranges/label/axis/class, categories) as the leaf's exact-span
  splices, and title text or structure there stays a typed refusal. Axis
  renames alone are always exact splices.

## 4. ODI geometry on fresh images (0282 G4)

The flat-image rewriter inserted `svg:`/`draw:`/`style:`/`text:`/`xlink:`
attributes without a declaration when the compact package had dropped the
unused one, so readback failed. The rewritten start tag now declares the
prefix of every inserted attribute the document root does not bind. Test:
`crates/litchi-odi/tests/fresh_geometry.rs`; the unified ODI test now sets
fresh geometry.

## 5. Object-format coalescing (0282 G8)

`format::odf_tree::coalesce::Run` joins consecutive sets of `File::edit`
and `batch` on ODG, ODC, ODI, ODB and the formula into one leaf `Edit` and
one commit (`write::set_groups`; every flavor's setter now stages a list of
(nodes, props) groups). A set joins only when its path is positional, every
key is coalescable (ODG page name and shape text/name/geometry; ODC chart,
axis and series keys; ODI frame and metadata; ODB query command; formula
token text), and no pending set writes the same (node, property family)
(position and size count as one geometry family). Queries, `/metadata`,
renames that change addressing (ODB tables/columns), the formula's
StarMath and anything else close the run first. Errors keep their item
index; failed batches leave the input unchanged.

- ODB leaf: scalar query setters on several queries in one transaction were
  refused ("one query per transaction"); they now stage, and
  `Patch::change()` reports the legacy single-query change only when one
  query was touched.
- Committed object handles build their tree on the first read
  (`OdfTree::from_commit`); ODB, whose tree parses the catalog fallibly,
  still builds eagerly.

## 6. Measurements

`crates/litchi-ops/examples/odf_gaps_bench.rs`, release, `taskset -c 9` on
a shared 32-core host (load 6–9 from other agents), ABBA × 12 with two
warm-ups, medians; every write scenario asserts outside the clock that the
unified output equals the leaf output byte for byte
(`results/0288/odf_gaps_bench.{txt,json}`). Batch legs: one hand-written
leaf `Edit` with N stagings and one commit vs one unified `File::edit` of
N sets.

| Format | Scenario | leaf µs | unified µs | ratio |
|---|---|---:|---:|---:|
| ODG (1,000 shapes) | batch 1 (shape text) | 13,049.6 | 13,067.0 | 1.001 |
| ODG | batch 10 | 45,492.4 | 45,676.4 | 1.004 |
| ODG | batch 100 | 370,351.1 | 370,095.1 | 0.999 |
| ODG | batch 1,000 | 3,613,078.2 | 3,617,156.2 | 1.001 |
| ODB (1,000 queries) | batch 1 (query command) | 2,269.6 | 3,057.4 | 1.347 |
| ODB | batch 10 | 9,843.7 | 10,653.3 | 1.082 |
| ODB | batch 100 | 85,996.5 | 86,941.8 | 1.011 |
| ODB | batch 1,000 | 930,517.8 | 916,063.7 | 0.984 |
| ODF formula (1,000 tokens) | batch 1 (token text) | 1,590.0 | 1,585.7 | 0.997 |
| ODF formula | batch 10 | 8,405.4 | 8,400.5 | 0.999 |
| ODF formula | batch 100 | 77,883.7 | 77,634.1 | 0.997 |
| ODF formula | batch 1,000 | 816,520.7 | 820,504.8 | 1.005 |
| ODG (`3page.odg`) | metadata set | 1,283.5 | 1,285.0 | 1.001 |
| ODG | insert page first | 1,703.3 | 1,710.0 | 1.004 |
| ODG | move page last | 1,611.8 | 1,612.9 | 1.001 |
| ODS (`calc-two-sheets`) | metadata set | 556.1 | 556.0 | 1.000 |
| ODP (`impress-basic`) | metadata set | 997.1 | 997.0 | 1.000 |
| ODC (typed chart) | title set | 149.3 | 150.7 | 1.009 |
| ODI (fresh image) | geometry set | 161.5 | 162.7 | 1.007 |

Against the PERFORMANCE BAR:

- **Batch (1k-op ≤ 2× one leaf Edit):** met for ODG (1.001), ODB (0.984)
  and the formula (1.005); every N from 1 to 1,000 is one leaf commit.
  10k-op batches were not measured: the leaves' own per-staging cost (ODG
  reparses `content.xml` per `set_shape_text`, ODB rescans per setter)
  makes them quadratic in the leaf leg itself (ODG 1k = 3.6 s).
- **Single-op write (≤ 10 %):** met for every new single-op path above.
  **Miss:** ODB single set on a 1,000-query database is 1.35× (and 1.08×
  at 10 sets): the unified handle rebuilds the object tree after the
  commit, which parses the 1,000-query catalog (~0.8 ms) the leaf leg never
  reads. The other object formats now defer that build to the first read;
  ODB does not, because its catalog parse is fallible (record 0282
  measured ODB single set 1.062 on a 6-query database). Follow-up below.
- A first run exposed a determinism bug: `PackageWriter` copied manifest
  directory entries in `HashMap` order, so two publications of the same
  edit could differ in `META-INF/manifest.xml`. Directory entries are now
  sorted in all three copy routes (`litchi-odf-common/src/core/writer.rs`);
  the byte-equality checks above pass since.

## 7. Gaps and follow-ups

Listed in `capabilities()` unless noted.

- **G-ODP-1 character formatting** (not attempted here): ODP run
  bold/size/colour/font need automatic text-style resolution and
  authoring in the scene splices. Follow-up unchanged from 0278.
- **ODB single-op tree rebuild** (not a capability gap): a fallible-free
  ODB tree (or a cached catalog on the committed package) would let ODB
  defer the build like the other formats.
- **ODG:** shapes other than text rectangles; a shape move into another
  page (`0288 ODG cross-page shape transfer`); removing a page another page
  references by `draw:page-name` (`0288 ODG reference rewriting`); layer,
  style, form and transition-style edits still use the compact-only
  whole-part rewrite.
- **ODC:** title/subtitle text, size and structure on producer charts the
  leaf cannot project losslessly (`0288 ODC lossless producer
  definitions`); data points, labels, regressions and series styles
  (`0288 ODC series detail writes`). Producer size writes need an `svg`
  prefix proof from the leaf.
- **ODI:** payload (`source`) writes; one frame per image.
- **Metadata language read-back** (not a capability gap): see §1; needs
  an owner decision on the fixed Markdown golden.
- **Metadata:** user-defined fields and statistics/dates (`0288 ODF
  user-defined and statistics metadata`); `/metadata` is set-only on the
  object formats other than ODI (`0288 ODF object metadata node`; reads go
  through `File::metadata()`).
- **Coalescing:** ODB renames, the formula's StarMath, `/metadata` and
  queries take the ordinary route.
- 0282 G5/G6/G9–G11 unchanged.

## 8. Shared files

Separated hunks only: `capabilities.rs` (object move/swap support),
`error.rs` (`caused_by_argument` for `odp`), `transaction.rs` (the ODS/ODP
`Leaf` patch types, the object coalescing run and its flush),
`format/mod.rs` (`odf_metadata`). No iWork crate and no `PROGRESS.md`
change. Tests: CLI (`litchi-cli/tests/cli.rs`, ODG producer
insert/move/swap/remove and `/metadata`), Python
(`test_write_api.py::OdfObjectTests`).

## 9. Gates (all exit 0 on the final tree)

| Command | Result |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-{odf-common,odt,odp,oth,odm,odg,odc,odi,odb,odf-formula} -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ods --lib --all-features -- -D warnings` (+ `--test metadata_family`) | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features f] --all-targets -- -D warnings`, f ∈ none, odt, ods, odp, oth, odm, odg, odc, odi, odb, odf-formula | 0 each |
| `cargo test --no-fail-fast --all-features -p litchi-{odf-common,odt,odp,oth,odm,odg,odc,odi,odb,odf-formula}` | 0 |
| `cargo test --all-features -p litchi-ods --test metadata_family` | 0 (the ODS suite needs a purged fixture; lib linted) |
| `cargo test --no-fail-fast -p litchi-ops --all-features` (lib, doc, and every integration test in five chunks to stay under the disk cap) | 0 |
| `cargo test --no-fail-fast -p litchi --all-features` | 0 |
| `cargo test --no-fail-fast -p litchi-cli` | 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --test capabilities --test registry`, `--all-features` and every feature but `encryption` | identical files; only the ODF goldens changed |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features` (the eleven ODF crates and `litchi-ops`) | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 (267 tests) |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `python3 tools/officecli_capability_delta.py --officecli <OfficeCLI>` then `--check` | 0, 0 |

The Python suite found a real defect before the final run: a `meta.xml`
with a byte-order mark got its splices three bytes early (the reader skips
the mark). The planner now works after the mark and shifts the ranges
back (unit test `plans_on_part_bytes_after_a_byte_order_mark`). The
OfficeCLI corpus (DOCX/XLSX/PPTX) is not affected by this change.
