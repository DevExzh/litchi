# 0282 — The remaining ODF formats in `File`: OTH, ODM, ODG, ODC, ODI, ODB and ODF formula

- Status: implemented on `uo2/odf-leaves`, cut from `feat/unified-ops-api`
  at `80e24f8745`. Plan [0248](0248-coordinator-restart-plan.md) §6 Phase F,
  step 23 part 2. Builds on the W3 split of [0268](0268-odf-w3-package-split.md)
  (follow-up G1) and the ODT/ODS/ODP adapters of records 0276–0278.
- Date: 2026-10-14
- Owned: `crates/litchi-{oth,odm,odg,odc,odi,odb,odf-formula}`, the new
  `crates/litchi-ops/src/format/odf_tree*`, small separated hunks in the
  shared ops files, plus two strictly additive reuse hooks (§4).
- `performance_claim`: paired leaf/unified medians in §6 (release, one
  pinned core of a shared host, ABBA × 12, three runs). No frozen-harness
  qualification.

Before this record `litchi::open` detected these seven formats and refused
them with `unsupported_format` (`not_enabled`); their capability reports
were whole-format refusals (`M6`). Every claimed ODF format now opens into a
`File` variant and is readable and writable through the unified verbs.

## 1. Leaf W3 (0268 G1)

Each of the seven crates now has one immutable `Package` owner with every
state verb and a read-only projection (commit `20098fc08e`):

| Crate | Projection | Primary family | Removed (banned) |
|---|---|---|---|
| `litchi-oth` | `Document` | `litchi_oth` | `Template`, `Commit::{template, into_template, changed}`, `Builder::build() -> Vec<u8>` |
| `litchi-odm` | `Document` | `litchi_odm::transaction` | `Master`, `apply_*_patch`, `Commit::{snapshot, into_snapshot}` (all three families) |
| `litchi-odg` | `Drawing` | `litchi_odg::edit` | the 37 root `Package*` aliases, `Snapshot`, `Transaction`, `Drawing` state verbs |
| `litchi-odc` | `Chart` | `litchi_odc` | `Chart::{open*, from_bytes*, from_definition*, edit, into_bytes}`, `Commit::{chart, into_chart, changed}` |
| `litchi-odi` | `Image` | `litchi_odi` | `Image::{open*, from_bytes*, edit, into_bytes, as_bytes}`, `Commit::{image, into_image, changed, semantic_patch}` |
| `litchi-odb` | `Database` | `litchi_odb` | `Database::{open*, from_bytes*, edit*, into_bytes}`, `Commit::{database, into_database, changed, protection_transition}` |
| `litchi-odf-formula` | `Formula` | `litchi_odf_formula` | `Formula::{create*, open*, from_*, edit, set_root, apply_patch, to_bytes, into_bytes, save*}`, root `Limits`/`LimitError`, `Builder::build_with_limits` |

- **State verbs on `Package`:** `new()` (deterministic, through the
  crate's `Builder`), `open`, `open_with_options`, `from_bytes(&[u8])`,
  `from_bytes_with_options`, `from_reader(_with_options)`,
  `from_prepared_package` (refuses another family), password routes where
  the crate had them (ODG, ODI, ODB, ODM), `edit`, `apply(&Patch)`, atomic
  `save`/`save_with_durability` (`litchi_odf_common::publish::save_bytes`),
  `to_bytes`, `write_to`, `as_bytes`, and the projection accessor.
- **Commit shape:** `Commit::{package, patch, diagnostics, into_package,
  into_parts}`, `Diagnostics::changed()`, `Patch::is_noop()`; public
  `Patch::apply` is gone (`Package::apply`). Every family exports
  `{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` (`Error` and
  `LimitKind` are `litchi_odf_common`'s); exact no-op commits return the
  source bytes. `Builder::build()` returns `Package`.
- **Limits:** ingress takes `litchi_odf_common::OpenOptions`. Source limits
  and the declared `content.xml`/`styles.xml` byte limits are applied by the
  additive `core::family::Package::*_with_options` constructors (§4);
  `SemanticLimitKind::Units` is checked against pages (ODG), top-level
  blocks (OTH), master-body items (ODM), plot-area series (ODC) and frames
  (ODI). ODB and the formula count no units (gap G9).
- **CRUD evidence:** a `tests/package_owner.rs` per crate (8–10 tests each):
  deterministic `new()`, `from_bytes` round trip, exact no-op commit, real
  edit plus `into_parts` + `Package::apply` reproducing the bytes, atomic
  replacement without a staging file, refused save through a symlink, and
  typed `Error::Limit` for input bytes / content bytes / units.
- **Glossary:** seven new profiles in `tools/api_glossary.json` (`oth`,
  `odm`, `odg`, `odc`, `odi`, `odb`, `odf-formula`; 213 bans, 7 families);
  the `docs/API_GLOSSARY.md` owner, builder, commit, patch, ingress, limits
  and save rows flip for them, with a new section. The `ops-odf` profile now
  also builds the seven new `litchi-ops` features.
- **Consumers:** `tools/native-resave` follows the ODB, ODG and formula
  owners. The frozen probe `docs/performance/probes/0502-odg-open` keeps the
  old API (not built by any gate).

## 2. OTH and ODM: the ODT adapter, reused

An OpenDocument HTML template (`text-web`) and a master document
(`text-master`) have the ODT body model. They open as `File::Oth` and
`File::Odm`, which hold the ODT handle (`format::odt::Odt`), now carrying
its package's `FileFormat`. Everything of record 0276 applies unchanged:
paths, queries, views, set/add/remove/move/swap, coalesced batches,
`/metadata`, atomic save and exact no-ops. Splices only touch
`content.xml`/`meta.xml`, so the package keeps its own MIME type (tested
against the bytes). Linked sections of a master document read as sections;
their links are kept, not followed.

- Nodes and errors report `Oth`/`Odm` (`Odt::retag` at the handle
  boundary); query plans compile through the ODT spec table.
- `create(Oth|Odm)` starts from the leaf's `Package::new()` and adopts it
  into the ODT engine; create-time metadata is the ODT set.
- Features: `oth = ["odt", "dep:litchi-oth"]`, `odm = ["odt",
  "dep:litchi-odm"]` (decision: the variants exist whenever `odt` is on, so
  the shared dispatch arms are one or-pattern; they are only constructed
  when their feature is on).
- Tests: `tests/oth_unified.rs`, `tests/odm_unified.rs`; the ODM corpus
  (3 files, two copied into `test-data` with provenance) runs the 34 ODT
  query selectors.

## 3. ODG, ODC, ODI, ODB and ODF formula: one object-tree adapter

`format::odf_tree` builds, at open, an arena of the elements each leaf
projection exposes, with properties aligned to a per-format spec table and
the leaf positions writes address. One generic `Nav<F: Flavor>` implements
path resolution, queries and node views for all five; `File::{Odg, Odc,
Odi, Odb, Odf}` hold the same `OdfTree` handle; `Node::Object` is its node.

| Format | Elements (paths) | `set` | `add` | `remove` |
|---|---|---|---|---|
| ODG | `/`, `/page[N|NAME]`, `…/shape[N]`, `…/layer[N]` | page `name`; shape `text`, `name`, `x`, `y`, `width`, `height` | page under `/`; text rectangle under a page (appended) | shapes |
| ODC | `/`, `/axis[N|NAME]`, `/series[N]`, `/legend` | axis `name` | — | axes, series |
| ODI | `/`, `/frame[N]`, `/metadata` | frame `name`, `title`, `description`, `x`, `y`, `width`, `height`, `anchorType`, `zIndex`, `layer`, `style`; metadata `title`, `author`, `subject`, `description`, `keywords` | — | — |
| ODB | `/`, `/table[N|NAME]`, `…/column[N|NAME]`, `/query[N|NAME]` | table and column `name`; query `command` | query under `/` | tables, columns, queries |
| ODF formula | `/`, `/element[N]` nested | token `text`; root `starmath` | — | elements |

- Views: text, annotated, outline and stats (the shapes of the other
  tier-3 adapters).
- Writes stage the leaf's own `Edit` operations and commit once per verb;
  the committed package becomes a new handle with a fresh tree; unchanged
  values are exact no-ops. Leaf staging refusals are `invalid_value`
  (`unsupported_capability` when the leaf says unsupported); a candidate the
  leaf refuses at commit is the edit's fault, not `corrupt_file`.
- `move`/`swap` (and `add`/`remove` where the leaf has none) are
  `unsupported_capability` with follow-up `0282 ODF object-format
  structure`, reported identically by `capabilities()`.
- `create` for all five; a fresh image keeps an empty `meta.xml` so its
  `/metadata` is settable.
- Features: `odg`, `odc`, `odi`, `odb`, `odf-formula`, each turning on the
  internal umbrella `odf-objects`. Facade (`litchi`) features of the same
  names; the `odf` aggregate and the CLI enable all seven.

## 4. Additive reuse hooks

- `litchi_odf_common::core::family::Package::{open_with_options,
  open_with_password_and_options, from_reader_with_options,
  from_bytes_with_options, from_bytes_with_password_and_options,
  from_owned_package_with_options}` — the `OpenOptions` ingress the six
  family-based leaves share.
- `litchi_odt::Package::from_prepared_text_package` — adopts a prepared
  ODT, OTH or ODM package through the ODT text model, reusing the
  detection index. Nothing else in `litchi-odt`, `litchi-ods` or
  `litchi-odp` changed.

## 5. Goldens and generated tables

- Query-corpus differential: ODM (34 selectors, 3 files), ODG (11, 9
  files), ODC (5, 2), ODB (7, 6) and ODF formula (5, 4) digests, blessed
  from the release build of this change; every existing key is
  byte-identical (only additions).
- Capability and registry goldens `oth`, `odm`, `odg`, `odc`, `odi`,
  `odb`, `odf` (`.json`), blessed with `--all-features` and with every
  feature but `encryption`: both builds wrote identical files and no other
  golden changed.
- `officecli-capability-delta.md` regenerated (seven new sections).
- CLI: the read table gained ODM/ODG/ODC/ODB/formula fixtures, the creatable
  loop all seven formats, and an ODG add/set/remove round trip. Python:
  the seven formats are creatable, ODM opens with `Document.open`, the
  object formats open as `File`, and the remaining detected-not-opened rows
  are the four object formats (no family opener reads them).

## 6. Measurements

`crates/litchi-ops/examples/odf_leaves_bench.rs`, release, `taskset -c 11`
on a shared 32-core host (other agents were compiling), per-call means
over 10–50 calls per sample, ABBA × 12, medians of three runs
(`results/0282/odf_leaves_bench.{txt,json}`). For OTH and ODM a third leg,
`engine`, is the same operation written by hand against `litchi_odt`, the
engine the unified layer drives.

| Format | Scenario | leaf µs | engine µs | unified µs | unified/leaf | unified/engine |
|---|---|---:|---:|---:|---:|---:|
| OTH (1,000 ¶) | open + first get | 446.7 | 125.3 | 126.1 | 0.28 | 1.006 |
| OTH | single set (¶ 500 text) | 723.1 | 331.9 | 332.1 | 0.46 | 1.001 |
| ODM (`master.odm`) | open + first get | 1432.6 | 622.0 | 617.5 | 0.43 | 0.993 |
| ODM | single set (¶ 1 text) | — | 2168.1 | 2173.8 | — | 1.003 |
| ODG (46 shapes) | open + first get | 2242.8 | | 2298.1 | 1.025 | |
| ODG | single set (shape name) | 10077.5 | | 10082.2 | 1.000 | |
| ODC | open + first get | 76.6 | | 78.0 | 1.017 | |
| ODC | single set (axis name) | 205.4 | | 208.4 | 1.015 | |
| ODI | open + first get | 56.9 | | 57.9 | 1.018 | |
| ODI | single set (frame name) | 143.4 | | 144.8 | 1.010 | |
| ODB | open + first get | 64.7 | | 66.0 | 1.019 | |
| ODB | single set (query command) | 291.6 | | 309.7 | 1.062 | |
| ODF formula | open + first get | 66.7 | | 78.3 | 1.174 | |
| ODF formula | single set (token text) | 309.7 | | 323.2 | 1.043 | |

Against the PERFORMANCE BAR:

- **Open** (≤ 5 % or ≤ 20 µs): met everywhere. The formula's +17 % is
  11.6 µs (the object tree is built eagerly for every MathML element);
  ODG +2.5 % (55 µs, 2.5 %).
- **Single-op write** (≤ 10 %): met everywhere (worst ODB +6.2 %, which
  includes rebuilding the object tree of the committed package).
- **OTH and ODM are faster than their own leaves** because the unified
  layer reads and writes through the ODT engine; against that engine the
  overhead is within ±1 %. The OTH/ODM leaf columns are the leaves' own
  APIs (OTH projects every structure at open; ODM has no paragraph text
  edit, so its leaf `set` leg is the engine).
- Batches were not measured for the object formats (each verb is one leaf
  commit; no coalescing yet, gap G8). OTH/ODM batches are the ODT
  coalescer of record 0276.

## 7. Gaps and follow-ups

Listed in `capabilities()` (operation gaps) unless noted.

- **G1 ODG structure on producer drawings.** The leaf's compact-XML audit
  refuses page/shape insertion into LibreOffice-authored `content.xml`
  (newline after the XML declaration), as ODP's was before 0278; inserts
  work on fresh drawings. Page removal and reordering have no leaf API;
  only text rectangles are added. Some producer shapes refuse text writes
  ("not one losslessly replaceable XML span").
- **G2 ODG metadata** is read-only (the leaf has no metadata edit).
- **G3 ODC** title/subtitle/type/size and series writes need a whole typed
  definition replacement; adding series/axes needs typed range values;
  structural removals on producer charts outside the leaf's lossless typed
  surface are refused (`unsupported_capability`).
- **G4 ODI geometry on fresh images.** A frame without geometry in a
  compact package whose `content.xml` dropped the unused `svg` namespace
  declaration fails the leaf's readback when geometry is inserted
  (`invalid_value`); payload (`source`) writes are a gap. Frames and
  metadata cannot be added or removed (one frame per image).
- **G5 ODB** typed table/column creation and query rename; data rows and
  forms/reports are not modelled.
- **G6 Formula** element insertion (typed MathML values) and layout-element
  text.
- **G7 Metadata families.** None of the seven leaves adopted the ADR 0034
  §7 `metadata` family; ODI/OTH/ODM metadata writes stay in the primary
  `Edit` (owner note 2026-10-14: preferred where built — follow-up).
- **G8 Object-format batches** commit per verb; a coalescing run like
  ODT's would cut 1k-op batches (not measured).
- **G9 Limits.** ODB and the formula count no semantic units; ODC's kept
  chart `Limits` are selectable only through `Builder::with_limits`; the
  formula's MathML parser limits at ingress are the defaults.
- **G10 Flat owners** (`.fodg`, `.fodt` templates, flat chart/image/formula)
  stay out of `File` (flat ODG detects and is refused with
  `unsupported_capability`, as flat ODT is).
- **G11 Security vocabulary** (ADR 0034 §9) unchanged in the seven leaves.
- **0268 G2/G6** (public mutable writers, 512 MiB vs 2 GiB default) were not
  in reach and remain open.
- **Query/find/merge.** `Verb::Replace` and `Verb::Merge` are refused for
  all seven (records 0273/0274 follow-ups).

## 8. Shared files

Separated hunks only: `file.rs`, `read.rs`, `write.rs`, `transaction.rs`,
`save.rs`, `create.rs`, `capabilities.rs`, `open.rs`, `open/{bytes,path}.rs`,
`node.rs`, `node/view.rs`, `query/{matches,plan}.rs`, `view/{json,document}.rs`,
`tolerance.rs`, `error.rs`, `lib.rs`, `locked.rs`, `detect/*` (cfg lists),
`path/vocabulary.rs` (five vocabularies), `format/mod.rs`,
`format/odt{,/node,/write}.rs` (format tag), `Cargo.toml`. No iWork crate
and no `PROGRESS.md` change.

## 9. Gates (all exit 0)

| Command | Result |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-{oth,odm,odg,odc,odi,odb,odf-formula,odf-common,odt,odf} -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --features <f> --all-targets -- -D warnings` for none, `odt`, `ods`, `odp`, `oth`, `odm`, `odg`, `odc`, `odi`, `odb`, `odf-formula` | 0 each |
| `cargo test --no-fail-fast --all-features -p litchi-{oth,odm,odg,odc,odi,odb,odf-formula,odf-common,odt,odf}` | 0 |
| `cargo test --no-fail-fast -p litchi-ops --all-features` | 0 (goldens and query-corpus differential included) |
| `cargo test --no-fail-fast -p litchi --all-features` | 0 |
| `cargo test --no-fail-fast -p litchi-cli` | 0 (21 CLI tests) |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --test capabilities --test registry`, `--all-features` and every feature but `encryption` | identical goldens, only the seven new files |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features` (same crates plus `litchi-ops`) | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 (33 profiles, 1,410 bans, 46 families) |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 (261 tests) |
| `python3 tools/officecli_capability_delta.py --officecli <OfficeCLI>` then `--check` | 0, 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `cd tools/native-resave && cargo check --all-targets --all-features` | 0 |
