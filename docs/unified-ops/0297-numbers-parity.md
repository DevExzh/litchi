# 0297 — Numbers to parity in the unified layer (M6)

- Status: implemented on `uo2/iwork-numbers`, cut from
  `feat/unified-ops-api` at `dbd990cad1`. Plan
  [0248](0248-coordinator-restart-plan.md) §6 Phase F (M6, Numbers).
- Date: 2026-10-15
- Owner decisions applied: **D1** (`File::Numbers` holds the leaf package)
  and **D2** (the iWork crates may be edited), 2026-10-15.
- Commits: `723056db68` (clippy lints in the leaf's test targets),
  `95b149bee0` (ADR 0034 W1), `5cc4712da9` (formula parser, sheet order
  without previews, `limits::Error`), `68e1a70181` (unified adapter, tests,
  goldens), and this record.
- Siblings: Keynote is record 0298, Pages record 0299. No shared
  `litchi-iwa-*` crate changed; `crates/litchi-ops/src/format/iwork/`
  changes are the Numbers files plus small hunks in `iwork.rs`, `error.rs`
  and `write.rs`.
- `performance_claim`: paired leaf/unified medians in §5 (release, one
  pinned core of a shared host, ABBA). No frozen-harness qualification.

## 1. ADR 0034 W1 for `litchi-numbers`

- **Owner and projection.** The immutable `Package` owns every state verb
  (`open`, `from_bytes`, `from_shared_bytes`, `*_with_options`, `new`,
  `edit_*`, `apply_*`, `save`, `write_to`). `Document` became `Workbook`
  (module `litchi_numbers::workbook`): owned through `Package::workbook()`,
  borrowed through `workbook_ref()`. As with the Pages `Document`, the
  projection keeps read-only ingress for app-authored package directories,
  which the owner refuses. `Package::object_count` and `Package::text` were
  removed (they were deprecated).
- **Limits** (`src/limits.rs`). `SourceLimits` and its builder: input
  bytes, members, entry bytes, aggregate bytes, component bytes, and the
  Numbers-specific `header_fields`. `SemanticLimits` and its builder:
  objects, sheets, tables, references, materialized cells, output text
  bytes, formula render work and depth. Both have `to_builder()`. Other
  types: `SourceLimitKind`, `SemanticLimitKind`, `LimitKind`, and
  `OpenOptions` with `source_limits_value()`/`semantic_limits_value()`.
  Setters and `build()` return the owning module's
  `limits::Error::InvalidLimit` (ADR 0034 §5). The source defaults stay at
  the hard ceilings (2 GiB input), a deviation from the shared 512 MiB:
  the focused transactions precharge their work budgets from the component
  ceiling. Before this record the profile was `ReadOptions`. It is now the
  package-private `Profile`.
- **Creation** (`src/create.rs`). `Package::new()` returns an embedded
  6,188-byte blank workbook (SHA-256 `13ec94a1…`): one sheet and one
  10 x 5 table. It was generated once from the migration host's
  `NumbersDocumentBuilder`, with the identity rewritten to `Identity::BLANK`
  and a build history of "Created by Litchi", so every call is
  byte-identical. `create::Builder` takes a required `create::Identity`
  (`generate`, `from_bytes`, `BLANK`) and names the sheet and table through
  the names transaction.
- **Errors.** `Error::UnexpectedFormat { detected }` and
  `workbook::ReadError::UnexpectedFormat { detected }` replace
  `NotNumbers`. `formula::Error::Syntax { offset }` is new.
- **Families.** The transaction families moved into focused modules with
  short names: `table::lock`, `table::merge::transaction`,
  `sheet::chart::{arrangement, data, metadata}`,
  `shape::image::adjustments`, `cell::comment`,
  `cell::comment::transaction`, and `table::dimension::transaction::Error`.
  The prefixed root aliases were deleted, not aliased. Selectors carry
  `Position(Position)`; `index()` stays as a constructor, and `as_index` was
  deleted. `SemanticLimitKind::FormulaWireBytes` became
  `FormulaCategoryBytes`.
- **Consumers migrated:** `litchi-iwa`, `litchi-iwa-structured`, the
  facade's `numbers`/`iwork` modules and fuzz targets, `litchi-ops`, and
  `pyo3-litchi`.
- **Glossary.** The profiles `numbers-minimal` and `numbers-full` ban 103
  names and register 27 complete families. Rows in `docs/API_GLOSSARY.md`
  ("W1 Numbers (record 0297)") were flipped in the W1 commit.
- **Leaf public-API check.** `tools/check_iwork_leaf_public_api.py` reports
  the same 45 findings at `dbd990cad1` and after this record. They are
  pre-existing `litchi_iwa_common::table::model` types reachable from the
  projection. W1 had added 19 findings of its own: the limit builders
  returning the package error, and the "wire" kind name. `5cc4712da9`
  removed them.

## 2. Coverage in the unified layer

| Verb | Element / key | Leaf call | Status |
|---|---|---|---|
| `create` | Numbers | `Package::new` | supported (Template flavor and metadata refused) |
| `set` | sheet `name` | `edit_names().rename_sheet` | supported |
| `set` | table `name` | `edit_names().rename_table` | supported |
| `set` | cell `value` (text, number, bool, null = clear) | `edit_table_cells` | supported |
| `set` | cell `formula` | `Expression::parse` + `Change::set_formula` | supported: local refs, ranges, functions |
| `set` | cell `numberFormat` | number/percentage display-format transactions | supported: decimal and percentage codes; `General`/null resets |
| `add` | `mergeCell` with `ref` | `edit_table_merges().merge` | supported |
| `remove` | `mergeCell` | `edit_table_merges().unmerge` | supported |
| `remove` | cell | `Change::clear` | supported (clears) |
| `move` | sheet | `edit_sheet_order` | supported |
| `move` | table into another sheet (`Placement::Into`) | `edit_table_relocation` | supported (appended) |
| `get`/`query` | `mergeCount`, `merge`, `mergeAnchor`, `mergeCell[ref]` | merge reader | supported |
| `add`/`remove` | sheet, table | — | gap (requests 4, 5) |
| `add`/`remove` | row, column | — | gap (request 6) |
| `set` | `font.bold` and other text styles | — | gap (request 7) |
| `set`/`get` | cell `comment` | — | gap (new request 19) |
| `set` | cross-table formula | — | gap (source-bound table handles) |
| `move` | table within its sheet | — | gap (relocation only appends) |
| `set` | `formula` on a created workbook | — | gap (new request 18: the blank has no formula owner) |

Gaps are typed: they are `unsupported_capability` with a follow-up, and
they appear in `capabilities()`, the registry golden and the OfficeCLI
delta.

**Batch coalescing.** Consecutive direct `value`/`formula` sets to distinct
cells in a batch or closure edit share one leaf cell transaction per table
(`format/iwork/numbers/coalesce.rs`). A repeated cell, any other key or
verb, or a query closes the run, so every operation sees the result of the
ones before it. The run resolves each table path once, compares the steps
rather than hashing the path text, and reads whether the cell is stored
from the leaf projection, not from a tree walk per cell. A 1,000-item batch
is one leaf commit (`diagnostics.commits() == 1`).

## 3. Tests and goldens

- `crates/litchi-ops/tests/numbers_unified.rs` (10 tests):
  - creation, both deterministic and writable;
  - renames, sheet moves, table relocation and the in-sheet refusal;
  - formulas; display formats and the comment refusal;
  - merges (add, read, unmerge, a Numbers-authored merge);
  - cell clears; batch coalescing, including the barrier on a repeated cell;
  - typed lifecycle gaps;
  - a corpus sweep over all 37 `test-data/iwork/numbers` files: 36 merge
    round trips, 2 sheet-move round trips, 0 refusals.
- `crud_matrix.rs`: fixtures `table-merges-native` and
  `table-relocation-native`, plus a `mergeCell` recipe. `query_corpus.rs`
  adds `mergeCell`, `table[mergeCount>0]` and `cell[mergeAnchor=true]`.
- Goldens blessed in both feature sets, with identical output:
  capabilities, registry, query digests (the Numbers lines only), the CRUD
  checklist and the fidelity report. In the fidelity report Numbers is
  40/40 no-op byte-exact, and its probe is now value-exact through a sheet
  rename.
- CLI (`crates/litchi-cli/tests/cli.rs`): `create blank.numbers` succeeds,
  and the created workbook takes a sheet rename, a value and a merge.
  Python: `create("numbers")` joins `CREATABLE`.
- `litchi-numbers`: parser unit tests and
  `a_package_without_previews_still_reorders`.

## 4. Findings outside the scope

- **Release-only `litchi-iwa` builder failure.** In a release build,
  `NumbersEditor::create()` and `NumbersDocumentBuilder::build()` fail with
  `Bundle("IWA package entry not found: Index/Document.iwa")`; debug builds
  succeed. 0297 changed only renames in `litchi-iwa`, so this is likely
  pre-existing. It does not affect `Package::new()`, which embeds bytes. The
  bench grid was generated in a debug build.
- A formula set on `table-relocation-native.numbers` (sheet 1, table 1) is
  refused by the leaf with `InvalidSource { path: Package }`. On
  `basic.numbers` the same set succeeds. The bench uses `basic.numbers`.
- A leaf merge draws fresh object identifiers, so two identical merges
  differ byte for byte. The bench compares merged regions.

## 5. Performance

`crates/litchi-ops/examples/numbers_unified_bench.rs`, built with
`cargo build --release -p litchi-ops --no-default-features --features
numbers --example numbers_unified_bench` from `68e1a70181`. The binary's
SHA-256 is `beed17bb6e86861b8397f9bcabe3c627ef645ce99ec41a846ede7d891b4db01e`.
It ran as `LITCHI_NUMBERS_BENCH_SAMPLES=7 taskset -c 7 …` on a shared
32-core host: 7 ABBA rounds, 14 samples per leg, two runs. The raw samples
are in [`results/0297/numbers-bench.jsonl`](results/0297/numbers-bench.jsonl)
(6.8 KB). Write outputs were compared before timing: byte-identical, and
merges region-identical.

| Scenario | Fixture | Leaf median | Unified median | Ratio (run 1 / run 2) |
|---|---|---:|---:|---:|
| `open_get` | bench grid | 100.1 µs | 116.8 µs | 1.167 / 1.175 |
| `set` (one value) | bench grid | 369.3 µs | 372.4 µs | 1.008 / 1.014 |
| `batch_1000` | bench grid | 997.5 µs | 1,472.1 µs | 1.476 / 1.439 |
| `formula` | basic | 2,743.9 µs | 2,760.1 µs | 1.006 / 1.003 |
| `rename_sheet` | bench grid | 197.1 µs | 199.2 µs | 1.011 / 1.008 |
| `merge_add` | bench grid | 184.6 µs | 191.0 µs | 1.035 / 1.031 |
| `sheet_move` | relocation | 2,216.4 µs | 2,213.6 µs | 0.999 / 1.000 |

The batch path was profiled with callgrind (`--toggle-collect` on the
out-of-line legs, line tables only). The first coalesced version resolved
each cell through the object tree. A cell that was not stored took a failed
full resolve and then a second resolve of its table, which made the batch
2.07x the leaf (unified 133.2 M Ir). Caching the table per run with a
hashed path-text key brought it to 1.64x (108.1 M Ir). Comparing the steps
instead brought it to 1.44–1.48x (97.4 M Ir; the leaf leg is 67.3 M Ir).

The residual is about 5 k Ir per item. Half of it is parsing the path text
of each batch item (`IntoTarget`), which all formats share. The other half
is the transaction's per-item planning. The commit itself costs the same in
both legs. `open_get` pays about 17 µs for the unified open, which builds
the object tree. The remaining single operations are within 4 % of the leaf.

Target directory: 13–16 GB during the work; pruned to examples and test
binaries between suites.

## 6. Verification

The gates ran on the final tree (`68e1a70181`, with the record added):

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-numbers -p litchi-ops -p litchi -p litchi-py -p litchi-cli -p litchi-iwa-structured --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features numbers / pages / keynote / pages,numbers / keynote,numbers] --all-targets -- -D warnings` | 0 (each) |
| `cargo clippy -p litchi-iwa --lib --all-features -- -D warnings` | 0 |
| `cargo test -p litchi-numbers --all-features` | 0 |
| `cargo test -p litchi-ops --all-features --no-fail-fast` | 0 after the re-bless (`corpus_fidelity` and `query_corpus_differential` take the new fixture) |
| `cargo test -p litchi --all-features` | 0 |
| `cargo test -p litchi-cli` | 0 |
| `cargo test -p litchi-iwa --all-features --lib --test 'numbers*'` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 (269 tests) |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers -p litchi-ops` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile numbers-minimal / numbers-full / ops-iwork` | 0 (each) |
| `python3 tools/check_iwork_leaf_public_api.py` | 1, the same 45 findings as at `dbd990cad1` |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test capabilities --test registry`, then the same with the explicit feature list | 0, 0 (identical goldens) |
| `python3 tools/officecli_capability_delta.py --officecli …/3rdparty/OfficeCLI` | 0 (Numbers adds `mergeCell`) |

The OfficeCLI corpus was not run: it has no Numbers documents, and no
Office-format behaviour changed. `litchi-iwa`'s test targets have
pre-existing Keynote clippy lints, and its `*_pie_chart`/`*_donut_chart`
examples have pre-existing compile errors. Both are outside this record.
The test `test_iwa_package_edge_is_archive_owned_and_host_edge_cannot_return`
in `tools/test_check_crate_boundaries.py` fails at the base too
(281 != 270).

## 7. Follow-ups

1. Sheet and table add/remove, and row/column insert/delete, in the leaf
   (requests 4–6). The host editor has them, but porting them means the
   leaf must own the drawable, calculation-engine and view-state graphs.
2. A formula owner in the blank workbook (request 18), so that formulas
   work on `create(Numbers)`.
3. Comment writes on Numbers-authored files (request 19).
4. Cell text styles (request 7); in-sheet table order; cross-table
   formulas.
5. The release-only `NumbersDocumentBuilder` failure in `litchi-iwa` (§4).
6. The cost of parsing path text per batch item, which all formats share.
