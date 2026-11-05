# 0272 — The perf-baseline harness test suite is green on trunk again

Date: 2026-10-13. Branch `uo2/harness-green`, cut from trunk
`feat/unified-ops-api` at `fd49666493` (after the merges of records 0264 §4,
0266, 0267, 0268, 0269 and 0270). Scope: `tools/perf-baseline`,
`tools/native-resave` (lockfile only), the CRUD coverage index and its tests,
and `.github/workflows/perf-baseline.yml` with its policy test. No product
crate changed.

## 1. What was failing

The coordinator's merge gate ran `cargo check --all-targets` in
`tools/perf-baseline`, which compiles tests but does not run them. On
`fd49666493`, `cargo test --all-targets --no-fail-fast` in
`tools/perf-baseline` fails 3 of 595 library tests (591 pass, 1 ignored).
Every other target passes. Three more problems would fail CI or the agents'
clippy gate. The run above hid row 4 because it did not pass `--locked`:

| # | Failure | Root cause | Fix |
|---|---|---|---|
| 1 | `producer_shape::selected_route_tests::{producer_medium_read_reports_the_first_selected_scanner_refusal_per_sheet, producer_medium_read_marker_removal_controls_reveal_downstream_refusals}`: "producer-complete-first-child-sheetViews disagreed for sheet 1: left Styles, right UnsupportedStructure" | An intended change to the leaf scanner, not a regression (§2) | Test expectations re-derived for the new admission rules (§2) |
| 2 | `tests::selectable_case_count_matches_current_enumeration`: 633 ≠ 605 | Record 0220 (`ba7755e4a0`) added 28 routes without updating the constant | Constant set to 633 after an audit (§3) |
| 3 | `cargo fmt --check` on the harness (run by the perf-baseline smoke job): 8 files unformatted | `8f50565cd8` (pptx `OpenOptions` ingress) rewrote harness call sites without formatting them | Formatting-only commit |
| 4 | `cargo test --locked` (CI) refuses to run: `tools/perf-baseline/Cargo.lock` and `tools/native-resave/Cargo.lock` are stale | `5ea005f78d` (record 0268) added `thiserror` to `litchi-odf-common`; `5ee73c18d0` added `caseless` and `unicode-normalization` to `litchi-xls`. Neither change updated the standalone harness lockfiles | Both lockfiles updated (dependency edges only; no new crate versions) |
| 5 | `cargo clippy --all-targets -- -D warnings` on the harness: 62 findings (58 `needless_borrow`, 4 `useless_conversion`) | Leaf constructors began taking owned or generic inputs in the ADR 0034 waves (e.g. ODP `from_bytes`, metadata accessors). No CI job lints the standalone harness | Mechanical `clippy --fix` plus rustfmt; smoke now lints the harness (§5) |

Item 4 of the task, the record-0220 routes, is §4. The CI hook is §5.

`tools/native-resave`: `cargo test --all-targets` passes. It has no tests;
it only builds. It failed only under `--locked` (row 4).

## 2. Producer selected-scanner controls (failure 1)

**Was the scanner change intended?** Yes. `5fe77e0bb6` ("feat(xlsx): add
retained grid reads and typed workbook writes", M1 L-XLSX-R1) documents it in
`crates/litchi-xlsx/src/raw/worksheet/selected.rs` (module docs "Inert
worksheet children", `Frame::{SheetFormat, Inert}`, `Scanner::new`). It also
adds tests for it in `raw/worksheet/tests.rs`. A **read** scan (a coordinate
or a rectangle) now admits the following children, which the materialized
parser does not interpret for cells:

- a `sheetViews` or `pageMargins` subtree, where only merge markup is checked;
- a `sheetFormatPr` that parser would read without a repair or a refusal;
- the row `spans` hint.

Anything irregular in these children still settles a not-eligible verdict.
**Target** (edit-verifier) scans keep the old slice, so commit readback is
unchanged. This is a deliberate widening of the fast path, guarded by its own
leaf tests, so nothing was changed in `litchi-xlsx`.

**The new chain**, derived from the scanner and confirmed by the test run.
The generated producer worksheet is `dimension`, `sheetViews`,
`sheetFormatPr defaultRowHeight="15" x14ac:dyDescent`, `cols`, then
`sheetData` with `spans` on every row, then `pageMargins`. Sheet 3, the only
one with a relationship, also has `pageSetup r:id`.

| Control | Before `5fe77e0bb6` | Now |
|---|---|---|
| complete producer | `UnsupportedStructure` (sheetViews) | `Styles` (cols) |
| − sheetViews | `UnsupportedStructure` (sheetFormatPr) | `Styles` |
| − sheetViews, sheetFormatPr | `Styles` (cols) | `Styles` |
| − cols only *(new)* | — | sheets 1, 2, 4 eligible; sheet 3 `UnsupportedStructure` (pageSetup) |
| − cols, defaultRowHeight removed *(new)* | — | `UnsupportedStructure` on every sheet (irregular sheetFormatPr) |
| − cols, pageSetup *(new)* | — | all four eligible: 4,096 / 3,686 / 410 |
| − sheetViews, format, cols | `UnsupportedStructure` (row spans) | same split as "− cols only" |
| − … and row spans | `UnsupportedStructure` (pageMargins) | same split |
| − … through pageMargins | sheet 3 refuses, others eligible | same split |
| − everything, incl. pageSetup | all eligible | all eligible |

**What the test still proves, and what it now adds:**

- The "marker removal reveals downstream refusals" chain is still exercised
  end to end. Removing `cols` reveals the `pageSetup` refusal on sheet 3
  only, and removing `pageSetup` makes every sheet eligible with the reviewed
  aggregate.
- The admitted markers never decide the verdict. Stripping them stage by
  stage gives the same split, and the **selected payloads are compared
  exactly** (debug rendering of every record and the dependency metadata)
  between the controls with and without the admitted markers.
- The admission is checked, not blanket. A `sheetFormatPr` without
  `defaultRowHeight` is one the materialized parser would repair, and it
  still defers every sheet on structure.
- The control names now describe the observed outcome.

The archive identity (22,739 bytes, SHA-256 `55e0901c…`) and the
4 × 32 × 32 counts are unchanged.

## 3. Selectable case count (failure 2)

`git log -S` on each name shows that the 28 new `Case` variants
(605 → 633) were all added by `ba7755e4a0` ("perf(ops): publish bounded PPT
text chunks together", record 0220). They are
`PptTextAggregate{Eager,Positional}{Native,Unified}{1,10,100,256,1000,4096,10000}`.
Each one is registered as an intentional route:

- `PPT_TEXT_BATCH_CATALOG` (28 entries, twin-paired);
- `parse_case`;
- the usage text;
- the CRUD coverage index, which lists them as excluded selectors.

`ppt_text_batch_routes_are_named_paired_and_opt_in` asserts that none is in
`Case::DEFAULT`. No other variant was added or removed since `af7e6ef49a`
(M0a S9). The constant is now 633, and its comment cites both records.

The test's other invariants still mean what they say:

- `Case::DEFAULT.len() == 41` is the default matrix. The array is typed
  `[Self; 41]`, so it is compile-checked as well.
- The four provider-axis cases stay opt-in.

## 4. Record-0220 routes: renamed `_aggregate_v2` → `_aggregate_v3`

**Decision: option (b), following the harness convention.** Record 0220
moved the `*_v1` selectors to `*_aggregate_v2` when the exact-byte output
changed above 256 targets. It retired `_v1` from the active harness, left
its frozen evidence unchanged, and said that "renaming the active routes
keeps the old exact-byte oracle intact rather than weakening it to text
equivalence". It cited the earlier `chunk_publication_v2` precedent of
record 0158.

Record 0264 §4 raised `litchi_ppt::package::edit::MAX_OPERATIONS` from 4096
to 16384. The routes derive their expected commits from that bound
(`count.div_ceil(PPT_TEXT_BATCH_NATIVE_OPERATION_LIMIT)`), so the 10k routes
now qualify one publication instead of three, and their output bytes differ
from `results/0220-ppt-aggregate-batch.json`. Record 0264 left the rename to
the coordinator. All 28 names are now `ppt_text_batch_{eager,positional}_{native,unified}_{count}_aggregate_v3`.
The whole family moved, as in 0220, where the counts whose bytes matched
were renamed too.

Changes:

- `Case::name`, `parse_case`, the usage text, the CRUD coverage index
  (`selector_names`, the excluded list, `selector_names_sha256`) and
  `tools/test_crud_coverage_index.py` (its expected digest).
- The `timed_calls` descriptions now say "16384-operation". They had said
  4096, which was stale since 0264 §4.
- `ppt_text_batch_routes_are_named_paired_and_opt_in` now asserts:
  - the `_aggregate_v3` suffix;
  - that both `_v1` and `_aggregate_v2` names no longer parse;
  - `PPT_TEXT_BATCH_NATIVE_OPERATION_LIMIT == 16_384`.

  The last check means a future bound change fails this test and forces the
  next rename, rather than silently changing frozen evidence.
- The frozen 0220 JSON and the 0264 record are unchanged. The case count is
  unchanged (633).

**Identity check against the frozen v2 evidence.** Each of the 28 v3 routes
ran once (debug build, one sample, no warmup; qualification only, no timing
claim). Evidence: [`results/0272-ppt-aggregate-v3-identity.json`](results/0272-ppt-aggregate-v3-identity.json).

| Counts | Source SHA-256 vs 0220 | Output SHA-256 vs 0220 | Publications (native/unified), v2 → v3 |
|---|---|---|---|
| 1, 10, 100, 256, 1000, 4096 (24 routes) | equal | **equal** | 1/1 → 1/1 |
| 10000 (4 routes) | equal | differs | 3/3 → 1/1 |

So v3 is byte-identical to v2 through 4096 targets. Those v2 timings stay
comparable as outputs, though not as timings across builds. Only the 10k
routes changed meaning. Their new timing is the 0264 §4 measurement
(`results/0264-ppt-bound.json`).

## 5. CI hook

**What CI already does.** The harness tests are already in CI twice:

- `rust-ci.yml` → job "non-iWork release gate" → step "Test the standalone
  performance harness" runs `python3 tools/non_iwork_gate.py harness-tests`.
  That is `cargo test --manifest-path tools/perf-baseline/Cargo.toml --lib
  --tests --no-fail-fast`, plus the serial allocator binary. Trigger:
  `**/*.rs` and others, on pushes to `main` and PRs into `main`.
- `perf-baseline.yml` → job `smoke` → step "Check standalone harness" runs
  `cargo fmt -- --check`, `cargo check --locked` and `cargo test --locked`.
  The same events, behind a path filter.

`--lib --tests` and the default `cargo test` run every harness test
function: no example or bench target in the harness defines a test. So
`--all-targets` would only add compilation of examples and benches, and
`cargo check --all-targets` already covers that.

**Why the failures accumulated anyway:**

- the trunk `feat/unified-ops-api` is a local branch (there is no
  `origin/feat/unified-ops-api`), so no workflow ever ran on it;
- the coordinator's local gate used `cargo check`, which compiles tests but
  does not run them;
- the local run omitted `--locked`, so stale lockfiles (row 4) were rewritten
  silently.

All four failures would have shown up only at the eventual PR into `main`.

**Measurement** (32 logical CPUs, shared with other agents, debug,
`CARGO_INCREMENTAL=0`, debug info off):

| Command | Wall |
|---|---:|
| `cargo test --all-targets --no-fail-fast` from an empty target dir (build + run) | 673 s |
| of which the library test binary | 568.6 s |
| the other targets (bins, integration tests) | < 1 min |

The `non_iwork_gate.py` comment records 581 s for the default suite at
default parallelism, and 1,870 s serialized. The suite fits the existing CI
jobs: the smoke job has a 120-minute timeout, and the gate job runs it
already. No new step is needed.

**Changes in this record:**

1. **Path filters.** `crates/litchi-ops/**` and `crates/litchi-xldm/**` were
   missing from the `perf-baseline.yml` push and pull-request filters. Both
   are in the harness's resolved dependency graph (`litchi-ops` is a direct
   dependency), so a change touching only them would not have triggered the
   smoke job's `cargo test`. Both are now added.
2. **Policy test.** `tools/test_perf_workflow_policy.py` gains
   `test_harness_path_dependency_closure_is_in_push_and_pull_request_path_triggers`.
   It reads the harness's resolved local-crate closure from
   `cargo metadata --locked` (features and optional dependencies resolved by
   Cargo; 33 directories today) and requires each `<dir>/**` in both
   filters. A new local dependency of the harness now fails the smoke job
   until the filter is extended. The test is skipped where `cargo` is absent.
   The smoke job sets up Rust before it runs this test.
3. **Harness lint in CI.** The smoke job's "Check standalone harness" step
   now runs `cargo clippy --locked --all-targets -- -D warnings` on the
   harness, between `check` and `test`. Its toolchain step installs
   `clippy, rustfmt`. Row 5 drifted because nothing linted the harness.
   `test_smoke_lints_formats_and_tests_the_standalone_harness` pins the fmt,
   clippy and `--locked` test commands and the clippy component.
4. **Route-identity binding** (§4). The route test pins the native bound.

**For the coordinator's merge gate (follow-up, not a code change):** replace
the harness `cargo check --all-targets` with
`python3 tools/non_iwork_gate.py harness-tests`, or with
`cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml --lib --tests --no-fail-fast`.
That costs about 10 minutes. `--locked` is what catches row 4. If the
program ever pushes the trunk, adding it to the `rust-ci.yml` branch
filters would make GitHub run the same gate.

## 6. Gates

All commands ran in the branch worktree with the program build environment.
Exit codes:

| Gate | Exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo fmt --manifest-path tools/perf-baseline/Cargo.toml -- --check` | 0 |
| `cargo clippy --locked --manifest-path tools/perf-baseline/Cargo.toml --all-targets -- -D warnings` (after row 5) | 0 |
| `cargo clippy -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cd tools/perf-baseline && cargo test --locked --all-targets --no-fail-fast` (644 pass, 1 ignored; 652 s) | 0 |
| `cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml --features allocator-metrics --bin docx_bounded_tail_append_compare -- --test-threads=1` | 0 |
| `cd tools/native-resave && cargo test --locked --all-targets` | 0 |
| `cargo test -p litchi --all-features` | 0 |
| `cargo test -p litchi-ops --all-features --no-fail-fast` | **101**: one pre-existing trunk failure, see below |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 |
| `python3 -m unittest tools.test_perf_baseline_source_policy tools.test_crud_coverage_index tools.test_perf_workflow_policy tools.test_non_iwork_gate tools.test_no_3rdparty_references` | 0 |
| `python3 tools/validate_crud_coverage_index.py …` (the CI arguments) | 0 |

**Pre-existing `litchi-ops` failure, not touched here (no product changes in
this record).** `tests/charts.rs::xlsx_chart_add_get_set_remove_round_trip`
fails on trunk `fd49666493` with `unsupported_element: this element kind
cannot be added to an XLSX workbook yet; only sheets can`. This is a semantic
merge conflict between two branches that were merged independently:

- record 0269 (`1714b7d867`) added the kind gate `new_kind(node.kind())?`,
  which admits only `sheet`, at the top of
  `crates/litchi-ops/src/format/xlsx/structure.rs::add_inner`;
- record 0270 (`51d513641b`) added the `chart` branch *after* that line in
  the same function.

The merged trunk therefore refuses every XLSX chart add before it reaches
the chart code. The likely fix is to admit `chart` in `new_kind` (a
`NewKind::Chart` arm), or to move the chart branch above the gate. Every
other `litchi-ops` test target passes.

## 7. Gaps and follow-ups

- **Trunk regression in `litchi-ops`** (§6): the XLSX chart add is refused
  by the record-0269 kind gate. It is owned by the 0269/0270 owners or the
  coordinator, because `structure.rs` is shared dispatch.

- The coordinator's local gate should run the harness tests (§5). This
  record cannot change that gate.
- Lockfile drift in the standalone projects (`tools/perf-baseline`,
  `tools/native-resave`) recurs whenever a leaf crate gains a dependency.
  CI's `--locked` catches it at the PR into `main`. A local
  `cargo metadata --locked` check in the merge gate would catch it earlier.
- No timing was taken for the v3 routes in this record. 0264 §4 holds the
  16384-bound measurements, and the next paired PPT measurement should cite
  the `_aggregate_v3` names.
