# 0356b — PPTX: retire `SourceBackedPresentationEditor`

- Status: implemented on `uo2/w2-pptx` after the coordinator accepted record
  [0356a](0356a-pptx-review-follow-ups.md) (tip `3e392d750f`); trunk merged
  at `59922faa26` (trunk `7f36f7b909`, records 0357a and 0356a); for
  coordinator review.
- Date: 2026-10-20.
- Input: the review item 0356b: retire the public
  `SourceBackedPresentationEditor`, move its cross-slide-copy and
  slide-removal plan routes onto `source::Package`, migrate perf-baseline,
  make the old owner crate-private and ban the name; measure those routes
  (A-B-B-A, within 5% of before); teach the inventory to classify
  `builder::*` drafts as drafts so the PPTX inventory has zero owner rows.
- Owned: as 0356a (`crates/litchi-pptx/**`, `tools/perf-baseline`, the
  `pptx-*` profiles, line W2, `tools/api_debt_inventory.py`,
  `results/0351/inventory.tsv`), plus the 0356b parser of
  `tools/unified_bench_matrix.py`.
- `performance_claim`: measured (§4). The three source-backed cross-copy
  routes are within −1.6%…−0.3% of before (after/before, A-B-B-A three
  times), with byte-identical output.

## 1. Result

| Item | Done |
|---|---|
| Cross-slide copy on `source::Package` | New family `litchi_pptx::source::cross_slide_copy`: `Package::plan_cross_slide_copy(&source, source_slide, layout_slide, insertion) -> Plan`, `Package::edit_cross_slide_copy() -> Edit` (`stage(&Plan)`, `copy(...)`, `reset`, `commit`), `Package::apply_cross_slide_copy(&Patch) -> Commit` (the core `Commit<Package, Patch>` triad); `Copied` (was `SourceBackedCrossSlideCopySnapshot`) reports what a staged copy publishes |
| Slide removal on `source::Package` | `source::Edit::remove_slide(Position)` (already the source-backed primary verb, with the same topology planner); the editor's `plan_slide_removal` / `publish_slide_removal_to_stream` route and `SourceBackedSlideRemovalPlan` are crate-private and test-only |
| Old owner crate-private and banned | `SourceBackedPresentationEditor`, `SourceBackedIntoEditorError`, `SourceBackedCrossSlideCopyPlan`, `SourceBackedCrossSlideCopySnapshot`, `SourceBackedSlideRemovalPlan` are `pub(crate)` and banned in `pptx-minimal` and `pptx-full` (with the 20 editor types of 0356a); 6 obsolete bans scoped on the editor are removed |
| perf-baseline migrated | every cross-copy route (`lib.rs`, `pptx_cache_retention`, `pptx_pair_lifecycle`, `pptx_provider_lifecycle`, `pptx_retention`, `bin/pptx_native_copy_probe`, `bin/pptx_external_cross_copy`) opens a `source::Package` and times `edit_cross_slide_copy → stage → commit → write_to`; `implementation`/`timing_scope` strings name the new route |
| Inventory: `builder::*` are drafts | `tools/api_debt_inventory.py` classifies any type under a `builder` module as a draft, neither a mutable owner nor an unregistered draft; the PPTX section of `results/0351/inventory.tsv` has zero rows |

## 2. What landed

### 2.1 `source::cross_slide_copy`

`crates/litchi-pptx/src/presentation/source/primary/cross_slide_copy.rs`.
A `Plan` holds the checked copy plan and the destination it was planned on;
staging it on another package (another destination source) is
`Error::PatchConflict`. A committed package is source-backed and deferred
(the W1 0355a shape, as `source::svg` in 0356a): `write_to`/`save` rerun the
bounded planner against both sources and stream the destination with the
copied slide, its media and chart parts added and every untouched member
raw-copied — the exact bytes of the former
`publish_cross_slide_copy_to_stream`. Reading or editing the committed
package materializes the published artifact once.

Several copies stage in one edit: `Edit::copy` plans against the staged
package, so a second copy sees the first. A plan made on the edit's source
cannot be staged after an earlier copy (its destination is no longer the
staged one); `copy` is the route for chained copies.

The engine and its tests stay: `source_cross_copy.rs` (planner and
publisher, `pub(crate)`), with the integration tests moved into
`src/presentation/source/engine_tests/` (they need the crate-private
editor). New family tests: `cross_slide_copy_family_publishes_a_source_backed_package`
(the committed package's bytes equal the streamed route's, the source
package is unchanged, replay, inverse on the commit's own package only, a
reopened destination refused, chained copy collision, no-op commit).

### 2.2 Retirement

- `presentation/source.rs`: the editor (`#[derive(Clone)]` for the deferred
  publisher), `try_into_editor` and `SourceBackedIntoEditorError` are
  `pub(crate)`; `lib.rs` and `presentation/mod.rs` drop the re-exports (a
  `#[cfg(test)]` crate-root alias keeps the in-crate tests' paths).
- `source_removal.rs`: `SourceBackedSlideRemovalPlan` is `pub(crate)`; its
  now-unused public accessors are removed, the remaining dead test-only
  items are `#[cfg(test)]`.
- `open_options.rs`, `leniency.rs`, `FEATURE_MATRIX.md` (rows "Source-backed
  cross-slide image/chart leaf copy" and "Source-backed SVG attachment
  lifecycle"), `tools/perf-baseline/README.md` and the W2 line of
  `docs/API_GLOSSARY.md` name the `source::Package` routes.
- `tests/source_related.rs`, `tests/pptx_tolerance_report.rs` use
  `source::Package`.

### 2.3 Inventory

`tools/api_debt_inventory.py`: `draft = "builder" in path[:-1]`, so
`litchi_pptx::builder::{Builder, ...}` are drafts by module, like the other
formats' builders. With 0356a's core-`Commit` alias rule, the PPTX section
is now just the header. `results/0351/inventory.tsv` is regenerated over the
merged tree (`(deleted-coord-artifact)`: rustdoc JSON of every
crate, then the tool): 57 `mut_owner` rows leave, every one a `builder::`
path — 10 PPTX, 44 DOCX and 3 PPT (the rule is not PPTX-specific); no
other row changes.

## 3. Behaviour (recorded for review)

- A cross-copy `Patch` is exact-source: it replays onto a package over the
  same opened destination source (same source lineage) and is refused with
  `PatchConflict` elsewhere — including another opening of the same bytes.
- Its inverse applies only to the commit's own package (it restores the
  package the commit was made on, by identity of content). It is not a
  durable removal of the copied slide; stage `source::Edit::remove_slide`
  for that.
- Signed, stale or foreign destinations keep the old typed refusals; the
  planner, limits and fences are unchanged.

## 4. Measurements

perf-baseline `pptx_source_backed_cross_copy_plain`,
`pptx_source_backed_cross_copy_plain_lifecycle`,
`pptx_source_backed_cross_copy_media_rich_lifecycle`: before = a release
binary built from `git archive 3e392d750f` (editor route), after = the
branch at `dcaa2aa4cb` (family route; the trunk merge changes no file under
`crates/litchi-pptx`, `litchi-core`, `litchi-opc` or `tools/perf-baseline`).
Order A-B-B-A three times, `taskset -c 7`, `--warmup 3 --samples 30`;
value = median of the per-run p50s.

| Case | Before (µs) | After (µs) | After/before | Same output |
|---|---:|---:|---:|---|
| `pptx_source_backed_cross_copy_plain` | 2,541.9 | 2,502.4 | 0.985 | yes |
| `pptx_source_backed_cross_copy_plain_lifecycle` | 2,973.1 | 2,937.6 | 0.988 | yes |
| `pptx_source_backed_cross_copy_media_rich_lifecycle` | 13,304.8 | 13,264.8 | 0.997 | yes |

Evidence: `results/0356b/` (12 run JSONs of ~110 KiB, `summary.json`,
`binaries.sha256`). The bench matrix parser `p0356b` adds the three rows
(budget 1.05, regression).

Command (each binary, from its tree):

```sh
taskset -c 7 target/release/litchi-perf-baseline \
  --case pptx_source_backed_cross_copy_plain,pptx_source_backed_cross_copy_plain_lifecycle,pptx_source_backed_cross_copy_media_rich_lifecycle \
  --warmup 3 --samples 30 --json <run>.json
```

The source-backed slide-removal route has no perf-baseline case; its public
route, `source::Edit::remove_slide`, is unchanged code.

## 5. Gates

On the merged tree (`59922faa26` plus this record's changes), sequential,
`(deleted-coord-artifact)`; every command exited 0.

| Gate | Command | Exit |
|---|---|---:|
| fmt | `cargo fmt --all --check` | 0 |
| clippy | `cargo clippy -p litchi-pptx -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| clippy, ops feature sets | `cargo clippy -p litchi-ops` with `--no-default-features`, `--no-default-features --features pptx`, and every feature but `encryption` (`--all-targets -- -D warnings`) | 0, 0, 0 |
| docs | `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-pptx -p litchi-ops` | 0 |
| PPTX tests | `cargo test -p litchi-pptx --all-features` (1,863 passed, 3 ignored) | 0 |
| ops tests | `cargo test -p litchi-ops --all-features --lib`, `--doc`, and every integration target in chunks of 20 (951 passed in lib + integration) | 0, 0, 0 |
| facade, CLI, Python | `cargo test -p litchi --all-features`, `cargo test -p litchi-cli --all-features`, `crates/pyo3-litchi/tests/run_tests.sh` | 0, 0, 0 |
| glossary | `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (44 profiles, 3,025 banned names, 437 families) | 0 |
| perf-baseline | `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| tool checks | `check_crate_boundaries.py`, `non_iwork_gate.py verify`, the tool unittests, `officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check`, `check_feature_matrix.py` | 0 each |
| goldens | `LITCHI_BLESS=1 cargo test -p litchi-ops --test capabilities --test registry` (all features; all but `encryption`), then `git status --porcelain -- crates/litchi-ops` empty | 0, 0, 0 |

## 6. Follow-ups

- None from 0356b's scope. The engine's internal `SourceBacked*` names
  remain inside the crate; renaming them is internal clean-up only.
