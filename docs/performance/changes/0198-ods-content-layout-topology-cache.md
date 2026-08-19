# Change 0198: ODS ContentLayout caches the derived table/row topology

Date: 2026-08-18

## Decision

Extend the 0195 per-owner content-layout cache to carry the derived worksheet
topology — the sorted direct table children of `office:spreadsheet` and each
table's sorted direct `table-row` children — so commits served by the cached
layout no longer re-scan the span vector for the table inventory and per-sheet
row lists. Accepted on the measured ODS source-backed selectors; no regression
pattern fired on any statistic.

## Mechanism and invariants

`crates/litchi-ods/src/worksheet/package.rs` `changed_row_edits` used to
re-derive the worksheet topology on every commit, including commits served by
the 0195 cached layout: one `one_spreadsheet` scan over all spans, one
`direct_children` scan for the table inventory, and one `direct_children` scan
per sheet for its physical rows. Each span visit performs namespace/local-name
string comparisons (~48-byte namespace URIs), so a commit over the two-sheet
corpus walked every span four times; profiling of the source-backed commit
path attributed 6.30% of commit-phase samples to `changed_row_edits` self time
plus a large share of the 5.65% `memcmp` total to these comparisons.

`ContentLayout` now additionally carries `tables: Vec<usize>` (sorted direct
table children of the single `office:spreadsheet`) and
`rows: Vec<Vec<usize>>` (each table's sorted direct `table-row` children),
built once by `build_layout` immediately after `scan`. The derivation is a
pure, infallible function of the spans, and the layout is only ever reused
with byte-identical input (the owner's retained `content_xml` is immutable),
so cached and uncached paths observe identical values.

Error ordering is unchanged on both paths: the per-call gates still run before
the layout is consulted; `one_spreadsheet` still runs at layout-build time on
the scanning path (same position relative to the scan as before); the
fallible inventory checks (`tables.len() != original.len()`,
`rows.len() != before.rows.len()`) still run per transaction against the
cached vectors. Early ineligible returns still yield `(None, None)` and never
retain a layout; the layout is returned only when the scan actually ran and
the transaction succeeded, matching the 0195 retention contract. The
row-inventory error text keeps its per-sheet index.

`replace_tables`, `try_replace_changed_rows_spliced`, and the flat
`replace_changed_rows` are behaviorally unchanged: the non-caching callers
build the same topology through `build_layout` with the same error order and
discard it.

## Matched release timing

Two frozen release binaries differ only in `worksheet/package.rs`; both carry
changes 0193-0196 as baseline and the identical 341-case selector matrix.
Control SHA-256 `db708afa17eddc7dab9911429c51d6d0cd676550f8c33f4458893f2ea1201cff`
(bit-identical to the 0196 candidate and the 0197 control, confirming build
reproducibility), candidate SHA-256
`4e5662b270a30fc93af0b72bbe8c80d03adc82a73612a85085e6e3ac020954e1`.
Fresh CPU-2-pinned processes ran `A1 control, B1 candidate, B2 candidate,
A2 control`, 30 warmups and 500 retained samples per leg over the three
existing ODS source-backed edit selectors. The predeclared p50/mean/p95/p99
drift ceilings are 5%/5%/10%/15%; a statistic is accepted only when both
paired directions are lower and both drifts pass its ceiling. Every leg
reports all embedded verification flags true.

### Repeated-edit selector (`ods_source_backed_repeated_edit`)

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| total p50 | 0.50% | 0.87% | 0.74% | 0.37% | accept |
| total mean | 0.48% | 1.29% | 1.03% | 0.21% | accept |
| total p95 | 0.25% | 3.26% | 3.00% | -0.11% | accept |
| total p99 | 0.19% | 5.77% | 3.68% | -2.12% | accept |
| stage mean | 2.05% | 0.30% | 0.13% | 1.92% | accept |
| stage p95 | 3.09% | 2.22% | 1.16% | 2.07% | accept |
| stage p99 | 4.07% | 5.10% | -0.86% | -1.92% | accept |
| commit p50 | 2.20% | 2.89% | 0.94% | 0.23% | accept |
| commit mean | 2.44% | 2.95% | 0.84% | 0.32% | accept |
| commit p95 | 4.03% | 3.98% | -0.27% | -0.23% | accept |
| commit p99 | 6.11% | 7.72% | 2.54% | 0.78% | accept |
| publication p50 | 0.08% | 0.48% | 0.65% | 0.26% | accept |

Stage p50 (+1.45%/-0.16%) and publication mean/p95/p99 straddle zero and are
withheld as neutral. No statistic shows a regression pattern.

### One-edit guardrail (`ods_source_backed_one_edit_save`)

Single transaction: the layout is built once and never reused, so the paths
are near-identical by construction. Lifecycle mean (0.15%/0.70% lower) and
p95 (0.53%/1.34% lower) and commit p99 (2.24%/3.75% lower) accept; all other
statistics straddle zero and are withheld as neutral. No regression trigger
fired.

### One-percent guardrail (`ods_source_backed_one_percent_edit_save`)

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| lifecycle p50 | 1.06% | 1.92% | 0.24% | -0.62% | accept |
| lifecycle mean | 0.94% | 1.87% | 0.28% | -0.67% | accept |
| lifecycle p95 | 0.39% | 2.89% | 0.64% | -1.89% | accept |
| commit p50 | 2.11% | 1.34% | -1.39% | -0.61% | accept |
| commit mean | 1.76% | 1.82% | -1.04% | -1.10% | accept |
| commit p95 | 1.20% | 5.46% | 0.96% | -3.39% | accept |

Lifecycle p99 measures the candidate 0.34%/0.57% slower in both directions
(sub-1% deep-tail; withheld) and commit p99 straddles (-4.56%/+4.94%;
withheld). No other statistic shows a regression pattern.

The claim is scoped to the measured ODS source-backed selectors; the flat and
spliced callers share the mechanism but are not re-measured here. No
allocation/RSS, physical-I/O, cold-cache, producer, or broad ODF claim is
made.

## Verification

```text
cargo test --locked -p litchi-ods --all-targets
cargo clippy --locked -p litchi-ods --lib --all-features --no-deps -- -D warnings
RUSTDOCFLAGS='-D warnings' cargo doc --locked -p litchi-ods --all-features --no-deps
cargo fmt --all -- --check
python3 tools/check_crate_boundaries.py
```

litchi-ods passes 342/342, including the two multi-row window regression
tests retained from the withheld 0197 exploration, which exercise
`changed_row_edits` through both the scanning and cached-layout paths. Scoped
strict Clippy, rustdoc, formatting, and crate-boundary checks pass. No
manifest changed, so `cargo sort` is not implicated. The full-workspace
`--all-features` debug gate does not fit on this host's disk; the change is
private to `litchi-ods` internals with no signature changes, so the
crate-scoped gate above is the relevant boundary.

Artifacts:

- repeated-edit: [summary](../results/ods-repeated-edit-0198-summary.json),
  [manifest](../results/ods-repeated-edit-0198-manifest.json)
- one-edit guardrail: [summary](../results/ods-one-edit-save-0198-summary.json),
  [manifest](../results/ods-one-edit-save-0198-manifest.json)
- one-percent guardrail:
  [summary](../results/ods-one-percent-edit-save-0198-summary.json),
  [manifest](../results/ods-one-percent-edit-save-0198-manifest.json)
- compressed A1/B1/B2/A2 raw reports listed in each manifest
