# Change 0195: ODS source-backed content layout scan cache

Date: 2026-08-18

## Decision

Cache the row-local edit layout scan of `content.xml` at most once per
`SourceBackedSpreadsheet` owner instead of rescanning the full document on
every source-backed commit.

The first source-backed commit pass
(`worksheet::package::replace_changed_rows_from_content_xml`) scans the
entire retained `content.xml` into an element-span layout before splicing
the changed rows. The scan input is the owner's immutable `content_xml`
projection, bound to the captured `SourceVersion` and cloned by `Arc` into
every staged snapshot, so the layout is a pure function of data that never
changes for the owner's lifetime. Repeated transactions on one owner
repeated identical work: profiling of the source-backed commit path
attributed 4.37% of commit-phase samples to `worksheet::package::scan` plus
1.46% to `resolve_namespace`, on top of the shared quick_xml event machinery
the scan drives.

The cache is a private `OnceLock<ContentLayout>` on the owner, mirroring the
0193 edit-protection cache. Only successful scans are cached: the first
commit on an owner runs the unchanged scanning path and retains the layout
it computed as a side effect; gate refusals before the scan and scan
failures leave the cache empty, so no error is ever cached and no refusal is
weakened. A concurrent first commit may scan twice, with one result
published.

## Mechanism and invariants

`changed_row_edits` now takes an optional pre-scanned `&ContentLayout` and
returns the layout it scanned. The per-call gates
(`validate_content_xml_size`, the physical-run limit, per-candidate
`validate_sheet`, and the sheet-count match) always run first, in the
original order, on both the scanning and the cached path — error priority is
unchanged because a cached layout can only exist after a successful scan of
the identical input. Refusal outcomes, error texts, the structural-edit
fallback (`None` ineligible result), and the publication-boundary reparse
and compare are untouched. Single-transaction lifecycles pay exactly one
scan, as before.

Three `pub(crate)` entry points replace the previous one:
`replace_changed_rows_from_content_xml_with_layout` (cached path),
`replace_changed_rows_from_content_xml_retaining_layout` (scanning path that
returns the layout for caching), and the removed wrapper's other callers
(`replace_changed_rows`, `try_replace_changed_rows_spliced`) keep their
signatures and behavior. A focused inline test proves the cache is empty at
open, populated after the first successful commit, reused by pointer
identity on a second commit, performs zero source reads in either case, and
leaves both commits' outputs correct.

## Matched release timing

Two frozen release binaries differ only in the litchi-ods commit layout-scan
path; both contain the 0193 edit-protection cache and the 0194
`validate_text` byte scan as baseline and the identical 341-case selector
matrix. (The control binary is bit-identical to the 0194 candidate binary.)
Control SHA-256
`e193366cd3b85a6e23a1b978be9e0e1e28fdc386c99e84f56a1cba559266d163`,
candidate SHA-256
`ff3d1d989214ad19ba597fe9c5254d8548124575d4d4afc20281b105c07c4ea5`.
Fresh CPU-2-pinned processes ran `A1 control, B1 candidate, B2 candidate, A2
control`, 30 warmups and 500 retained samples per leg. The predeclared
p50/mean/p95/p99 drift ceilings are 5%/5%/10%/15%; a statistic is accepted
only when both paired directions are lower and both drifts pass its ceiling.
Every leg reports all embedded verification flags true.

### Repeated-edit selector (`ods_source_backed_repeated_edit`)

Four-transaction totals:

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| p50 | 7.00% | 5.55% | -0.98% | 0.57% | accept |
| mean | 7.11% | 5.96% | -1.00% | 0.23% | accept |
| p95 | 8.14% | 6.79% | -2.04% | -0.60% | accept |
| p99 | 7.62% | 12.49% | 0.06% | -5.22% | accept |

Commit phase (per-sample sum of the four commits):

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| p50 | 27.47% | 26.14% | -1.33% | 0.47% | accept |
| mean | 27.10% | 26.47% | -1.00% | -0.14% | accept |
| p95 | 26.28% | 26.92% | -1.68% | -2.53% | accept |
| p99 | 25.13% | 36.98% | 8.42% | -8.74% | accept |

Three of the four layout scans are removed per sample. The accepted
reduction exceeds the profiled scan self-cost because the scan also drove
the shared quick_xml namespace machinery and allocator traffic. Stage phase:
mean and p95 accepted (1.80%/0.27% and 2.15%/1.83% lower), p50 and p99
withheld (paired directions disagree inside noise). Publication phase: mean,
p95, and p99 accepted (1.32%/0.11%, 3.20%/1.36%, 3.18%/3.65% lower), p50
withheld (directions disagree). No regression trigger fired.

### Guardrail selectors

`ods_source_backed_one_edit_save` (single-cell lifecycle): lifecycle
p50/mean/p95/p99 all withheld as neutral — the first commit on an owner
still scans once by design. Commit phase p50/mean/p95/p99 all accepted
(5.50%/4.55%, 5.02%/4.06%, 3.62%/2.19%, 1.98%/2.51% lower); the scanning
path itself got marginally cheaper from the refactor's codegen, a scoped
codegen-level observation, not a mechanism claim.

`ods_source_backed_one_percent_edit_save` (21-cell lifecycle): lifecycle p99
accepted (1.65%/2.22% lower), lifecycle p50/mean/p95 withheld (directions
disagree within 1.9%). Commit phase p50/mean/p99 accepted (1.39%/2.91%,
1.10%/3.19%, 2.77%/1.64% lower), p95 withheld (directions disagree). No
regression trigger fired on any withheld statistic.

## Verification

```text
cargo test --locked -p litchi-ods --all-targets
cargo clippy --locked -p litchi-ods --lib --test source_cell_transactions -- -D warnings
RUSTDOCFLAGS='-D warnings' cargo doc --locked -p litchi-ods --no-deps
cargo fmt --all -- --check
python3 tools/check_crate_boundaries.py
```

The litchi-ods suite passes 340/340 including the new laziness/sharing test.
Scoped strict Clippy, rustdoc, formatting, and crate-boundary checks pass.
Unrelated pre-existing strict-Clippy failures in untouched litchi-ods test
files (`facade_round_trip.rs`, `tracked_changes*.rs`) reproduce identically
without this change and are outside its scope.

Artifacts:

- repeated-edit: [summary](../results/ods-repeated-edit-0195-summary.json),
  [manifest](../results/ods-repeated-edit-0195-manifest.json)
- one-edit guardrail: [summary](../results/ods-one-edit-save-0195-summary.json),
  [manifest](../results/ods-one-edit-save-0195-manifest.json)
- one-percent guardrail:
  [summary](../results/ods-one-percent-edit-save-0195-summary.json),
  [manifest](../results/ods-one-percent-edit-save-0195-manifest.json)
- compressed A1/B1/B2/A2 raw reports listed in each manifest
