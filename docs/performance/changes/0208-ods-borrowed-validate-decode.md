# Change 0208: ODS commit validate-reparse decodes to borrowed strings

Date: 2026-08-19

## Decision

**Withheld — reverted.** The deterministic allocation win (commit
transactions -43.0% allocations, -53.5% commit-only) is real, but the
latency gate failed: both executed-phase workloads that showed an
adverse both-directions pattern (one-edit-save commit p50/mean/p95/p99,
repeated-edit commit p50/mean) **reproduced** the pattern in their
single permitted rule-2 reruns (one-edit commit p50/mean/p95 adverse at
-2.12% to -6.94% with clean drifts; repeated-edit commit all-four
adverse at -2.40% to -5.84% with clean drifts). Per banking rule 2 the
change is withheld regardless of the floor. The pattern is consistent
with code-layout shift (the sibling one-percent workload showed
accepted-favorable below-floor readings on the identical code path, and
the change strictly removes work), but the methodology does not bank on
that argument — a reproduced adverse pattern on an executed phase is
binding. The source change was reverted; the tree is back at the banked
0207 state (verified by bit-exact rebuild against the 0207 control
binary).

## Mechanism and invariants

Commit-path profiling (commitonly transactions) showed the path is
allocation-dominated. Deterministic counting-allocator attribution placed
15,753 of 36,616 allocations per transaction (43.0%) in
`validate_rewritable_row` / `validate_modeled_attributes`
(`worksheet/package.rs`): the changed-row validate-reparse materialized
an owned `String` per element local name, per element/attribute
namespace URI, and per End-tag name — every one used only for equality
comparison against constants and dropped immediately (~250 owned Strings
per re-parsed row; the staged-edit window covers most of the 64-row
corpus).

The change replaces the owned `decode()` / `resolve_namespace()` calls
with borrowed variants (`str::from_utf8` + identical `map_err` message,
same call sites, same order) and classifies namespaces through a
`NamespaceKind {Table, Office, Text, Other}` equality partition where
quick-xml's `ResolveResult<'_>` borrow ties prevent direct comparison.
`scan()` keeps owned Strings (they are stored in `Span`). The orphaned
helpers `is_element_name` / `is_modeled_row_element` are removed.

Invariants:

- Identical error messages at identical call sites in identical order;
  the UTF-8 error branch is kept at its original position although
  structurally unreachable (the source is a `&str` sliced at ASCII
  delimiters).
- Namespace classification by URI equality is provably the same partition
  as the historical `Option<&str>` constant comparisons — including
  custom prefixes bound to the office/table/text URIs (pinned by two new
  unit tests, since the pre-existing literal body-marker check in
  `litchi-odf-common` rejects custom-prefix documents at the facade
  level).
- No public API change; `package.rs` internals only.

Verification: the full `litchi-ods` suite (376 tests, +2 new) passes;
fmt, clippy (`-D warnings`), rustdoc (`-D warnings`), and
`tools/check_crate_boundaries.py` pass. Counting-allocator driver:
commit transactions 36,616 → 20,863 allocations (-43.0%; commit-only
29,473 → 13,720, -53.5%) and 5,042,068 → 4,698,216 bytes; source-open
unchanged at 8,665 allocations / 1,829,857 bytes (exactly the 0207
baseline — the open path executes no changed code).

## Matched release timing

Two frozen release binaries differ only in the borrowed validate-reparse
decoding; both carry changes 0192-0196, 0198-0202, 0204, 0206, and 0207.
Control SHA-256 `57270d24894a7047682146f4a6a68d428ecc51b3d1270e5b93d90d0fddcb284b`
(the banked 0207 binary), candidate SHA-256
`6edff9a2858b17b3904c06106edfabdfd7e81765de11be2512e1a2f783754a6b`.
Fresh CPU-2-pinned processes ran `A1 control, B1 candidate, B2 candidate,
A2 control`, 30 warmups and 500 retained samples per leg, drift ceilings
5%/5%/10%/15% (p50/mean/p95/p99). The 0205 floor rule applies: accepted
statistics below the calibrated floor are neutral, not claims; adverse
both-directions readings within the floor on phases executing no changed
code are layout readings and do not block.

### ods_file_source_open (no changed code)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | -3.43% | -5.07% | 1.04% | 2.64% | withheld; adverse both dirs, at/below floor 5.5% → layout reading |
| mean | -3.04% | -4.91% | 0.93% | 2.77% | withheld; adverse both dirs, below floor 5.5% → layout reading |
| p95 | -0.30% | -4.70% | -0.68% | 3.68% | withheld; adverse both dirs, 0.2pp over floor 4.5% → borderline layout reading (no mechanism: open path executes no changed code, allocation counts bit-identical) |
| p99 | 2.07% | 3.10% | 2.29% | 1.22% | accepted, below floor 36% → neutral |

### ods_file_eager_open (no changed code)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | -0.31% | -3.03% | -2.48% | 0.16% | withheld; adverse both dirs, 0.03pp over floor 3.0% → borderline layout reading |
| mean | -0.54% | -2.95% | -2.66% | -0.32% | withheld; adverse both dirs, below floor 3.0% → layout reading |
| p95 | -1.66% | -4.50% | -4.28% | -1.61% | withheld; adverse both dirs, 1.5pp over floor 3.0% → layout reading (no mechanism: eager open executes no changed code) |
| p99 | -3.88% | -1.02% | -2.45% | -5.14% | withheld; adverse both dirs, below floor 7% → layout reading |

Eager-open p95 exceeds the calibrated floor by 1.5pp on a phase that
executes zero changed code (source-open allocation counts are
bit-identical to the 0207 control, and the commit path is not touched by
either open workload). With no causal mechanism available, these are
recorded as layout readings; the single-rerun allowance is reserved for
phases that execute the changed code.

### ods_source_backed_one_edit_save (commit executes changed code)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| lifecycle p50 | -0.53% | -0.73% | 0.08% | 0.28% | withheld; adverse both dirs |
| lifecycle mean | -0.49% | -0.65% | 0.14% | 0.30% | withheld; adverse both dirs |
| lifecycle p95 | -0.33% | -0.98% | -0.02% | 0.63% | withheld; adverse both dirs |
| lifecycle p99 | -0.15% | 1.52% | 3.28% | 1.57% | withheld (disagreeing directions) |
| commit p50 | -2.39% | -1.97% | 0.22% | -0.19% | withheld; adverse both dirs, within floor 3.7% |
| commit mean | -2.44% | -2.20% | 0.44% | 0.20% | withheld; adverse both dirs, within floor 3.7% |
| commit p95 | -2.29% | -3.83% | -0.06% | 1.45% | withheld; adverse both dirs, within floor 5.8% |
| commit p99 | -1.31% | -4.03% | 7.95% | 10.85% | withheld; adverse both dirs, within floor 17% |

Adverse both-directions on the executed commit phase → rule 2 applies:
the single permitted rerun of this workload must clear the pattern
before banking.

#### Rule-2 rerun (one-edit-save)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| lifecycle p50 | -2.22% | -1.25% | 0.66% | -0.30% | withheld; adverse both dirs REPRODUCED |
| lifecycle mean | -1.89% | -1.56% | 0.13% | -0.19% | withheld; adverse both dirs REPRODUCED |
| lifecycle p95 | -0.75% | -2.36% | -1.07% | 0.50% | withheld; adverse both dirs REPRODUCED |
| lifecycle p99 | -1.94% | -9.37% | -6.78% | 0.01% | withheld; adverse both dirs REPRODUCED |
| commit p50 | -3.22% | -3.96% | -0.52% | 0.19% | withheld; adverse both dirs REPRODUCED |
| commit mean | -3.19% | -4.51% | -1.22% | 0.05% | withheld; adverse both dirs REPRODUCED |
| commit p95 | -2.12% | -6.94% | -3.45% | 1.10% | withheld; adverse both dirs REPRODUCED |
| commit p99 | 5.58% | -21.19% | -17.74% | 5.58% | withheld (disagreeing directions; control drift over ceiling) |

The rerun **reproduces** the adverse both-directions pattern on the
executed commit phase (p50/mean/p95, clean drifts) — and strengthens
it. Per rule 2 the change is withheld regardless of the floor.

### ods_source_backed_one_percent_edit_save (commit executes changed code)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| lifecycle p50 | 0.43% | -0.09% | -1.59% | -1.07% | withheld (disagreeing directions) |
| lifecycle mean | 0.50% | -0.41% | -1.61% | -0.71% | withheld (disagreeing directions) |
| lifecycle p95 | 0.24% | -2.24% | -1.81% | 0.64% | withheld (disagreeing directions) |
| lifecycle p99 | 0.71% | -4.59% | -0.16% | 5.18% | withheld (disagreeing directions) |
| commit p50 | 3.79% | 2.17% | -2.87% | -1.23% | accepted, min-paired 2.17% below floor 3.1% → neutral |
| commit mean | 3.35% | 1.88% | -2.63% | -1.15% | accepted, min-paired 1.88% below floor 3.1% → neutral |
| commit p95 | 3.72% | 0.04% | -5.01% | -1.37% | accepted, min-paired 0.04% below floor 4.6% → neutral |
| commit p99 | 2.75% | -2.07% | -4.01% | 0.75% | withheld (disagreeing directions) |

No adverse pattern on this workload.

### ods_source_backed_repeated_edit (commit executes changed code)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| total p50 | -0.07% | -0.03% | 0.47% | 0.43% | withheld; adverse both dirs, below floor 1.8% |
| total mean | 0.08% | -0.14% | 0.27% | 0.48% | withheld (disagreeing directions) |
| total p95 | 1.04% | -0.48% | -0.56% | 0.96% | withheld (disagreeing directions) |
| total p99 | 1.65% | -0.61% | -0.44% | 1.86% | withheld (disagreeing directions) |
| stage p50 | 3.84% | 1.74% | -0.49% | 1.69% | accepted, min-paired 1.74% below floor 3.8% → neutral |
| stage mean | 4.49% | 1.57% | -0.84% | 2.19% | accepted, min-paired 1.57% below floor 3.8% → neutral |
| stage p95 | 4.65% | 0.75% | -0.85% | 3.21% | accepted, min-paired 0.75% below floor 5.3% → neutral |
| stage p99 | 4.82% | -3.17% | -3.14% | 4.99% | withheld (disagreeing directions) |
| commit p50 | -1.48% | -2.97% | -0.71% | 0.74% | withheld; adverse both dirs, below floor 4.4% |
| commit mean | -0.99% | -2.91% | -0.89% | 1.00% | withheld; adverse both dirs, below floor 4.4% |
| commit p95 | 1.37% | -3.73% | -2.55% | 2.49% | withheld (disagreeing directions) |
| commit p99 | 4.93% | -3.84% | -3.65% | 5.24% | withheld (disagreeing directions) |
| publication p50 | 0.03% | 0.38% | 0.63% | 0.27% | accepted, below floor 1.1% → neutral |
| publication mean | 0.12% | 0.30% | 0.52% | 0.34% | accepted, below floor 1.1% → neutral |
| publication p95 | 0.94% | 0.42% | 0.05% | 0.58% | accepted, below floor 2% → neutral |
| publication p99 | 1.50% | -0.32% | 0.32% | 2.17% | withheld (disagreeing directions) |

Commit p50/mean show an adverse both-directions pattern on the executed
phase → rule 2 applies: the single permitted rerun of this workload must
clear the pattern before banking.

#### Rule-2 rerun (repeated-edit)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| total p50 | -1.04% | -0.56% | -0.02% | -0.48% | withheld; adverse both dirs REPRODUCED |
| total mean | -0.80% | -0.48% | 0.05% | -0.27% | withheld; adverse both dirs REPRODUCED |
| total p95 | -0.29% | 0.16% | 0.28% | -0.17% | withheld (disagreeing directions) |
| total p99 | -0.70% | -1.14% | 0.47% | 0.90% | withheld; adverse both dirs REPRODUCED |
| stage p50 | 0.23% | 2.14% | -0.71% | -2.60% | accepted, min-paired 0.23% below floor 3.8% → neutral |
| stage mean | 0.07% | 2.37% | -0.57% | -2.87% | accepted, min-paired 0.07% below floor 3.8% → neutral |
| stage p95 | -1.32% | 2.82% | 1.30% | -2.84% | withheld (disagreeing directions) |
| stage p99 | -6.34% | 2.47% | -0.04% | -8.32% | withheld (disagreeing directions) |
| commit p50 | -3.32% | -2.57% | 0.19% | -0.54% | withheld; adverse both dirs REPRODUCED |
| commit mean | -3.40% | -2.56% | 0.38% | -0.43% | withheld; adverse both dirs REPRODUCED |
| commit p95 | -3.01% | -2.40% | 0.28% | -0.30% | withheld; adverse both dirs REPRODUCED |
| commit p99 | -3.46% | -5.84% | -0.14% | 2.16% | withheld; adverse both dirs REPRODUCED |
| publication p50 | -0.49% | -0.26% | -0.06% | -0.28% | withheld; adverse both dirs, below floor 1.1% |
| publication mean | -0.34% | -0.13% | 0.01% | -0.20% | withheld; adverse both dirs, below floor 1.1% |
| publication p95 | -0.14% | 0.24% | 0.23% | -0.15% | withheld (disagreeing directions) |
| publication p99 | -0.14% | -1.60% | 0.69% | 2.16% | withheld; adverse both dirs, at floor 1% |

The rerun **reproduces** the adverse both-directions pattern on the
executed commit phase — now on all four statistics with clean drifts —
and the primary-run mixed patterns on total/publication also flip
adverse. Per rule 2 the change is withheld regardless of the floor.

## Verdict

**Withheld.** No claim is made. The counting-allocator allocation
reduction (-43.0% per commit transaction) is documented above as
mechanism evidence only; the reproduced adverse latency pattern on the
executed commit phases blocks banking. The change has been reverted from
the tree. Raw artifacts: `docs/performance/results/*-0208-*` (primary)
and `*-0208r-*` (rule-2 reruns).

Lesson for the series: on this corpus the commit phase is sensitive
enough to code layout that an allocation-only change with no latency
mechanism beyond allocator pressure can produce a reproducible
sub-5% adverse reading in both A/B directions. Allocation-only commit
changes should expect this gate outcome unless the counting-allocator
delta is accompanied by an above-floor latency mechanism.
