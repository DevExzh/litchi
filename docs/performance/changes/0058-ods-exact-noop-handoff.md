# Change 0058: ODS exact no-op handoff

Date: 2026-08-12

Status: accepted after review-trigger analysis

## Decision

Stop an exact unified ODS worksheet no-op at the nested worksheet commit and
construct its durable empty patch without rediscovering package effects. The
patch still retains the exact source allocation in both reversible directions,
so deterministic transfer, source applicability, inverse behavior and history
weight remain unchanged.

This is private ODS transaction work. It adds no public API, dependency,
archive type, cache, global state, runtime, lock, unsafe code or persisted
index. OLE2, OOXML, RTF and the other ODF format crates are unchanged. iWork
and IWA remain excluded while their crates are being modified independently.

## Problem and attribution

`Edit::worksheets` previously ignored the nested worksheet commit's unchanged
result. It cloned the complete candidate into another `Vec` and handed it to
generic package staging. The final unified commit then built an empty durable
patch by calling `changed_effects(source, source)`. That path opened and
validated the same ZIP package twice, compared every raw member, and performed
logical fallback work merely to prove that an already-established exact no-op
had no effects.

The frozen large exact-no-op control profile attributed repeated XML
validation, namespace resolution, comparisons and moves to
`changed_package_files -> Patch::build`. For example, 3.31% of process self
cycles in `validate_content_xml` was directly under that redundant stack;
related resolver, `memcmp` and `memmove` samples appeared under the same owner.
The complete profile has 2K cycle samples and zero lost samples.

## Implementation and invariants

- `Edit::worksheets` returns immediately when the nested worksheet commit is
  unchanged. It does not clone the unchanged snapshot or call the generic
  staging/diff path.
- `Patch::build` recognizes only the exact empty-step case. It first proves
  that source and target bytes are equal; a changed target without a semantic
  owner remains a typed refusal.
- The no-op patch hashes and inserts the retained source `Arc<[u8]>` once, then
  clones the content-addressed bundle for the reverse direction. Both bundles
  therefore retain the same immutable allocation and the same BlobId.
- The canonical wire envelope is unchanged: it contains no operations and one
  identical package blob in each direction. It remains independently
  decodable because the package bytes were not removed from the wire format.
- Changed commits still run security policy enforcement, package bounds,
  authored-part validation, complete snapshot reopen/readback, effect
  attribution and semantic-owner checks. Signed, encrypted and protected exact
  no-ops retain their established permission and exact-byte behavior.

Focused tests prove shared allocation identity and content address, empty
forward operations, symmetric deterministic JSON, inverse behavior, wire
decode/apply, exact bytes, and refusal of a changed ownerless package.

## Measurement method

Base revision: `0e31859fbb67f69dc4b0b48ad7e5eff75d9044bb`.

Both binaries use the unchanged standalone harness and release profile:

- control SHA-256:
  `b381629b2b30a860b3547c310b62cca0e736d7990ac15ceac1196744780b8912`;
- candidate SHA-256:
  `daa6e18a97d703757e663205cd543e1f60c1590af42cb9eeb6832e6058fb9366`.

Environment: Rust 1.95.0 / LLVM 22.1.2, Linux 6.8.0-101-generic,
x86-64 AMD EPYC 9575F VM, Rust system allocator, CPU 11 pinned with `taskset`.
The deterministic large corpus contains 32,768 cells in a 98,892-byte ODS;
its archive SHA-256 is
`7f0c43561602aedac7c5e91915f55b3515371d327ae69ac7fc0fe42b655db3f2`.
Every iteration verifies exact unchanged bytes, empty patch semantics, patch
apply/inverse and complete ODS readback outside the measured interval.

The primary used three direction-balanced ABBA pairs. Each leg had 10 warmups
and 50 timed samples, yielding 300 observations per state. Pooled samples and
confidence intervals are in the
[`primary summary`](../results/ods-exact-noop-handoff-primary-summary.json).

## Latency result

| Large exact no-op edit/save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 64.024 ms | 49.133 ms | **-23.26%** |
| mean | 64.305 ms | 49.383 ms | **-23.21%** |
| p95 | 67.327 ms | 51.876 ms | **-22.95%** |
| p99 | 70.299 ms | 53.521 ms | **-23.87%** |

The before mean 95% interval is 64.055-64.554 ms; the after interval is
49.242-49.524 ms. Every individual primary leg improves, with p50 reductions
between 22.0% and 23.9%.

## Guards and review trigger

The standard tiny, medium and large ODS guards used one direction-balanced
ABBA pair per shape. Tiny and medium have 200 pooled observations per state;
large has 60. The materially timed results are:

| Case | Tiny p50 | Medium p50 | Large p50 |
|---|---:|---:|---:|
| Exact no-op edit/save | **-39.24%** | **-22.38%** | **-22.79%** |
| Open | +1.33% | +0.54% | -0.53% |
| One-cell edit/save | -2.20% | +0.65% | +1.71% |
| Full cell text | +4.17% | -0.15% | -1.28% |

The media-rich one-cell changed-publication guard is also neutral: p50
+0.48%, mean +0.86%, p95 +0.96% and p99 +1.51% across 40 samples per state.
Its summary is
[`here`](../results/ods-exact-noop-handoff-media-guard-summary.json).

The guard matrix did trigger review for read-only public cell access, which
cannot execute either changed source branch. A 2,000-sample-per-state repeat
measured large cell-sweep p50 at 377.5 us before and 458.8 us after (+21.5%)
and one-cell p50 at 1.602 us versus 1.773 us (+10.7%). This result is disclosed
in the
[`repeat summary`](../results/ods-exact-noop-handoff-readonly-repeat-summary.json),
not hidden in the ODS aggregate.

Source inspection confirms that no facade/cell-locator code changed. The
frozen binaries contain the same 0x38f-byte `Spreadsheet::cell` function, and
normalized disassembly has the same instruction stream; only three
RIP-relative link addresses differ. The exact
[`layout diff`](../results/ods-exact-noop-handoff-cell-accessor-layout-diff.txt)
records those addresses. This is a real observation for these two linked
executables, but it is placement-sensitive and not an executed source-path
regression attributable to the optimization. The ordinary full-text guard,
which exercises the same cell access plus material text work, improves 1.28%
at large shape. No production workaround based on one executable's link
placement was added.

Complete pooled guard distributions are retained in the
[`tiny`](../results/ods-exact-noop-handoff-guards-tiny-summary.json),
[`medium`](../results/ods-exact-noop-handoff-guards-medium-summary.json) and
[`large`](../results/ods-exact-noop-handoff-guards-large-summary.json)
summaries.

## CPU and allocation evidence

Matched whole-process `perf stat` ABBA legs used five warmups and 20 timed
large no-op commits per process:

| Counter | Before mean/process | After mean/process | Delta |
|---|---:|---:|---:|
| Cycles | 15.629 billion | 14.059 billion | **-10.05%** |
| Instructions | 64.308 billion | 57.530 billion | **-10.54%** |
| Branches | 13.214 billion | 11.956 billion | **-9.52%** |
| Branch misses | 10.609 million | 8.682 million | **-18.16%** |
| Cache misses | 18.890 million | 16.090 million | **-14.82%** |

The matched candidate profile removes every resolved
`changed_package_files -> Patch::build` sample and retains zero lost samples.
Reports are
[`before`](../results/ods-exact-noop-handoff-before-perf-report.txt) and
[`after`](../results/ods-exact-noop-handoff-after-perf-report.txt); raw counter
CSVs use the `ods-exact-noop-handoff-stat-*` prefix.

Heaptrack over three measured commits reports 5,947,560 -> 5,945,478
allocation calls (-2,082), 1,464,278 -> 1,463,972 temporary allocations
(-306), identical 41.70 MiB peak heap, and 47.93 -> 47.60 MiB
Heaptrack-inclusive RSS. Leak reports remain 1.78 KiB. Reports are
[`before`](../results/ods-exact-noop-handoff-before-heaptrack.txt) and
[`after`](../results/ods-exact-noop-handoff-after-heaptrack.txt).

Four uninstrumented GNU Time processes per state report mean maximum RSS of
43,153 KiB before and 43,121 KiB after (-0.07%, flat at process resolution).

## Validation

Passed on the final source:

- the focused allocation-sharing and public deterministic-wire no-op tests;
- complete `litchi-ods --all-features` unit, integration and doc-test suites;
- warning-denied ODS production Clippy, changed integration-test Clippy and
  rustdoc;
- all 33 performance-harness tests and warning-denied all-target Clippy;
- warning-denied ODF common all-target Clippy and rustdoc, rechecking the
  deprecation cleanup from `1194fbc7f`;
- formatter, JSON parsing, link-target, whitespace and final-diff checks.

The broader ODS all-target Clippy command still has the previously documented
unrelated test-only lints; the changed test target is clean. No dedicated ODS
fuzz manifest exists in the repository.

## Limitations and next work

This change removes work only for exact unified ODS no-ops and the worksheet
handoff that establishes them. It does not remove the mandatory initial
snapshot readback, nested worksheet parse/validation or durable source hash.
Changed commits receive no shortcut.

Remaining high-ROI non-iWork work includes native XLS final-owner publication,
source-backed XLSX page-break and other OOXML edits, and broader RTF
formatting/media paths. Those remain separate OLE2, OOXML and RTF tranches.
