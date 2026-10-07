# 0959 — XLSX numeric ownership candidate and attribution correction

This batch independently corrects the 0958 stack-owner classification and
tests a private shared-string representation for numeric cells. **The Arc
candidate is rejected:** it reduces allocation calls but provides no primary
latency benefit and fails the one-edit control. Production remains unchanged.

## Corrected attribution

The 0958 helper scanned stack frames before candidate owner names. A nested
`MultiSourceEdit::commit` frame could therefore win even when an exact
`benchmark::stage_timed_commit` frame appeared later in the same stack. Its
commit-labelled bucket did not prove that the timed stage was unresolved.
The old helper also accepted substring owner matches rather than requiring
the full benchmark namespace. The 0958 packet is retained unchanged and its
main report now points to this correction.

The new independent selector requires both fully qualified
`litchi_perf_baseline::xlsx_managed_compression::benchmark` callsites and
rejects unknown frames. All six captures admit exactly the same calls and
requested bytes as before: **2,081,214 / 274,320,408**, across 5,160 selected
trace aggregates. Every one of the 2,604 formerly commit-labelled trace rows
also contains the exact stage and timed-iteration frames. Thus the earlier
allocation totals and numeric clone census stand; the two resolution labels
are superseded. Whole-process totals remain 13,032,599 / 3,406,101,409, with
six unknown allocations totaling 442,368 requested bytes retained outside
selection. Independent review reselected the raw rows and verified this result.

## Concrete allocating owners

Across the same six captures, ranking by the first meaningful XLSX owner
above allocator frames gives these largest requested-byte contributors:

| Owner | Requested bytes |
| --- | ---: |
| Parser::finish_cell | 102,826,752 |
| Parser::finish_parse | 32,669,568 |
| Store::merge_omitted_cells | 32,360,448 |
| MultiSourceEdit::apply_batch | 28,570,176 |
| validation::scan_attributes | 22,985,856 |

`finish_cell` grows the intermediate `RawCell` vector. `finish_parse` reserves
and materializes final `Stored` cells; these are separate live representations.
Complete source parsing establishes semantic ownership, source freshness and
error authority. Reduced output readback and omitted-cell merging already
limit subsequent parsing. Removing the source parse is not justified by these
allocation counts. A future bounded capacity hint must account for its scan
cost, source limits, error precedence and fallible fallback.

The complete ranking, individual cases and allocation sizes are retained in
the [bounded packet](results/change-0959/README.md). XML/parse, duplicate-check,
merge, numeric and copy tags overlap and must not be summed. This is requested
allocation-byte attribution, not copied bytes, resident memory or CPU time.
Numeric clone and construction tags each account for roughly 0.37% of selected
requested bytes; their allocation-call significance is larger. Historical
attribute-singleton strategies remain rejected without new end-to-end evidence.

## Candidate mechanism and semantic boundary

The isolated candidate changes private `Number(Box<str>)` storage to
`Number(Arc<str>)`. Omitted cell copies share exact lexical bytes. The public
`Number::new(impl Into<Box<str>>)` signature remains unchanged and still
validates before creating the number. A crate-private borrowed constructor
lets the full parser and both selected parser paths allocate the Arc directly.
Exact integer conversions use the existing itoa dependency's stack buffer;
all supported integer types are finite, so their existing infallible path
needs no redundant validator or assertion.

The public boxed constructor may now allocate twice: conversion to its boxed
argument followed by Arc allocation. The plain selected scanner also copies a
String into Arc where the baseline could move it into a Box. Atomic clone/drop
operations and control-block bytes are additional costs. These tradeoffs need
measurements; fewer clone allocations are not automatically a useful change.

Baseline and candidate probes measure the actual integrated types on Rust
1.95.0 x86_64, including their Option forms. All are unchanged:

| Type | Size, bytes | Alignment, bytes |
| --- | ---: | ---: |
| Number | 16 | 8 |
| Value | 24 | 8 |
| Cell | 72 | 8 |
| Stored | 184 | 8 |
| Formula | 72 | 8 |
| Content | 72 | 8 |

The inline alternatives in the probe are representation sketches, not
integrated candidate layouts. They do not prove the size of enclosing types
under an inline implementation. Only the Arc candidate proceeds to comparison.

The candidate preserves content-based equality/hash, Debug/Display, signed
zero, exponent spelling, Unicode whitespace and long lexemes. Nonfinite and
malformed input still uses the existing validator and tolerance/fallback paths.
Complete XML validation, inverse restoration, source matching and resource
limits remain mandatory. No unsafe code, ambient executor, public API,
compression policy or iWork changes are involved.

## Semantic and tooling checks

The final external candidate passes **1,400 library tests** and warning-denied
library Clippy. Focused checks cover numeric traits and Unicode whitespace,
invalid/nonfinite values, exact sparse parsing, malformed lexical repair,
selected scanner paths and their lexical/prolog differential controls. Scoped
format checks pass. The initial layout probe failed the existing
unused-qualification lint; the corrected probe passed and both attempts are
retained. Source hashes are checked across test and build boundaries.

The initial ranking drafts exposed two offline reporting bugs (whole-process
size aggregation under a selected-scope label and insufficient owner matching).
Their outputs and correction receipts are retained as superseded. The v3
ranking recomputes exact owner selection and selected size buckets. The first
0958 correction receipt also had a trace-count transcription error, corrected
from 2,620 to 2,604 in v2; calls and bytes were unaffected.

## ADR review for the isolated candidate

| Constraint | Candidate evidence and adoption condition |
| --- | --- |
| 0001 priorities; 0006 preservation and validation | Existing finite validator, lexical Eq/Hash/Debug/Display and parser refusal paths remain; 1,400 tests pass. Native and memory gates still determine usefulness. |
| 0002/0024 crate ownership | Only three XLSX implementation files change; existing std Arc and itoa dependencies suffice. No facade or archive coupling is added. |
| 0003 immutable snapshots | Shared numeric bytes remain immutable and Send + Sync. Existing clone/patch/inverse behavior is covered by library tests and capture oracles. |
| 0005 measured memory and performance | Actual owner sizes remain unchanged. Arc control blocks, atomic operations and constructor/scanner copies are explicitly retained tradeoffs; allocation savings alone cannot authorize adoption. |
| 0008 verification | Focused and full XLSX tests, Clippy and format pass in the isolated tree. No broad completion or workspace-wide verification claim follows. |
| 0031 explicit execution | Existing finite-context benchmark uses one worker. The candidate adds no runtime, worker, I/O or scheduling behavior. |


## Comparison protocol and capture history

The frozen native schedule contains three cases, six alternating ABBA/BAAB
blocks per case, and 20 warmups plus 200 measured samples per fresh process:
72 children. The allocator lane follows with three alternating AB/BA/AB pairs
per case, two warmups and ten measured rows: 18 children. Workers remain one,
with affinity to CPUs 12–15. The baseline reuses the byte-bound 0957 binaries;
the candidate uses the same Rust 1.95.0 opt-level 3, fat LTO, one codegen unit,
debug level 1 and disabled incremental compilation. RUSTFLAGS are unset.

Edit/commit and total p50/p95 are primary. Publication is a control; all phase
p99 and RSS observations remain visible. Native uncertainty uses 10,000
bootstrap resamples (seed 959) over the six block geometric means, not pooled
iterations. Allocator counters remain descriptive, with region peaks separate
from process RSS. Per-case aggregate upper bounds above 1.05 require review;
all individual leg/block regressions over 5% are retained. A named 5% benefit
requires point ratio at most 0.95 and an upper confidence bound below 1.0.
The selected-cell read control is separately planned and not implied by this
three-case managed edit capture.

Before launch, actual retained native and allocator reports exposed schema
assumptions in both helpers: allocator identity is nested under `case`, and
one reducer scope string was incomplete. Versioned fixes passed all three
case report checks and the offline controls before the final capture plan.
The initial release build also failed because direct RUSTFLAGS combined fat
LTO with disabled embedded bitcode. Its replacement uses the same Cargo
release-profile environment as the retained baseline; the failed attempt
remains recorded.

The first capture epoch (`capture-native-v1`) failed its first baseline child
because the runner precreated the benchmark output directory. The benchmark
refused at its create-new guard before calling `run`, so no fixture or measured
sample ran. The failure terminal and directory remain untouched. The v7 helper
requires an absent output path and lets the benchmark create it. An offline
control proves that path is absent at launch and that an interruption after
parent exit still cleans the process group and records failure. A new GO and
`plan-final-v8.json` authorize `capture-native-v2`; candidate, schedule, gates
and oracles are unchanged. This is a corrected capture epoch, not a replacement
sample inside the failed epoch. No timing was discarded.

## Measured result and decision

Both lanes completed with all semantic/resource validations passing: 72 native
children (14,400 measured iterations) and 18 allocator children (180 measured
iterations). The frozen reducer initially stopped before reading packet data
because it required legacy source-binding field names. A versioned adapter
validates the final plan's actual descriptors and binds the original reducer;
report checks, schedules, statistics and gates are unchanged. The original
failure, helper, correction and control receipts are retained. No capture was
rerun or excluded to repair this offline compatibility issue. Independent
review verifies all 90 child records and reproduces all 63 aggregate metrics
without a discrepancy.

Candidate/baseline ratios below use the six block geometric means; brackets
show the two-sided 95% bootstrap interval. Lower is better. These are managed
serial-route measurements on the retained host/build, not an ordinary-route
or cross-platform claim.

| Case | Edit/commit p50 | Edit/commit p95 | Total p50 | Total p95 | Native RSS |
| --- | --- | --- | --- | --- | --- |
| medium-one-percent | 1.0205 [1.0137, 1.0265] | 1.0189 [1.0127, 1.0247] | 1.0095 [0.9984, 1.0192] | 1.0083 [0.9978, 1.0181] | 0.9956 [0.9740, 1.0171] |
| dense-sparse-one-percent | 1.0303 [1.0198, 1.0402] | 1.0301 [1.0184, 1.0414] | 1.0120 [1.0058, 1.0181] | 1.0119 [1.0062, 1.0177] | 0.9760 [0.9690, 0.9817] |
| medium-one-edit | 1.0395 [1.0276, 1.0521] | 1.0469 [1.0315, 1.0616] | 1.0881 [1.0422, 1.1370] | 1.1019 [1.0640, 1.1454] | 1.0391 [1.0181, 1.0614] |

No primary metric reaches the prespecified 5% useful-benefit threshold. The
one-edit edit/commit, total and RSS upper bounds exceed 1.05, so its primary
safety gate fails. Its total p99 ratio is 1.1039 [1.0674, 1.1438]. The medium
and dense primary safety gates pass, but their edit/commit point estimates are
about 2.1% and 3.0% slower rather than useful improvements.

Commit-drop p50 rises 11.9%, 16.3% and 7.3% for medium-percent, dense-percent
and one-edit respectively. Publication p50/p95 ratios are 0.9896/0.9885,
0.9806/0.9821 and 1.1530/1.1851. The one-edit publication control is variable
and regresses; these observations do not isolate atomic operations as the
cause. Every phase p99, individual pair/block flag and confidence interval
remains in the [bounded machine-readable results](results/change-0959/analysis/comparison-summary-v3.json). No grand average hides them.

Allocator ratios describe the drained core and its measured region peak;
they are separate from native process RSS and lifetime high-water marks.

| Case | Allocation calls | Requested bytes | Region peak live bytes |
| --- | ---: | ---: | ---: |
| medium-one-percent | 0.9136 | 1.0052 | 1.0068 |
| dense-sparse-one-percent | 0.9120 | 1.0087 | 1.0112 |
| medium-one-edit | 0.9180 | 1.0064 | 1.0027 |

The allocation-call reduction is real in this instrumented route, but requested
bytes and region peaks increase slightly. Passing allocator safety does not
override the native result. The candidate is therefore **not adopted**. Its
patch and isolated source remain reproducible evidence; production keeps
`Number(Box<str>)`. The separately planned selected-cell control was not run,
because the primary trial already fails. Public boxed-constructor and scanner
tradeoffs remain unqualified, not implicitly cleared by the semantic tests.

The next measured opportunity is the larger parser backing-allocation owner,
starting with a bounded capacity-hint hypothesis and full error/semantic
checks. This ranking does not establish a latency benefit or authorize removing
complete source validation. Existing 0956/0957 memory flags and the earlier
ordinary one-edit p99 flag remain open.

## Reproduction and retained evidence

The [packet README](results/change-0959/README.md) contains source/build
bindings, the retained candidate patch, capture and reduction commands,
per-case statistics, review receipts and external artifact descriptors. The
full comparison remains outside Git at
`/home/zhuhe/litchi-performance-0959/analysis/comparison-v3.json`
(1,447,293 bytes; SHA-256
`9dbaefcd25b348c6b3f3c23ff4ed94b6e160fa7c95a5e37667a7efafa80242eb`).
All large source manifests, raw captures and executables remain external.
Owned build caches and scratch were removed after reduction: 4,542 files,
4,079,318,512 logical bytes. Retained sources, binaries and replay evidence
were verified unchanged. No push was performed.
