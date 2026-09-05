# Change 0417: representative CRUD timing and resource baseline

> Correction (0421): allocator `peak_live_bytes_*` values in this record
> may under-report peaks and must not support high-water claims.
> [Counter correction](0421-allocator-peak-counter.md) explains the defect.
> Live-byte totals, allocation requests, normal timing, RSS and Heaptrack
> metrics are unaffected by this specific bug. Raw reports are unchanged.

`performance_claim: none`

`claim_authorized: false`

## Problem and result

Recent performance work concentrated on narrow container paths. This batch
establishes a current, descriptive baseline for all 30 existing representative
selectors in the non-iWork CRUD index, spanning 14 executable categories.
Dynamic calculation/refresh remains unsupported. The measured revision is
`b10d6c25a13242ca260a8c897946f4d80ae06c61`; no production runtime code changes
in this batch.

The [evidence bundle](../results/change-0417/) retains 60 normal processes
(30,000 observations), 60 separate allocator processes (1,800 observations),
and 30 one-sample preflight processes. Every report has its original corpus
catalog and whole-process time/RSS log. The new reusable summarizer validates
raw vectors, producer statistics, identities and oracles, and the bundle
verifier checks matrix completeness, exact commands, ordered non-overlapping
runs, historical input hashes and portable replay.

The XLS validation corpus contract now lists its supported `tiny` and `large`
writer shapes; it previously incorrectly included `medium`. A regression test
checks that boundary. The captured XLS input was `large`. The index's 10
measured and 20 correctness-only statuses, 425-selector registry, and
36-case / 198-row default matrix remain unchanged. This separate opt-in
baseline does not satisfy the default/full-run acceptance contract or establish
complete CRUD lifecycle coverage.

## Measurement scope

CPU 2, one configured worker, generated warm in-memory sources, and a shared
AMD EPYC 9R45 KVM host are fixed by the protocol. Two normal fresh processes
per selector retain 500 samples after 20 warmups. Allocator processes retain
30 samples after 3 warmups. Repeat 1 follows matrix order and repeat 2 reverses
it. Task-owned builds, tests, profiling and measurement are serialized;
unrelated host background activity is uncontrolled. Samples share their process.

The [timer audit](../results/change-0417/timing-boundaries.md) identifies each
case's included phases. Already-open query and commit-only cases cannot stand
in for open/edit/save lifecycles. `CountingSink` retains full output; 64 KiB
limits individual writes. The RTF 37-byte and XLSX 4 KiB authoring windows do
not bound whole-process memory. Join and three-way branch preparation can
use threads outside the elapsed interval despite the one-worker configuration.

The release build explicitly used Rust/Cargo 1.98.1 with frame pointers and
unwind tables. Embedded binary compiler metadata confirms Rust 1.98.1. Raw
reports say `rustc_version: 1.95.0` because the harness invokes the pinned
checkout's runtime compiler command when collecting environment metadata.
That runtime field does not identify the build compiler; both records are
retained without rewriting reports.

## Timing, uncertainty and allocation gaps

The [complete table](../results/change-0417/baseline-table.md) retains both
repeats' p50/p95/p99, scoped operations/second and process peak RSS. The
summary retains IID median order-statistic intervals, integer-nanosecond
midpoint p50 and nearest-rank tails. IID intervals do not account for host
drift; repeat observations are not pooled.

Media-rich PPTX copy has p50 **1,151.843 / 1,152.214 ms**. Defined-name and
sheet-protection eager publication have p50 near **236 ms**. These figures
describe different workloads and timing boundaries; their ratios are not
speedups or a comparison of interchangeable APIs.

Five selectors exceed the predeclared absolute 5% repeat-drift flag:

- ODF validation p99: **−6.67%**.
- ODP text-to-sink p99: **+5.86%**.
- XLS validation p50/p95/p99: **+11.55% / +11.53% / +11.58%**.
- XLSX defined-name publication p99: **+86.32%**.
- XLSX sheet listing p99: **+9.09%**.

These observations remain visible; flagged quantiles do not establish a
stable baseline or a regression cause. No speedup is claimed.

Only two selectors emit allocation attribution. Both repeats report identical
p50/p95/p99 calls and bytes: RTF streaming **16,387 calls / 1,507,536 bytes**;
DOCX story hyperlink redaction **4,518 calls / 8,548,090 bytes**. DOCX's allocation
region includes package cleanup after elapsed timing stops. The other **28
selectors remain unavailable**, including media-rich PPTX and XLSX streaming.
Whole-process RSS and lifetime allocator peak snapshots do not fill this gap.
Instrumented allocation elapsed values do not enter a latency comparison.

## Validation and remaining work

The separate PPTX CPU diagnostic retains 33,347 user-cycle stacks with zero lost
samples and 0.479% unresolved leaf weight. Deflate accounts for **71.929%** of
whole-command leaf weight and SHA-256 for **20.66%**. Call chains connect
compression to candidate building, commit re-planning, post-apply physical
fingerprinting and publication. This confirms substantial compression work;
it does not give elapsed-time shares for the timed phases.

The separate whole-command PMU observation has approximately 83% counter
scheduling coverage, IPC **2.384**, and branch misses **2.301%**. L1 alias zeroes
remain unvalidated and LLC events unsupported. Setup, preflight, warmups,
verification and reporting descendants are included. Compressed raw traces,
symbolized stacks, counter CSV and replayable attribution are retained.

All 120 formal report/catalog pairs and 30 preflight pairs pass the full bundle
verifier. All 145 focused Python tests pass, including summary corruption and
copied-path replay tests, coverage-index contracts, corpus bindings and the
existing comparator suite. Coverage-index, crate-boundary and report-claim
classification checks pass. Initial wrapper path and index-edit failures are
retained with the final successful logs.

After removing the task worktree and both binaries, an isolated export replays
successfully twice and rejects 12 deliberate corruptions, including source and
binary identities, ordering/overlap, corpus inputs, raw statistics, summary data,
duplicate keys and invented zero allocation. Shared build targets and the
user's untracked `docs/GOAL.md` are preserved.

The [static PPTX audit](../results/change-0417/pptx-static-audit.md) identifies
repeated image compression in planning, commit checks and publication, and
an existing source-backed media-copy capability for a matched next experiment.
Preservation and source-identity checks must remain equivalent before removing
work. The plain source-backed case in this baseline is not a matched media-rich
counterpart.

The reusable verifier and XLS shape correction are committed at `4d5cdcdeb`.
Measured runtime identity remains the earlier clean `b10d6c25a` revision.

The program remains open: operation allocation coverage, aligned end-to-end
boundaries, broader sizes and producers, cold/remote sources, failure matrices,
native interoperability, and concurrency/scaling evidence are still required.
