# 0688 — inline checked CFB chain traversal

Status: final validation passed; scoped retention independently reviewed. `performance_claim: none`.
OLE2/OOXML remain active; iWork is excluded.

## Measured hypothesis

After [0687](0687-cfb-ascii-directory-keys.md), a fresh two-million-query
profile assigns 59.07% self samples to shared CFB `next_chain_sector` on the
54016 owned selected-cell route. Disassembly confirms two actual calls from
`stream_cursor_at_hinted`. The 261-byte helper reserves 96 stack bytes and
writes its successful `Result` through a hidden return pointer; the caller
reloads its discriminant and value. This establishes repeated call and result
handling work, beyond the necessary dependent FAT/MiniFAT table load.

## Implementation and correctness

The private checked helper is marked inline. Two private cold, non-inline
helpers retain marker/index error formatting. Current-marker validation,
explicit `usize::try_from`, bounds-checked lookup, and next-marker validation
remain in exactly that order. A current `ENDOFCHAIN` is still an invalid marker;
a next `ENDOFCHAIN` still succeeds. A regular next value outside the table is
returned and fails only on a following lookup. Error variants and text remain
unchanged. The cursor loop and legacy `file.rs` helper are unchanged.

Two focused tests compare exact values and typed errors against frozen
pre-change helpers across small tables, marker boundaries, missing indices,
short walks, and arbitrary resume states. Independent literal assertions pin
error precedence, termination, delayed out-of-table failure, and zero/backward
walk behavior. Existing hinted/unhinted and checkpoint corruption/source-fence
tests continue to run. No cache, persistent state, public API, dependency,
unsafe code, or source-read policy is introduced.

## Evidence scope

The [packet](results/change-0688/README.md) binds baseline `9df782b7c`, final
sources, unchanged probes and fixtures, binaries, commands and raw captures.
Native A/A then A/B/B/A cover 24 case/source groups, 14,400 fresh owners and
115,200 selected queries. The eighth query and sum of open plus eight query
timers are distinct metrics. Separate allocation/counting probes preserve all
96 allocation groups and 12 source-I/O routes. Long-loop controls use nine
process samples per leg, each averaging 50,000 queries with the default 2 MiB
limit. They are distributions of loop means, not individual query latencies.

Hardware-counter diagnostics subtract N=10 from N=100,010 and divide by
100,000 extra queries; setup and the timing wrapper remain in whole-process
counts. Native-child RSS is captured separately from allocator gauges. All
measurements use CPU 12 on the recorded shared host and warm OS file caches.
No physical cold-device, remote, concurrency, cross-platform, native Office,
malformed-chain performance,
clean-build, universal memory-bound or general XLS speedup claim follows.

## Paired results and tradeoffs

The eighth-query median changes are:

| Target | Owned | File, warm OS cache |
|---|---:|---:|
| 54016 first, default index budget | −11.51% / −12.50% | −4.88% / −4.21% |
| 54016 late | −14.77% / −17.32% | −8.16% / −9.80% |
| Plan1 first | −10.29% / −8.70% | −1.77% / −0.89% |
| Plan1 late | −8.11% / −13.16% | −3.21% / −3.18% |
| 45365-2 late | −15.38% / −15.38% | −2.17% / −3.28% |
| Formula refusal | −10.34% / −9.24% | −8.15% / −6.37% |

For example, 54016 first owned falls from 1,260/1,280 ns to 1,115/1,120 ns;
its late target falls from 1,760/1,790 ns to 1,500/1,480 ns. The separate
50,000-query controls confirm 54016 first owned loop means improving
13.50–14.85%, late owned 15.35–15.56%, late file 9.30–9.54%, and Plan1
owned 10.92–11.08%. Missing indexed-query loop means remain approximately
flat. There is no aggregate average that hides individual routes.

Large indexed 54016 open-plus-eight windows remain approximately flat
(owned first +0.30%/+0.14%, late −0.14%/−0.14%). Formula-refusal windows
improve 4.27–4.77% owned and 3.77–4.16% file. The 45365-2 first-target
windows improve 0.32–3.55% owned and 2.49–5.66% file; its file A/A control
moves −3.07%, limiting interpretation of the smaller workflow changes.
No general opening or short-workflow gain is claimed.

Costs remain visible. With indexing disabled, 54016 owned q8 regresses
4.34–4.68% and open-plus-eight 3.77–4.72%; corresponding file windows rise
1.15–3.18%. Simple stored owned q8 rises 2.13–4.35%, while its longer loop
means rise 3.27–3.68%; file q8 rises 2.40–4.39%. The Simple missing owned
q8 median is 70 → 80 ns in both pairs (+14.29%), despite that route doing
no chain lookup. This is a retained observed cost; the experiment does not
isolate its cause. Timer quantization does not justify dropping it.

The [median triggers](results/change-0688/regressions.md) and
[mean/tail triggers](results/change-0688/tail-regressions.md) retain all paired
changes over 5%. These include Simple stored owned warm-query means and
workflow p99, missing-query tails, and several open/first-query p99 flags.
The full comparison retains single-leg changes, 1,000-draw paired median
bootstrap intervals and A/A/ABBA drift. Shared-host drift remains possible,
and 100-sample p99 is descriptive, not a population tail guarantee.

## Mechanism and resource costs

The candidate has no standalone `next_chain_sector`. The compiler instead
outlines the whole `cursor_chain_sector` walk: one call per walk, with all
checks inline in the successful link loop. That loop contains neither a
per-link call nor a successful `Result` store. It still performs the table
load, bounds and marker checks. The whole-walk result still uses the existing
return ABI. The source-level cursor loop was not rewritten.

In the measured binary, cursor construction shrinks 1,201 → 958 bytes; the
new walk is 389 bytes and the two cold helpers are 101 bytes each. Baseline's
standalone link helper was 261 bytes. Whole-binary `.text` grows 636,083 →
637,523 bytes (+1,440, about 0.23%); local cursor shrinkage is not a claim of
whole-program code shrinkage. Cursor stack reservation changes 168 → 88 bytes;
the candidate walk reserves 120 bytes, while baseline's per-link helper
reserved 96. These are disassembly observations, not a stack-usage bound.

The matched two-million-query profile takes 2.224 → 1.907 seconds (−14.25%).
Candidate `cursor_chain_sector` still accounts for 56.65% self samples.
Attribution moves with inlining, so percentages alone do not establish an
absolute work reduction. Separate whole-process diagnostics do:

| Owned target | Instructions per extra query, before → after | Cycles, before → after |
|---|---:|---:|
| 54016 first | 25,579 → 15,790 | 4,907 → 4,196 |
| 54016 late | 45,448 → 24,594 | 7,614 → 6,464 |
| Plan1 first | 11,569 → 9,559 | 2,556 → 2,293 |
| Simple first | 5,862 → 5,853 | 1,416 → 1,484 |

Instructions fall 38.27% and 45.89% on the first/late 54016 owned routes,
with cycles falling 14.48% and 15.10%. Simple's cycles rise 4.86%, consistent
with its measured small-route cost. Branches fall on the long-chain routes;
branch misses, cache misses and page faults remain in the raw comparison.
Very small differenced miss/fault counts are not used to claim locality gains.

All 96 allocation groups match exactly across three repeats per binary:
calls, requested bytes, peak live delta and retained delta. All 12 counted
routes preserve opening and eight queries' reads, bytes, freshness observations
and outcomes. These query allocation captures exclude opening and stack usage.
Native-child RSS has no paired median increase above 5% in this diagnostic
matrix; the largest is Simple owned long at 2,564 → 2,680 KiB (+4.52%),
followed by Plan1 owned long at 3,000 → 3,128 KiB (+4.27%). No uniform RSS
reduction or process-memory bound is claimed.

## Validation and disposition

Six CFB/XLS/facade gates and DOC/PPT consumer tests pass: 4,394 tests passed,
zero failed, 27 existing ignored. The two new boundary tests are included.
All 126 real XLS fixtures match exactly on owned/file sources; the generated
70,001-cell fixture retains full-visitor count and semantic-digest parity.
Formatting, all-target checks, warning-denied Clippy/rustdoc, boundary and
evidence gates are retained in the packet. No new fuzz campaign or native
Office round trip is claimed.

Independent source review finds no correctness blocker. Performance retention
is scoped to repeated checked-chain queries and selected refusal routes, with
small/no-index costs and code growth disclosed; final independent disposition
is recorded in the [review](results/change-0688/review.md).
Remaining dependent chain traversal, within-sheet/SST work, small-route costs,
broader sources and concurrency remain open. The broader GOAL remains active;
no registered performance claim or CRUD coverage row is promoted.
