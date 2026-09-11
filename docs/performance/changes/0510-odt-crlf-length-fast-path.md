# 0510: ODT XML 1.0 CRLF length fast path

The ODT semantic text parser's XML 1.0 decoded-length helper now retains
UTF-8 validation, uses one `memchr` search for the first carriage return, and
returns the raw length when no carriage return is present. When a carriage
return is present, it starts the existing scalar CRLF walk at that first byte.
This preserves the established normalization, error, budget, writer and
progress boundaries while avoiding the full byte walk for the common no-CR
case. It adds no public API, dependency, archive, or iWork change.

The candidate is retained with documented newline-workload tradeoffs.
The matched ordinary export medians improve 3.92–8.87%, while the longer
newline-heavy guardrail retains a dense-small median +6.86% flag and a
sparse-medium p99 +14.11% flag. All individual results remain visible below.
This is an export-only result on synthetic corpora, with no universal
newline, peak-memory, full CRUD or cross-format improvement claim.

## Selected algorithm and diagnostic limits

The first-CR hybrid validates `raw` with `str::from_utf8` before searching.
After validation, a missing carriage return returns `raw.len()` exactly. The
first carriage return is a valid byte boundary, so the unchanged scalar loop
can preserve every later CRLF subtraction and lone-CR result. The helper has
no access to the sink or pending block frontier.

Two unconditional iterator alternatives were rejected in the isolated helper
probe:

| Diagnostic variant | Worst helper p50 regression | Case |
| --- | ---: | --- |
| `memchr_iter` | +672.3% | 65,536-byte all-CR input |
| `memmem_iter` | +754.3% | 65,536-byte dense-CRLF input |

The selected hybrid has a remaining +33.3% tiny dense-CRLF flag and +8.4%
large sparse-CRLF flag in that probe, while its no-CR cases improve
substantially. Those results are diagnostic aggregates only: the retained
helper CSV stores per-row min/p50/p95/max values rather than every sample
duration, and per-call durations use integer division of batched elapsed time.
One 49-byte sparse fixture has no CRLF and duplicates its no-CR control. These
figures cannot replace the full public guardrail or be used for a cross-probe
percentage comparison.

## Public newline guardrail

The required end-to-end guardrail is a separate public-export probe over nine
newline cases (dense, all-CR and sparse CRLF at 49, 1,024 and 65,536 bytes).
It uses an independent byte-counting `Sink`, keeping its timing scope
separate from the hashing baseline. Before editing, the control freezes exact
preflight output bytes, object counts, archive digests and output digests;
the candidate has a matching exact preflight. The retained preflight rows
match on those semantic identities.

The initial capture runs 500 samples per case with ten warmups, serially in
before/r1, after/r1, after/r2, before/r2 order on CPU 2. Opening and exact
byte checks remain outside the timer, and hashing is not timed. A posthoc
1,000-sample-per-case ABBA follow-up examines the initial adverse and drift
flags using the same frozen binaries. Both sets remain in the evidence;
no runs are discarded.

## Current instrumented profile

The control and candidate use the same ODT large semantic export shape. The
Callgrind collection is toggled only inside `write_text_blocks_to_writer`,
including the checking SHA sink, for five exports with no warmups. Heaptrack
runs the whole child for 20 exports. The retained comparison is:

| Measure | Control | Candidate | Scope and interpretation |
| --- | ---: | ---: | --- |
| Inclusive Callgrind instruction references | 298,831,445 | 275,608,266 | **-7.77%**; simulated references in the parser scope |
| `normalized_xml10_decoded_len` exclusive references | 28.35M | 1.6M | Helper attribution in the same Callgrind scope |
| `normalized_xml10_decoded_len` inclusive references | 35.5375M | 12.2875M | Helper plus its callees in the same scope |
| `append_sink_precharged` allocation calls | 20 | 20 | 20 large exports; the 0509 reusable-buffer count is unchanged |
| Whole-child peak heaptrack heap | 9.91M | 9.91M | Rounded display, unchanged; no peak-memory claim |

The Callgrind values are simulated instruction references, not hardware
cycles or native elapsed time. The allocation count is stack-filtered and
does not represent total process allocations. Whole-child heaptrack includes
fixture generation, preflight, opening, validation and report writing; its
RSS also includes profiler overhead. The harness, parser writer boundary,
checking sink and document-open/setup scope remain unchanged.

## Correctness and resource boundaries

The source review finds no source-level correctness blocker. UTF-8 validation
and its existing typed error mapping remain before the fast path. Valid text
still computes normalized length, charges the one cumulative
`SinkTextBudget`, materializes through the existing XML 1.0 decoder, and then
passes through the existing fallible append/precharge checks. Depth, block,
space and decoded-byte limits remain in their existing paths.

The focused differential cases compare the hybrid with the prior scalar helper
for empty, no-CR, lone-CR, repeated and mixed endings, Unicode, invalid UTF-8,
and generated valid inputs. Integration cases cover both text and CDATA,
owned and source-backed output, exact normalized output limits, and progress
after a refusal. These checks were inspected by the source and ADR reviews;
all 1,495 Rust tests/doctests pass, with one existing producer test ignored.

The helper does not touch `PendingSinkBlocks`, writer calls, nested frontiers,
source bytes or parser traversal. XML 1.0 normalization remains the existing
semantic behavior, including CRLF length subtraction and lone-CR handling.
The isolated helper flags therefore remain a workload limitation rather than
a license to claim a broad newline optimization.

## Matched native results and admission

The unchanged full-feature baseline uses its hashing discard sink inside the
export clock, with document opening outside it. Each primary child runs
500 samples and ten warmups for each of tiny, medium and large, serially in
ABBA order on CPU 2. All six paired identities match. The 6,000 raw durations
independently reproduce the statistics; primary repeat drift stays within the
predeclared ceilings. The host is shared: this batch ran its captures serially
after its own compilation finished, but does not establish system-wide idle
isolation or a cause for the observed drift. Percentage changes below are
candidate versus control.

| Repeat | Shape | Control p50 ns | Candidate p50 ns | p50 | Mean | p95 | p99 | Throughput |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| r1 | tiny | 6,140 | 5,890 | -4.07% | -3.96% | -2.56% | -3.11% | +4.12% |
| r1 | medium | 38,271 | 35,910 | -6.17% | -6.05% | -4.34% | -4.83% | +6.44% |
| r1 | large | 1,865,142 | 1,699,761 | -8.87% | -8.93% | -8.71% | -14.89% | +9.80% |
| r2 | tiny | 6,130 | 5,890 | -3.92% | -1.73% | -2.88% | -2.80% | +1.76% |
| r2 | medium | 38,020 | 35,525 | -6.56% | -6.56% | -5.69% | -4.63% | +7.02% |
| r2 | large | 1,818,927 | 1,692,306 | -6.96% | -6.98% | -6.80% | -10.51% | +7.50% |

The separate public counting-sink guardrail retains 18,000 initial and 36,000
follow-up raw samples, with the same nine archive/output identities and exact
byte checks. Its timing scope differs from the primary hashing baseline;
percentages must not be combined into a cross-sink aggregate.

| Capture | Newline pattern–paragraph bytes | p50 | Mean | p95 | p99 | Throughput | >5% adverse flag |
| --- | --- | ---: | ---: | ---: | ---: | ---: | --- |
| r1 | all_cr-1024 | -0.48% | -0.45% | -0.45% | -0.50% | +0.45% | no |
| r1 | all_cr-49 | +1.21% | +1.13% | +0.88% | -1.89% | -1.11% | no |
| r1 | all_cr-65536 | +3.14% | +3.18% | +3.56% | +0.09% | -3.08% | no |
| r1 | dense_crlf-1024 | -2.52% | -2.54% | -2.45% | -0.55% | +2.61% | no |
| r1 | dense_crlf-49 | +2.30% | +2.41% | +3.91% | -1.85% | -2.35% | no |
| r1 | dense_crlf-65536 | -0.23% | -0.20% | +0.44% | +1.11% | +0.20% | no |
| r1 | sparse_crlf-1024 | -6.83% | -6.07% | -2.76% | +8.33% | +6.46% | yes |
| r1 | sparse_crlf-49 | -1.50% | -1.58% | -3.03% | -2.47% | +1.60% | no |
| r1 | sparse_crlf-65536 | +38.39% | +38.35% | +34.63% | +33.30% | -27.72% | yes |
| r2 | all_cr-1024 | +0.46% | +0.57% | +0.54% | +5.60% | -0.57% | yes |
| r2 | all_cr-49 | +1.32% | +1.37% | +1.51% | +1.93% | -1.35% | no |
| r2 | all_cr-65536 | +0.11% | +0.12% | +0.22% | -0.26% | -0.12% | no |
| r2 | dense_crlf-1024 | +0.49% | +0.67% | +1.06% | +1.11% | -0.67% | no |
| r2 | dense_crlf-49 | +2.11% | +2.17% | +2.19% | +8.34% | -2.12% | yes |
| r2 | dense_crlf-65536 | -0.06% | +0.04% | +0.02% | +0.10% | -0.04% | no |
| r2 | sparse_crlf-1024 | +0.91% | +2.49% | +19.66% | +35.20% | -2.43% | yes |
| r2 | sparse_crlf-49 | -1.52% | -1.66% | -4.14% | -3.60% | +1.69% | no |
| r2 | sparse_crlf-65536 | -7.97% | -8.64% | -18.15% | -18.09% | +9.46% | no |
| followup-r1 | all_cr-1024 | +0.46% | +0.24% | -1.14% | -4.63% | -0.23% | no |
| followup-r1 | all_cr-49 | +1.58% | +1.52% | +1.37% | -1.31% | -1.50% | no |
| followup-r1 | all_cr-65536 | +0.33% | +0.26% | +0.30% | -4.93% | -0.26% | no |
| followup-r1 | dense_crlf-1024 | +0.62% | +0.62% | +1.12% | -0.40% | -0.61% | no |
| followup-r1 | dense_crlf-49 | +6.86% | +6.73% | +6.61% | +1.56% | -6.31% | yes |
| followup-r1 | dense_crlf-65536 | +0.23% | +0.16% | +0.26% | +0.20% | -0.16% | no |
| followup-r1 | sparse_crlf-1024 | +0.57% | +1.17% | +1.07% | +14.11% | -1.16% | yes |
| followup-r1 | sparse_crlf-49 | +3.47% | +3.39% | +3.34% | +0.89% | -3.28% | no |
| followup-r1 | sparse_crlf-65536 | +2.08% | +2.09% | +2.16% | -0.54% | -2.05% | no |
| followup-r2 | all_cr-1024 | -0.29% | -0.55% | -0.61% | -4.60% | +0.55% | no |
| followup-r2 | all_cr-49 | +0.14% | +0.17% | +0.33% | +0.29% | -0.17% | no |
| followup-r2 | all_cr-65536 | -0.68% | -0.70% | -0.72% | -0.58% | +0.70% | no |
| followup-r2 | dense_crlf-1024 | -1.05% | -0.98% | -0.12% | +1.39% | +0.99% | no |
| followup-r2 | dense_crlf-49 | +2.72% | +2.59% | +1.70% | +2.61% | -2.52% | no |
| followup-r2 | dense_crlf-65536 | -1.21% | -1.16% | -1.08% | -0.82% | +1.18% | no |
| followup-r2 | sparse_crlf-1024 | -2.83% | -3.10% | -2.63% | -2.74% | +3.20% | no |
| followup-r2 | sparse_crlf-49 | -0.70% | -0.69% | -0.44% | -1.99% | +0.69% | no |
| followup-r2 | sparse_crlf-65536 | -2.03% | -2.03% | -2.22% | -7.30% | +2.07% | no |

The initial sparse 65,536-byte R1 median +38.39% reverses to −7.97% in R2,
with candidate repeat drift −38.48%. Its longer follow-up gives +2.08%/−2.03%
and no follow-up drift flags. This does not establish the cause of the
initial instability, which remains retained. The initial sparse 1,024-byte
candidate tail drift and all initial tail flags are also preserved.

The follow-up still has a dense 49-byte R1 p50 +6.86%, mean +6.73%, p95 +6.61%
and throughput −6.31%; its R2 median rises 2.72%. This artificial workload
contains 10,000 paragraphs consisting of CR/LF bytes. That measured cost is
accepted in exchange for the repeatable ordinary-export gain and instruction
reduction. The sparse 1,024-byte follow-up R1 p99 rises 14.11%, versus −2.74%
in R2. This is not a regression-free or newline-tail improvement claim.
The [admission disposition](../results/change-0510/acceptance.json) enumerates
every guardrail flag and is checked against recomputed results.

## Memory and hardware scope

Primary whole-child RSS is 30,068→30,068 KiB and 30,336→30,004 KiB. It includes
all three shapes and setup. Initial guardrail RSS is 6,208→6,208 KiB and
6,196→6,208 KiB; follow-up is 6,272→6,208 and 6,236→6,208 KiB. It includes
nine fixtures, opening and retained exact-preflight output. No RSS pair has
an adverse >5% flag, and there is no operation-local peak-memory claim.
Whole-child heaptrack allocation calls are 362,176→362,178; the two-call
setup/report-level difference does not change the stack-filtered one buffer
allocation per large export. Rounded peak heap remains 9.91M.

Hardware access became available (`perf_event_paranoid=1`). Four serial ABBA
`perf stat` children each perform 1,000 large exports with zero warmups.
Grouped cycles, instructions, branches and branch misses have identical event
runtime and 100% running time. Whole-child cycles fall 7.29%/7.88%, instructions
12.67%/12.50%, and branches 12.25%/12.06%. Branch misses change +0.88%/−25.36%;
context switches change 70→76 and 77→74, and migrations remain zero.
Counters include fixture creation, opening, validation, exports and JSON
reporting, so these are not operation-local counter percentages.
Cache events are excluded because the broad capability probe was unreliable.

A 499 Hz DWARF cycle profile of 1,000 exports per binary records zero lost
samples. The flat control helper self share is 10.35%; the candidate helper
is below the report's 1% cutoff. Optimized stack unwinding and kernel symbol
restrictions limit call-chain interpretation. Retained flat samples are
coarse attribution only, not a precise serial fraction or scaling claim.
All 6,050 instrumented export samples (hardware, sampling, heaptrack and
Callgrind) are excluded from native latency comparisons.

## Custody, checks and cleanup

The control base is `869bbc3ed470453b8324893180360a2d69e7c315`. Its full harness
binary SHA-256 is `4808a9d6e1d0f1c5c1aa903f1c88ff7713d41584a1a920336e2d99c862ebeb24`;
the candidate is `6e25e32d926756354de400bd089ea3ecdb214e338bbf8831d452ad7305407eab`.
The final source manifest is
`6f453fbf386665b174abc108b4a60e35fdc34922782fdc4dbd57af486dc2356a`.
Only the private ODT helper and its two test locations change. Direct-rustc
public guardrail binaries separately bind the same owner rlibs and probe;
they are not the full harness binary.

Formatting, 1,012 ODT tests/doctests, 483 harness tests, strict scoped Clippy,
strict rustdoc, crate boundaries and strict claim validation pass. One
existing opt-in producer test remains ignored. All 29 accepted ADRs plus the
README are hash-checked unchanged. The 41-case/213-row/43-corpus default,
18 correctness-only mapped selectors and ten-entry claim registry remain
unchanged. No native-application or broad concurrency coverage is added.
The [replay verifier](../results/change-0510/verify.py) checks source/binary,
corpus/output, raw statistics, negative truncated captures, flags and gates.
Owned temporary build/probe files are removed with a retained
[cleanup receipt](../results/change-0510/cleanup.json); profiles, raw durations,
commands and logs remain under the evidence directory and its SHA256SUMS.

## Follow-up priority

After this batch, the requested optimization priority is OLE2 and OOXML.
Further ODF optimization work is deferred until that optimization goal is
complete.

The retained [profile summary](../results/change-0510/profile-summary.json),
[admission record](../results/change-0510/admission.json),
[source review](../results/change-0510/source-review.md),
[ADR review](../results/change-0510/adr-review.md),
[helper guardrail](../results/change-0510/helper-guardrail/README.md),
[public guardrail capture](../results/change-0510/export-guardrail/capture.py),
[control preflights](../results/change-0510/export-guardrail/before-preflight.csv),
and [candidate preflights](../results/change-0510/export-guardrail/after-preflight.csv)
provide the evidence and scope for this retained change.
