# 0509: bounded ODT sink buffer reuse

The ODT semantic text sink now keeps one operation-local spare `String` for
completed paragraph buffers. A subsequent paragraph start can take that
spare after the previous paragraph has been written successfully. The spare is
retained only when its actual capacity is at most **4,096 bytes**; a larger
active paragraph is still allowed, but its allocation is dropped after
emission. This preserves the parser's existing text budget, frontier ordering,
writer boundary, typed failures and progress behavior. It adds no public API,
dependency, archive, or iWork change.

The fresh retained profile supports a bounded allocation claim for the ODT
large semantic export. Native timings show modest median reductions, with a retained initial
p99 flag and a longer follow-up below the review threshold. It makes no
hardware-counter or general peak-memory improvement claim.

## Implementation and semantic boundaries

`PendingSinkBlocks` owns the spare only for the current parser operation. A
completed block is recycled after its existing ordered
`write_object(TextObjectKind::Paragraph, ...)` call succeeds and after the
checked frontier increment. A sink, output-byte, or object-limit error returns
before recycling the failed value, so the existing accepted-byte and completed
object progress remains observable at the same boundary. Nested blocks still
publish in start order; a nested completion cannot become a spare before its
outer frontier is emitted, and a later smaller nested spare cannot replace the
larger outer spare.

Reuse does not make the 4,096-byte value a text or allocation limit. The active
block can grow beyond the retention cap, while the monotone `SinkTextBudget`
continues to charge every decoded contribution against the existing cumulative
limit. Capacity is checked before clearing; a short string with an oversized
allocation is therefore discarded. The spare is not global, thread-local, or
shared between documents. Self-closing empty blocks continue to use the existing empty
string path.

The candidate source review covers successful reuse, actual-capacity rejection,
nested-frontier deferral, writer failure, exact owned/source-backed output,
empty blocks, oversized content, output/object limits, malformed tails and
source staleness. It identifies two modest fixture gaps: no dedicated test
drives the 64 MiB decoded-text ceiling across several blocks while a spare is
present, and the 4,096-byte equality boundary is not a separate fixture. The
implementation uses `>` for the latter, with tests below and one byte above
the bound. The source review was read-only and did not execute these checks;
the executed gate results are recorded below.

## Fresh profile and admission evidence

The control is the byte-identical 0508 final binary. Heaptrack runs the whole
child for 20 large ODT exports, with the allocation stack filtered through
`append_sink_precharged`. Callgrind collects inclusive parser references only
inside matched `write_text_blocks_to_writer` calls for five large exports. The
retained comparison is:

| Measure | Control | Candidate | Scope and interpretation |
| --- | ---: | ---: | --- |
| `append_sink_precharged` allocation calls | 200,000 | 20 | 20 exports of 10,000 paragraphs; 10,000 calls/export becomes one call/export |
| Inclusive Callgrind instruction references | 304,455,496 | 298,891,534 | Five large exports; **-1.83%** simulated instruction references |
| Whole-child peak heaptrack heap | 9.91M | 9.91M | Rounded profiler display; no peak-memory reduction claim |
| Retained spare capacity | — | 4,096 bytes maximum | Actual `String::capacity()`, operation-local |

The allocation reduction is **-99.99%** for this filtered stack. It is not a
whole-process allocation total: whole-child heaptrack also includes fixture
generation, preflight, document opening, validation and report writing. The
Callgrind values are simulated instruction references, not hardware
instructions or cycles. A `perf stat` capability probe was denied, so no
hardware counters, cache, branch, memory-bandwidth, or operation-local RSS
result is available. Heaptrack RSS includes profiler overhead, and the rounded
9.91M displays do not establish exact byte equality.

The unchanged full baseline harness, deterministic ODT semantic generator and
caller-owned checking sink remain in use. The candidate does not change the
sink write boundary or parser open/setup behavior. The native timing
protocol keeps document generation, preflight and opening outside the export
clock; instrumented heaptrack and Callgrind runs are excluded from any native
latency comparison.

The corrected retained strict claim check reports ten performance claims
validated in `claims-final.log`. The registry remains unchanged; this new scoped record is not an added
registry entry.

## Native timing and tail review

The full unchanged harness was built with pinned Rust 1.95.0, release mode,
no debug information/incremental compilation, and two build jobs. The control
binary SHA-256 is `ab874929448dfa3a4c17fd9072e702081b4e78b84cae39f3be120518a0c7f6f2`;
the candidate is `4808a9d6e1d0f1c5c1aa903f1c88ff7713d41584a1a920336e2d99c862ebeb24`.
Each manifest binds 7,195 Rust/TOML/lock files. Only the ODT parser and its
integration test file differ. All gates preceded native captures, without
subsequent source changes.

The initial serial A1/B1/B2/A2 matrix uses ten warmups and 500 measured
samples for each of tiny/medium/large, totaling **6,000 native samples**.
Children are pinned to CPU 2 on the shared AMD EPYC 9R45 host; there is no
isolated-core claim. Opening, fixture generation and preflight are outside
the clock. The bounded hashing discard sink updates its digest inside the
clock; final digest/progress checks are outside. All paired corpus hashes,
output digests, sink call counts and output sizes match. Large emits 499,999
bytes in 19,999 writes, largest 49 bytes, retaining no sink output.

| Pair | Shape | Control p50 (µs) | Candidate p50 (µs) | p50 change | p95 change | p99 change | Throughput change |
|---|---|---:|---:|---:|---:|---:|---:|
| R1 | tiny | 6.230 | 6.170 | -0.96% | -1.10% | +2.77% | +0.54% |
| R1 | medium | 39.270 | 38.270 | -2.55% | -6.75% | -3.75% | +2.76% |
| R1 | large | 1867.037 | 1830.166 | -1.97% | +2.24% | +9.14% | +1.43% |
| R2 | tiny | 6.230 | 6.170 | -0.96% | -1.88% | -1.22% | +0.97% |
| R2 | medium | 39.090 | 38.100 | -2.53% | -3.14% | -3.47% | +2.77% |
| R2 | large | 1884.757 | 1837.941 | -2.48% | -2.59% | +1.38% | +2.58% |

Throughput change is the reciprocal of mean elapsed time for identical work.
Initial whole-child peak RSS is 30,212→30,316 KiB (+0.34%) in R1 and
30,072→30,304 KiB (+0.77%) in R2. These children include all three shapes,
opening and verification; this is not operation-local peak RSS.

**The initial large R1 p99 increase of 9.14% is retained as an adverse flag.**
R2 p99 is +1.38%. That triggered a separately recorded, post-hoc large-only
A1/B1/B2/A2 follow-up with 2,000 samples per child, ten warmups and unchanged
binaries/CPU/sink: **8,000 additional native samples**. Follow-up p50 changes
are -2.05%/-0.50%, mean -1.97%/-0.40%, p95 -1.88%/-0.33%, and p99
+2.18%/+2.78%. RSS changes are -0.41%/0.00%. No follow-up pair exceeds the
5% latency/throughput/RSS threshold. Every initial and follow-up repeat stays
within the recorded drift ceilings (5% median/mean, 10% p95, 15% p99), though
candidate follow-up median drift reaches 4.37%, larger than the observed gain.

The acceptance is a bounded reduction in confirmed allocation churn with
modest observed median changes and no persistent >5% tail flag in the longer
follow-up. It is **not** a tail-latency improvement or a claim that the initial
flag was proven to be noise. The post-hoc follow-up supplements rather than
replaces the original evidence; shared-host variability limits timing precision.
Peak memory does not improve in the retained profile.

## Validation and disposition

All **1,008 ODT tests and doctests** pass, including four private resource/
ownership regressions and two integration regressions. The unchanged full
harness passes **483 tests**, with one existing opt-in real-producer test
ignored. Total: **1,491 passing Rust tests/doctests**. Workspace formatting,
ODT all-target Clippy, strict ODT rustdoc, harness Clippy, crate boundaries,
and unchanged strict claim registry checks pass. The evidence verifier checks
all 14,000 native samples plus 50 instrumented samples, source/build/gate
bindings, identical corpus/output/sink data, drift and adverse flags, and
rejects a deliberately short sample vector. Instrumented samples are excluded
from native comparisons. Final artifact checks and scratch cleanup are recorded
in `gates.json`, `verification.json`, and `cleanup.json` in the evidence folder.
The owned scratch cleanup removes 9,230 files and 2,228,387,840 allocated
bytes (2.23 GB); the verifier reproduces its summary after removal.

The source and ADR reviews found no blocker and no exception needed. The
4 KiB cap is an explicit retention bound, not a tuned optimal size. Current
residual profiling keeps text-length checking as a possible later ODT candidate;
ODS joined-output preflight and local provider scheduling still need their own
admission evidence. This batch does not close those priorities, historical
ODG comparisons, native-producer coverage, or the full non-iWork goal.

The retained [0509 evidence and replay instructions](../results/change-0509/README.md),
[admission record](../results/change-0509/admission.json),
[profile summary](../results/change-0509/profile-summary.json),
[profile scope](../results/change-0509/profile-scope.md),
[source review](../results/change-0509/source-review.md), and
[ADR review](../results/change-0509/adr-review.md) provide the machine-readable
bindings and limitations for this record.
