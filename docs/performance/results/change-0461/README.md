# ODP attribute matching experiment

The split between attribute matching and decoding is **rejected** under the
frozen practical gate. Normal p50 improves in every measured lane, but R1 medium
and large improve only 2.060% and 2.155%, below the required 3%. R2 medium/large
clear that threshold; all six normal confidence intervals lie below zero.
Every regional allocator metric remains exactly unchanged, so there is no
practical memory benefit to justify retention. No production improvement from
this experiment is claimed. The accepted 0460 staging optimization remains.

The production source is restored byte-exact to baseline revision `05f432d48`.
The measured candidate, full source text, assembly, tests and all raw reports
remain available for independent inspection.

## Hypothesis and generated code

The 0459 resolved profile puts `ElementAttrs::get` at 18.50% of commit samples
and 30.48% of snapshot-open samples, with its `lookup` helper at 6.12% / 9.99%.
Those inclusive diagnostic shares overlap and include warmups. The 0461 baseline
includes 0460 staging fusion, which does not change this initial/readback parser.

Baseline assembly shows a call to `lookup` for each cached attribute, followed
by result loads and discriminator checks. The tested change keeps the exact
namespace-first/local-name predicate and separates lazy value decoding. Candidate
assembly inlines matching into `get`, removes the `lookup` symbol, calls
`decode_value` only on a hit, and shrinks the `get` frame from 0x148 to 0x128 bytes.
This confirms the mechanism, but does not establish a sufficiently useful
end-to-end result by itself. See [source review](source-review.md), both
assembly receipts and their bound text artifacts.

The private state machine retains cache/source order, first match, lazy decoding,
malformed/duplicate attribute reachability and error wording, and borrowed
namespace resolution. Drawing-attribute harvesting is untouched. The new test
uses the independent raw-attribute reader and covers cached versus freshly
scanned invalid values, including the old advance-without-cache behavior after
a fresh decode error. Candidate owner tests pass 372 cases, with strict all-target
Clippy and formatting also passing. See [integration notes](integration-notes.md).

## Primary matrix and decision

The unchanged ordinary owned ODP append harness measures snapshot opening,
transaction, add, commit and sequential sink publication. Each variant retains
twelve reports: 64/4,096/8,192 slides, normal and allocator binaries, R1/R2,
three warmups and thirty retained samples. All 24 reports / 720 operations pass
the independent oracle. Rust 1.98.1, equal release flags, CPU 2 and serialized
workloads are frozen in [protocol.json](protocol.json); sequence is A1/B1/B2/A2.
Baseline revision is `05f432d48`. Both executable/source epochs are authenticated.

| Repeat | Instrumentation | Shape | Baseline p50 ms | Candidate p50 ms | Delta | 95% median-delta interval |
|---|---|---|---:|---:|---:|---:|
| R1 | normal | tiny | 1.837 | 1.799 | -2.058% | [-2.259%, -1.718%] |
| R1 | normal | medium | 70.214 | 68.767 | -2.060% | [-2.363%, -1.819%] |
| R1 | normal | large | 142.033 | 138.973 | -2.155% | [-2.591%, -2.060%] |
| R1 | allocator | tiny | 1.983 | 1.991 | +0.383% | [+0.241%, +0.612%] |
| R1 | allocator | medium | 76.983 | 76.284 | -0.908% | [-1.067%, -0.758%] |
| R1 | allocator | large | 153.683 | 153.328 | -0.231% | [-0.466%, -0.007%] |
| R2 | allocator | large | 154.807 | 153.371 | -0.928% | [-1.074%, -0.587%] |
| R2 | allocator | medium | 76.819 | 76.180 | -0.832% | [-1.100%, -0.692%] |
| R2 | allocator | tiny | 1.986 | 1.984 | -0.115% | [-0.321%, +0.471%] |
| R2 | normal | large | 142.790 | 138.323 | -3.128% | [-3.406%, -2.904%] |
| R2 | normal | medium | 70.864 | 68.118 | -3.875% | [-4.099%, -3.652%] |
| R2 | normal | tiny | 1.841 | 1.799 | -2.294% | [-2.646%, -1.794%] |

R1 allocator-tiny p50 rises 0.383%; all other allocator p50 changes are small
reductions. Every individual p95/p99, throughput and RSS value remains in
[summary.json](summary.json). No elapsed or RSS comparison exceeds +5%.
Process-lifetime maximum RSS varies from -2.537% to +1.852%, with no reduction
claim. Allocated bytes, allocation/deallocation/reallocation calls, peak above
entry and retained live delta are exactly equal across matched allocator lanes.
The smaller generated stack frame is not a measured regional heap/RSS benefit.

The keep gate requires four normal medium/large p50 improvements of at least 3%,
each with a bootstrap upper delta bound below zero. Two of the four fail the
practical threshold. No extra timing rerun is used to seek a more favorable
outcome, and the threshold is not relaxed after observing the results.
Quantiles use midpoint p50 and nearest-rank p95/p99. Intervals use 10,000 seeded
independent median-ratio resamples; they neither remove run-order effects nor
correct for multiple comparisons. See [decision.json](decision.json).

## Supplementary diagnostics

Four separate large normal phase reports retain 120 operations. Snapshot opening
p50 falls 3.942% / 4.291%, commit 4.168% / 3.056%, and transaction 3.096% / 0.406%
in R1/R2. These are separate diagnostic clocks, not the acceptance matrix or
causal proof for every small phase shift. The transaction-phase changes are
not attributed to this optimization. See [phase-summary.json](phase-summary.json).

Two separate 100-operation whole-process counter runs include fixture setup,
warmups, validation and reporting. Instructions fall 3.257%, cycles 3.231%,
branches 1.543% and branch misses 3.359%; cache misses rise 0.258%. Raw runtimes
and perf scaling percentages are retained. These are not operation-only counters
and do not override the practical latency gate.

## Reproduction and limits

The bundle retains immutable captures and command receipts, exact source delta,
source/build manifests, binary and assembly bindings, frozen protocol/drivers,
raw rows, derived summaries and rejection. `source-delta.json` binds complete
before/after `.txt` artifacts of the single changed Rust file to compiled epochs.
Independent final reopen, semantic readback, patch/no-op and preservation checks
remain required. No new dependency, API, worker, unsafe code or cache layout is
introduced by the experiment.

Run `python3 -B verify.py --portable` in a complete copied bundle to authenticate
all identities and recompute both summaries. Before cleanup, `--precleanup`
also checks retained executables and the restored baseline source. `negative.py`
requires a summary mutation to fail even after resealing. `finalize.py` records
precleanup and fresh-copy replay, then inventories and removes only
`/tmp/litchi-goal-0461`; `seal.py` refreshes the seal between proof-producing steps.
All final gates pass: 387 candidate harness tests (one ignored), restored-source
Clippy/rustdoc/formatting/boundaries, precleanup, portable replay and tamper
rejection. Cleanup removes four temporary executables totaling 233,048,120 bytes.
The full outcomes are in their receipts.

No cold-cache, range-source, multiworker scaling, other-CRUD or Office GUI claim
is made. No new coverage is promoted: 439 selectors / 36 defaults remain. The
full non-iWork goal stays open. Repeated cached matching remains a measured
opportunity; reducing its algorithmic work needs another separately validated
and measured design rather than retaining this below-gate factoring. See the
[bounded index proposal](next-work.md).
