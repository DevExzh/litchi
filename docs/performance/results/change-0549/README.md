# 0549: checked bitset test-and-mark rejected

The private test-and-mark candidate is rejected. Every primary XLS workflow
regresses in both repeats, despite lower instruction counts. The final runtime
is restored to the exact measured baseline. Candidate code, tests, and all
paired evidence remain in this bundle; no production speedup is retained.

| Primary XLS workflow | Repeat 1 p50 change | Repeat 2 p50 change |
| --- | ---: | ---: |
| `xls_source_backed_open` | +7.097% | +2.080% |
| `xls_source_backed_open_one_cell` | +8.254% | +5.058% |
| `xls_owned_source_open` | +10.876% | +7.053% |
| `xls_owned_source_open_one_cell` | +9.107% | +6.026% |

Positive values are regressions. The frozen requirement is at least 3%
improvement in every case and repeat. None of the eight primary comparisons
passes. CFB few-large p50 also regresses 3.870% and 3.706%.

The hypothesis does change emitted work: the candidate uses register `bts`
and a separate `bt` with an ordinary store. It does not reuse the `bts` carry
result for cycle detection and does not emit a memory test-and-set instruction.
Collector self Ir falls 2.939% for XLS-owned and 2.940% for CFB few-large in
both repeats; XLS-owned constructor inclusive Ir falls 1.415–1.461%. These
are positive timed-constructor counts, not operation-local hardware cycles.
The shorter executed instruction sequence does not prove lower latency, and
this campaign does not establish a particular opcode or cache mechanism as
the cause of the observed slowdown.

Allocation calls, allocated bytes and incremental region peak pass their
separate gates. All 12,800 public-guard samples match the exact oracle. Maximum
invalid p50/mean ratios are 1.312x same-invalid baseline and 1.071x baseline-valid,
below the frozen 4x and 2x limits. Every adverse and repeat-variation flag is
retained for individual review rather than hidden by the admission envelopes.

The implementation changes only a private checked bitset operation and the
scratch exact-chain call site. It preserves logical/backing bounds, diagnostic
strings, first-cycle ordering before duplicate append, reservation labels/order,
zero-fill, scratch reset/reuse and physical ownership validation. An already-set
word is assigned its unchanged value before reporting a cycle. The helper uses
safe Rust and adds no public API or dependency. Candidate-only tests cover
word boundaries, exact errors, missing backing words, unchanged storage, and
376,264 exhaustive table/start/count combinations per CFB feature configuration.

The candidate passed all 15 quality checks with 4,388 test executions.
After exact baseline restoration, all 15 final checks passed with 4,382 test
executions: CFB 310 in each feature configuration, XLS 1,345, DOC 1,187,
and PPT 1,230. Workspace compilation, Clippy, rustdoc, formatting, crate
boundaries, and the claim registry also passed.

The campaign retains 48,000 main native samples, 1,440 allocator samples,
16 profile children with 80 timed constructor dumps and 12 CFB setup dumps,
four whole-child hardware diagnostics, 64 measured public-guard children and
32 single-sample smoke children. Main instrumented elapsed and smoke samples
are excluded from native admission. All build/capture children run serially;
other shared-host activity is not controlled. No stable-tail, provider, I/O,
cache, concurrency or scaling improvement follows.

The common public guard is the unchanged tracked 0548 example. The baseline
collector disassembly exactly matches its sealed 0548 counterpart, as bound
by [continuity evidence](baseline-loop-continuity.json), but all performance
measurements are fresh. See [baseline loop review](baseline-loop-review.md),
[source review](source-review.md), [result review](results-review.md),
[decision](decision.json), [quality summary](quality-summary.json),
[protocol](protocol.md), and [ADR matrix](adr-compliance.md).

The final runtime-plus-test manifest equals the measured baseline manifest;
retained baseline binaries supply that exact source’s measured evidence. All
six binary hashes are verified before the sole target is removed. Strict
replay validates source, receipts, reviews, decision, documentation and cleanup.
Raw/intermediate artifacts are preserved, including a disassembly line-label
correction described in [analysis execution scope](analysis-execution-note.md).

Cargo-fuzz and nightly are unavailable in the recorded host observations; no
sanitizer campaign is claimed. This batch reports OLE owner test executions
separately from workspace checking. OLE2/OOXML remain the active priority; ODF
is deferred until that optimization goal completes and iWork is excluded.
After two unsuccessful CFB loop replacements, [the next target](next-target.md)
examines larger OOXML work while preserving the established CFB baseline.
