# 0722 writer-local DOCX fusion packet

The candidate is retained. All 96 primary and 16 independent read-control hard
gates pass. Edit p50 improves 16.16–19.32% on generated-medium and 6.23–6.80%
on NumberedList. Allocation-region peak and net live bytes are unchanged.
Three read tail/max regressions and ten p99/max repeat flags remain visible,
including a pinned-media pair-4 p99 regression of 18.58%. Independent review
supports retention under the frozen gates with this tail limitation.

[Change report](../../0722-docx-writer-local-fusion-pilot.md),
[all paired measurements](measurements.md),
[design and independent review](design-review.md),
[trace matrix](trace-details.md), and
[aligned memory diagnostics](memory-diagnostics.json) describe the scope.
The broader non-iWork performance goal remains active.

The pilot uses CPU 12, generated-medium and `NumberedList.docx`, edit and
lifecycle phases, two ABBA cycles, 100 warmups, and 200 native samples. The
allocator lane follows native capture and covers the first four stages with
three samples and no warmups. The 64 outer invocations comprise 32 primary
native, 16 read-control and 16 allocator invocations, retaining 9,600 native
and 48 allocator samples. Eight pinned filesystem controls each use 100
warmup, 200 priming and 200 measured inner processes: 4,000 inner processes.
No pair, sample or stage was excluded or rerun.

Primary gates require at least 3% edit improvement and at most 3% lifecycle
regression in every paired p50 and mean. Allocation requests, requested bytes
and per-operation peak above start must regress at most 3%; net live bytes
must not increase. Zero allocation baselines require exact equality.
Read-control p50 and mean regressions are limited to 3%. Tails and repeat
drift above 5% are review flags, not hidden hard-gate substitutions.

The original public codec stays at its original module path and only gains
private child declarations. The shared `namespace.rs` remains byte-identical.
Existing read consumers are unchanged; only the obsolete writer helper and
its unused imports are removed from `document_part.rs`. New dual-state code
is confined to `alt/codec/document_scan.rs`, with independent frozen
scanner tests wired through the writer. Parser scratch is explicitly dropped
before the two separate, ordered MCE calls. The existing final body reader
and package publication contracts remain in place.

`source-guard.py --check` verifies exact transformations against archived
baseline/candidate sources. Native binaries were built and frozen before
trace instrumentation. `trace.py` only patches guarded sources temporarily;
`trace-run.py` restores the candidate census after each lane. All 21 traced
boundaries preserve ordered MCE and structured scanner metadata, with repeated
stderr and public report parity. Baseline range-reader ends remain null where
borrowed events prevent observation; reader and observer counts are separate.
The trace reduces successful structural reads from 2,820 to 1,410 and from
356 to 178 on the two package corpora. This is not an I/O, instruction or
physical-memory claim.

Candidate qualification covers 1,471 DOCX unit/integration tests, 75 passing
doctests with 31 existing ignored cases, formatting, warning-denied Clippy and
rustdoc. The harness passes 535 tests with one opt-in test ignored. Six
repository evidence gates and 11 corruption checks validate the final packet.
[Quality receipts and totals](quality-summary.json) retain exact commands.

Build and source maps, host metadata, fixture identity, child stdout/stderr,
raw reports, receipts, the 56-input pre-capture freeze and both initial/final
analyses are retained. The final source is the candidate. Cleanup removes
only the owned target, frozen binary, filesystem and trace scratch roots;
eight exact binary identities remain in `cleanup.json`. Replays accept a
live matching binary or that exact cleanup witness, with stable output.

From the repository root, reproduce the retained analyses and custody checks:

```text
python3 -B docs/performance/results/change-0722/audit.py
python3 -B docs/performance/results/change-0722/source-guard.py --check
python3 -B docs/performance/results/change-0722/memory-diagnostics.py --check
python3 -B docs/performance/results/change-0722/trace-details.py --check
python3 -B docs/performance/results/change-0722/quality-summary.py --check
python3 -B docs/performance/results/change-0722/artifact-seal.py --check
```

The audit replays `analyze.py`, `read-controls-analyze.py` and the trace
comparison into guarded temporary outputs. `negative-checks.py` retains
11 verified corruption refusals; it intentionally refuses to replace its
existing result. `artifact-manifest.json` covers every packet file except
itself and rejects Python caches and symlinks. To repeat builds or captures,
use a fresh packet/output location and the archived sources and exact
commands; the original capture is intentionally immutable.

No RSS, leak, cold-cache, hardware-counter, throughput, scaling or universal
read-tail improvement is claimed. Allocation lifetimes do not prove identical
host allocator-exhaustion scheduling. The 0721 rejection remains unchanged.

The terminal audit initially expected obsolete dictionary metadata for derived
allocation formulas. `terminal-audit-correction/` preserves that validator,
the initial corruption-check result, the exact audit diff and its reason.
Only the audit assertion changed; captures, frozen analyzers and their outputs
did not. All 11 corruption checks and the full audit passed again after
cleanup with the corrected assertion.
