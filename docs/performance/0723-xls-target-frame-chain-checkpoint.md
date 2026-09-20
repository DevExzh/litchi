# 0723 — rejected XLS target-frame chain checkpoint

The candidate is rejected. It removes the repeated worksheet-chain prefix walk
and improves owned `54016.xls` late-target q8 p50 by 79.31–79.86%, but fails
14 of 24 native groups and four of 16 repeated-query groups. The baseline
implementation is restored; the candidate and all observations remain archived.
The non-iWork performance program remains active. `performance_claim: none`.

| Owned repeated-query case | B1/A1 p50 delta | B2/A2 p50 delta | Mean delta range |
| --- | ---: | ---: | ---: |
| 54016 late target | −81.64% | −81.82% | −81.59% to −81.56% |
| 54016 missing target | +7.58% | +6.63% | +4.89% to +6.58% |
| Plan1 first target | +7.75% | +6.84% | +7.11% to +8.53% |
| Simple first target | +9.70% | +7.72% | +7.71% to +8.73% |
| 45365 first target | +6.71% | +7.89% | +7.30% to +7.81% |

Negative latency deltas are faster. Repeated-query statistics describe nine
fresh processes per leg, each timing a 50,000-query loop after two preparatory
queries. They are not individual-query latency percentiles. The two comparisons
are B1/A1 and B2/A2 within one ABBA cycle, preceded by A/A controls.

The regressions are not confined to tiny warm queries. Simple owned q2 p50
increases 9.72–11.11%, and Plan1 owned q2 increases 5.73–5.91%. The second
file-backed 54016-late pair regresses 6.13% in open-plus-eight p50 and 6.19%
in its mean, despite the faster indexed query. Both positive-benefit gates
pass, but these results do not satisfy the frozen non-regression conditions.
No samples were removed, recaptured, or pooled to change the decision.

The fixed plan requires at most 5% regression in both p50 and mean for every
paired metric. Only native q3, q8 and q3-to-q8 mean have a predeclared 10 ns
absolute exception; q1, q2, open, complete workflow and repeated loops do not.
The complete matrix has 66 failed central-statistic checks, 194 paired
p95/p99/maximum regression flags over 5%, 15 A/A central drift flags, and 31
within-phase central drift flags. Drift flags use absolute changes above 5%.
These flags are descriptive and do not identify a scheduling or hardware cause.
[Every comparison and failure](results/change-0723/measurements.md) is retained.
The repeated early-target regressions remain enough to reject this candidate
without interpreting noisy observations as an excuse to relax the gates.

The candidate adds one optional `(frame_offset, StreamChainCheckpoint)` to each
admitted worksheet index, retaining the original worksheet-start and SST
checkpoints. After a successful complete scan, it finds the first matching raw
frame and constructs a metadata-only cursor at that exact offset. This avoids
capturing a scanner cursor that has read ahead. Replay chooses the target
checkpoint only for matching frames at or after that offset; earlier queries
keep the original worksheet checkpoint. Later targets can still walk a suffix.

Diagnostic tracing confirms that original-target replay removes 1,044 FAT links
for 54016-late, 72 for Plan1-late and 258 for 45365-late. First targets avoid
only 1–14 links. The extra construction walk occurs once per fresh owner/index publication.
The five-query late-build → late-publish → earlier-target → original-target
→ missing-target probe preserves exact semantic reports, source read offsets
and lengths, and source-version call counts. Later-target behavior is covered
by the source test, not this route capture. Earlier-target replay retains its 14-link prefix;
missing-target replay has no indexed slot and makes no cursor call. All 12
primary counted-source cases retain exact read counts, bytes, length calls and
version calls. This is metadata work elimination, not reduced physical I/O.

The fixed logical index charge increases from 224 to 264 bytes. Successful
stored-target q2 usually adds 56 allocation bytes, one allocation/deallocation
pair, and 40 retained bytes; the 16-byte temporary path vector is released.
All q1/q3/q8 allocator counters and zero-budget/refusal counters are exact.
Missing-target and generated-capacity cases have different q2 deltas because
the increased fixed charge can reduce slot capacity; all remain within the
frozen 64-byte allowance. All 96 allocator groups pass. The 34 budget-fence
comparisons preserve outcomes, while Simple budgets 489 and 490 expose capacity
route changes. Allocation-region counters are not RSS measurements.

Source/execution fences, complete scan validation, selected-frame decoding,
duplicate order, fallible reservations and publication cancellation remain.
The bounded metadata walk adds work between existing cancellation checks; it
does not remove a check. There is no decoded-value/error cache, public API,
dependency, unsafe code, ambient library I/O or parallel execution change.
[Independent design review](results/change-0723/design-review.md) found no
source correctness blocker. The regression mechanism is not isolated by this
combined layout, selection and construction change. Any future experiment
must address early/missing-target costs and q2 construction before fresh timing.
The large late-target benefit alone does not justify retaining this version.

The baseline is `45cb480eaaac75f6279e7596995f8e9d257872c2`. Both phases use
unchanged release probes from 0684/0686, Rust 1.95.0, CPU 12 on AMD EPYC 9R45,
Linux 7.0.0-1012-aws and glibc 2.43. Owned bytes and files use warm OS caches.
Source, probe, fixture, tool and executable identities are bound at freeze and
capture completion. Native timings cover 14,400 fresh measured owners and
115,200 queries, plus 432 warmup owners. Repeated loops cover 864 processes
and 43.2 million measured queries. The allocator lane has 576 processes.
Diagnostic instrumented binaries and separate two-million-query profiles are
not used as paired performance evidence. No cold-device, remote, concurrency,
scaling, hardware-counter, cross-platform or native Office claim follows.

Candidate qualification passed formatting, all-target/all-feature CFB/XLS
checks, warning-denied Clippy and rustdoc, 1,904 CFB/XLS tests/doctests with two
existing ignored tests, and 61 facade tests. All 126 real XLS fixtures in both
source modes and a complete generated 70,001-cell visit preserve exact outcomes
and source metrics. All six repository evidence gates pass. Twelve verifier
controls detect altered outcomes, timing failures and allocation-boundary
violations. Earlier diagnostic compile/argument failures are archived separately
and excluded from native measurements. No registered claim or CRUD coverage
promotion is made. [Evidence and replay index](results/change-0723/README.md).

Post-cleanup replay reproduces the same rejected gate results. The initial trace
replay required a live diagnostic binary; its validator was extended to accept
an absent binary only with an exact cleanup identity witness. The original
validator, failure and correction are archived under
`post-cleanup-trace-correction/`. The frozen native driver, analyzer, plan and
raw measurements are unchanged. The independent terminal audit also corrected
its build-kind binding, repeat-benefit indexing and allocator group-count lookup;
these audit-only corrections preserve the recorded hard-gate rejection.
