# ROMAN/ARABIC evaluation harness review

This is the retained measurement-boundary record for the bounded ROMAN/ARABIC
formula-function profile. The independent static review is in
[independent-harness-review.md](independent-harness-review.md); no source,
harness, build, or benchmark process was changed by this review.

The corpus has 53 comparable cases and 41 candidate-only cases. Each case runs
through three phases, yielding 159 baseline comparable rows, 159 candidate
comparable rows, and 123 candidate-only rows (441 rows total). Every capture
used three warmups and 15 measured iterations. Fixed cases use 128 repeats;
64-, 256-, 1024-, and 4096-unit scale cases use 64, 32, 8, and 2 repeats.
The candidate-only corpus covers ROMAN values 3888, 499, and 998 in formats 0
through 4, ARABIC spelling and nesting, zero/truncation/logical format
semantics, domain and typed refusal paths, 64-to-4096-symbol ARABIC scans, and
64-to-4096-call ROMAN concatenations. Expected values and typed failures are
fixture data, including the checked ODF format-2/3 vectors; they are not
inferred from the implementation during measurement.

| Phase | Timed operation | Outside the timer |
| --- | --- | --- |
| `parse` | Parse and drop a new expression on each repeat | Context setup and preflight |
| `evaluate` | Evaluate one parsed immutable tree on each repeat | Parse/tree/context setup and preflight |
| `parse-evaluate` | Parse and evaluate on each repeat | Context setup and preflight |

The binary's `Instant` samples the operation above. A one-time preflight checks
the expected scalar or typed refusal before warmups. The evaluator result is
dropped after every iteration, while the context remains alive for the
`live_after` observation. The `/usr/bin/time -v` `max_rss_kib` sidecar covers
the whole process, including startup and setup, and therefore is not a timer
for the operation.

The counting allocator resets after setup. `alloc_calls`,
`requested_bytes`, and `released_bytes` are totals in the timed region;
reallocation contributes its new request and old release. `peak_live_delta`
is the tracked live-byte high-water delta. `output_reserved_bytes` is the
largest evaluator-result reservation observed in a sample. These values are
allocator/accounting instrumentation and tracked heap bytes, not OS allocator
or peak-RSS measurements.

All 159 comparable rows on each side and all 123 candidate-only rows completed
with status 0. The comparable baseline is the exact
[d0c1ca700ceda177c27a820acffe0de534720390](../README.md) source without the
Roman module; candidate-only Roman/ARABIC rows consequently have no valid
before/after speedup comparison. For the 159 comparable rows, status/failure,
success/refusal counts, checksums, allocator call counts, requested/released
bytes, live-before/live-after, tracked peak-live bytes, and output reservation
fields matched exactly in the retained p50/max records. RSS remains a
process-level noise-sensitive measurement and is excluded from that parity
statement.

The four-round ABAB replay contains 328 rows for the 41 initially selected
comparable flags. Every replay row has status 0. `p95` and `p99` are quantiles
of only 15 samples and often equal the sample maximum; they are recorded tail
flags rather than production percentile estimates. The final disposition is
still open because six repeated comparable p50 regressions exceed the 5%
review trigger; see [report.md](report.md).

The five retained hardware-counter receipts use `cycles`, `instructions`,
`branches`, and `branch-misses`. They are whole-process `perf stat` counts over
the same harness binary invocation, including startup/setup/warmups/output;
they are not operation-only counters and do not establish CPU causality. The
receipts and statuses are in [perf-stat/](perf-stat/), and all five status
sidecars contain `0`.

Custody and replay evidence:

- [baseline source hashes](baseline/source-sha256.json) and [candidate source hashes](candidate/source-sha256.json)
- [baseline binary provenance](baseline/binary-provenance.json) and [candidate binary provenance](candidate/binary-provenance.json)
- [baseline comparable raw rows](baseline/comparable/raw.csv), [candidate comparable raw rows](candidate/comparable/raw.csv), and [candidate Roman raw rows](candidate/roman/raw.csv)
- [ABAB raw rows](abab/raw.csv), [ABAB summary](abab/summary.json), and [initial flag selection](initial-flags.json)
- [gate receipt](../gates/results.json)

The profiles were pinned to CPU 2, but the host was shared with unrelated Cargo
activity targeting `/home/zhuhe/litchi-goal-0557-target`. A separate benchmark at
`/home/zhuhe/litchi-goal-0557-target/retained/baseline/normal` was also observed
pinned to CPU 2 at about 96% CPU during the window; the resulting host noise was
reported as greater than 300%. Treat the CPU 2 receipts as provisional diagnostic
evidence rather than a clean isolated baseline. Root is probing the preserved
binaries on CPU 6. The environment snapshot is [environment.json](environment.json);
the shared host and lack of hard isolation limit timing generalization.
