# ROMAN/ARABIC evaluation harness review

This document records the final CPU6 measurement boundary for the bounded
ROMAN/ARABIC formula-function profile. The independent static corpus review is
in [independent-harness-review.md](independent-harness-review.md). The prior
CPU2 captures and their diagnostic report were moved to
[exploration/](exploration/); this document does not treat those CPU2 results as
clean final evidence.

The final runner is `roman-harness/run.py`, whose current SHA-256 is
`c547a088da9d0029d0cf85109d5a360e971cb3a347412b60eafa5c8978792056` and whose
command pins each child with `taskset -c 6`. The other retained harness hashes
are `src/main.rs`
`72d2ccc147fc04db683b64d426adafe2bde7243c9d47707b0e68651e3a4806c5`,
`Cargo.toml`
`638c9176b895d60fada4f715e990880b8a4c0bc77491ede3555c8fc1bc229d81`, and
`Cargo.lock`
`8f2aff51ed4df52cf1be91363494b82feca0cdcf75eb2e17bb3699f51c37efaa`.
The final candidate gate source hashes are `evaluation.rs`
`5b2320a6204b96185b6e91eaa27d3f166ac19555e71b307c79a0600329975cf4`,
`evaluation/roman.rs`
`3441233624f68e5efc72157953d3b7137bbb86f6c9835fd809eee40674342da1`, and
`ods_formula_roman_evaluation.rs`
`4bcc00fd7278787c1ae77a4bd58118677a743a92bf023f6f725309bd6b8877ff`.
The full source manifest is [../gates/results.json](../gates/results.json).

The corpus has 53 comparable cases and 41 candidate-only cases. Each case is
sent through `parse`, `evaluate`, and `parse-evaluate`, giving 159 baseline
comparable rows, 159 candidate comparable rows, and 123 candidate-only rows
when a complete capture is present. Fixed cases use 128 repeats; the
64/256/1024/4096-unit scale cases use 64/32/8/2 repeats. The candidate-only
cases cover ROMAN values 3888, 499, and 998 in formats 0 through 4, ARABIC
case and nesting, zero/truncation/logical format behavior, domain and typed
refusals, 64-to-4096-symbol ARABIC scans, and 64-to-4096-call ROMAN
concatenations. Expected values and typed failures are fixture data, including
the checked ODF format-2/3 vectors, and are not derived from the implementation
during measurement.

| Phase | Timed operation | Outside the timer |
| --- | --- | --- |
| `parse` | Parse and drop a new expression on every repeat | Context setup and preflight |
| `evaluate` | Evaluate one parsed immutable tree on every repeat | Parse/tree/context setup and preflight |
| `parse-evaluate` | Parse and evaluate on every repeat | Context setup and preflight |

The binary's `Instant` samples the operation in the table. One preflight checks
the expected scalar or typed refusal before warmups. Results are dropped after
each iteration while the context remains alive for `live_after`. The
`/usr/bin/time -v` `max_rss_kib` sidecar covers the whole process, including
startup and setup, and is not an operation timer.

The counting allocator resets after setup. `alloc_calls`, `requested_bytes`, and
`released_bytes` are totals in the timed region; reallocation contributes its
new request and old release. `peak_live_delta` is the tracked live-byte
high-water delta. `output_reserved_bytes` records the largest evaluator-result
reservation observed in a sample. These are allocator/accounting instruments
and tracked heap bytes, not OS allocator or peak-RSS measurements.

The exact baseline is
`d0c1ca700ceda177c27a820acffe0de534720390`; it has no ROMAN/ARABIC
implementation, so candidate-only rows have no before/after speedup comparison.
For the earlier 159-row CPU2 comparable capture, status/failure, counts,
checksums, allocator call counts, requested/released bytes, live-before and
after, tracked peak-live bytes, and output reservations matched exactly in the
retained p50/max records. That heap parity does not make CPU2 timing or RSS
clean, and final CPU6 receipts must be checked against their own raw rows and
source custody.

The prior CPU2 ABAB replay recorded six repeated comparable p50 regressions over
the 5% review trigger: `evaluate/control-utf8-left-1024` (+8.40%),
`evaluate/control-utf8-left-4096` (+11.25%),
`evaluate/bitwise-coerce-text` (+10.05%),
`parse-evaluate/control-coerce-4096` (+5.16%),
`parse-evaluate/control-utf8-left-1024` (+6.69%), and
`parse-evaluate/control-utf8-left-4096` (+8.78%). The original unchanged binaries also reproduced the principal UTF-8 and
coercion regressions in a separate CPU6 probe. During later experiments CPU2
was occupied by an unrelated benchmark at
`/home/zhuhe/litchi-goal-0557-target/retained/baseline/normal`, observed at
about 96% CPU; contaminated experiment deltas exceeded 300%. Final measurements
therefore use CPU6 and the revised windowed scanner. These two changes must
not be conflated when interpreting historical results.

Root verification of all final CPU6 rows confirms exact comparable result and
tracked-heap parity. Four A/B pairs cover all 34 initial flags (272 rows).
Two combined-phase p50 regressions remain (+7.60% bitwise text coercion and
+5.12% radix fractional-input error), together with eight tail-flag lanes; no
repeated RSS median exceeds 5%. The full qualification and follow-up are in
[report.md](report.md); the feature extension is not a no-regression claim.

Five primary and four investigative CPU6 `perf stat` captures pass status,
source/binary custody and accounting checks. Counters use `cycles`,
`instructions`, `branches`, and `branch-misses`; they include process setup,
preflight, warmups, timed work and output. They are whole-process evidence and
do not establish CPU causality. Final scratch roots and benchmark ELFs were
removed only after these checks and source review closed.

The final host snapshot is [environment.json](environment.json). Historical
CPU2 raw rows, ABAB replay, counters, environment, and report remain under
[exploration/](exploration/) for traceability.
