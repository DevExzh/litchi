# Reference-metadata performance capture analysis

The single authorized revised frozen-source capture used session `38187`, three warmups, and fifteen fresh child processes in both `evaluate` and `parse-evaluate`. It retained 840 baseline rows (28 controls) and 4,080 candidate rows (28 controls plus 108 metadata cases), for 4,920 timed rows. The candidate freeze is [`gates/freeze.json`](../../gates/freeze.json) with SHA-256 `b4a0f0c9665e10f67fafa1b179bfbfdd8be2c65e6ec897bac86db1b2afc9330a`; the contract hash is `87f20c419139bf721221d06c8714edd900ae7a57cbedecf070878e4f6cf64d8a` and the isolated gate lock hash is `58b4be6c…40a3e3`.

## Capture integrity

Both preflights passed: 28 baseline cases and 136 candidate cases. The raw receipts contain all expected samples and phase groups, source/profile hashes remained stable before and after timing, and target cleanup was recorded. The candidate matrix checks exact values, shapes, typed errors and failures before timing. The retained `performance/verify.py` had one known stale bound: it classified every `reference-metadata-*` case as zero-read, while the six computed ROW/COLUMN IF lanes each read exactly two selected cells. That expected verifier failure is retained at [`diagnostic-stale-read-bound-verifier/`](diagnostic-stale-read-bound-verifier/). The independent [`root-performance-audit.json`](../../root-performance-audit.json) applied the reviewed exact read contract and passed all 4,920 retained rows; the raw receipts and preflight/case-matrix evidence record the same two-read contract.

The setup-only failed launch is retained at [`diagnostic-revised-freeze-path-failure/`](diagnostic-revised-freeze-path-failure/). Its relative freeze path was resolved from the batch directory, so it created no build, preflight, or timed rows. It is not part of the 4,920-row capture.

## Matched controls

Across 56 control/phase groups, evaluate and parse-evaluate p50 latency deltas range from `-4.571%` to `+6.284%`; the only latency flag above 5% is `reference-conditional-256x4-sumifs` evaluate at `+6.284%` (`83,060.5` to `88,280.0` ns/repeat). Its parse-evaluate delta is `-1.852%`. The canonical root audit used exact `elapsed_ns_p50/repeat` values with an independent unpaired median-ratio bootstrap (5,000 resamples, derived seed `5554865353495640452`, percentile bounds), giving the descriptive 95% interval `[ -0.937%, +12.014% ]`. A 100,000-resample replay with the same exact normalized floats and seed gives `[ -0.728%, +12.032% ]`; both intervals describe sample variability and are not causal regression claims.

RSS deltas range from `-5.436%` to `+4.722%`; one group is below -5% and none is above +5%. Work, resolver reads, allocator calls, requested/released bytes, peak-live bytes, result-live budget, and input/output byte counters are identical across all comparable control groups. SUMIFS evaluate has work `1,839`, resolver reads `1,792`, allocator calls `42`, and bytes/repeat `52` on both builds. The candidate adds metadata dispatch and reference-shape state; `conditional.rs` adds `SourceReference` policy arms, but this ordinary-reference SUMIFS fixture does not enter that branch. Shared evaluator code layout remains a possible mechanism; the receipts do not establish one.

## Metadata lanes

The 108 metadata cases form 216 phase groups: 202 normalized zero-read groups, two one-read groups for `SHEET(ABS(range))`, and twelve two-read groups for the computed ROW/COLUMN IF, IFERROR, and IFNA value-array refusals. Expected outcomes are 78 finite-number groups, 24 logical groups, 80 `#VALUE!` groups, 10 `#REF!` groups, two `#N/A` groups, 16 typed unsupported-reference groups, four typed reference-cell resource failures, and two cancellation groups. The cancellation lane uses four repeats and records zero raw and normalized reads. No metadata lane changes allocator/work/read accounting beyond its intended descriptor/shape workload; detailed p50 values remain in [`performance-report.md`](performance-report.md) and raw receipts.

## Prior diagnostic and host limits

The superseded pre-computed-array capture remains under [`diagnostics/computed-array-preflight/performance-results/`](../../diagnostics/computed-array-preflight/performance-results/): 840 baseline plus 3,810 candidate rows, with a +5.865% SUMIFS evaluate flag and descriptive interval `[+0.528%, +6.831%]`. The revised capture is the result for the expanded 108-case matrix; both observations remain disclosed. Together, the two complete captures retain 9,570 timed rows.

The revised run has environment manifests for a 32-CPU host but no retained `capture-context-before.json` with load or affinity. Root observed unrelated preparation workload, and the host was not isolated. Timing and RSS deltas therefore remain descriptive observations without causal attribution. The profile measures evaluator and typed resource paths, not save, recalculation, native producer acceptance, or cross-platform timing identity.
