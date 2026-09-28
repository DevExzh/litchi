# 0815 — borrowed PPTX event-arm bindings

The candidate is rejected. Although borrowing removes the two targeted payload
copies, no capture or lifecycle row meets the frozen benefit threshold.
Large capture is 4.270% slower and large lifecycle is 5.299% slower; the latter
triggers the latency veto. Production is restored exactly to the baseline.

[0814](0814-pptx-current-native-attribution.md) localized native samples to
two post-dispatch 32-byte copies in the private notes scanner. This candidate
changes the `Start` and `Empty` bindings to `ref` and removes four redundant
argument borrows. The event bodies, resolver timing, parser/error ordering,
limits, and complete buffered differential test module remain unchanged.
No public API, dependency, unsafe code, or ownership contract changes.

The baseline is `55bb2ead34`. All 9,196 production source hashes equal sealed
0813 after, allowing exact-source reuse of its six production quality gates.
All 35 previously read normative inputs remain unchanged. The candidate passes
six fresh production gates: formatting, all-feature/all-target checking, 1,241
tests with zero failures and three ignored across 85 suites, warning-denied
Clippy and documentation, and the full crate-boundary checker. Both probe legs
pass formatting, 36 tests, and warning-denied Clippy. iWork is excluded.

Before application, eighteen baseline qualification reports match the sealed
source/output/full-semantic oracles. Exact ordinary and profile assembly shows
two arm-local copy sequences before and zero after; scanner size falls from
2,320 to 2,107 bytes and instruction count from 434 to 393. Both pre-dispatch
windows already have zero vector moves. The static mechanism therefore passes,
but it is not evidence of a workflow benefit.

The fresh trial contains six paired native blocks in order AB BA AB BA BA AB,
two allocation blocks, and four scoped Callgrind publications: 310 reports and
6,718 measured outputs including qualification. No historical timings are pooled.
Timing processes have three warmups and thirty samples; allocation processes
have no warmup and three samples. The CPU is AMD EPYC 9R45, x86_64 Linux, pinned
to logical CPU 12. Rust/Cargo 1.95.0 use optimization 3, thin LTO, one codegen
unit, and unwind. Native builds have no instrumentation or frame-pointer override.

The frozen policy requires at least 3% capture/lifecycle p50 improvement with
bootstrap upper endpoint below one, rejects any p50 ratio above 1.05 with
lower endpoint above one, and requires non-increasing paired allocation calls,
bytes, net live bytes, and peak above entry. Quantiles use nearest rank; ratio
intervals use 10,000 resamples of six paired blocks, seed 815815, endpoints
250/9749. The outcome is zero benefits, one latency veto, and equality for all
144 paired allocation comparisons.

[Protocol](results/change-0815/protocol-review.md),
[candidate design](results/change-0815/candidate/design.md),
[source review](results/change-0815/source-review.md), and
[assembly gate](results/change-0815/codegen-gate.json) delimit the hypothesis.

The table gives medians of six process p50s in milliseconds. The ratio is the
median of six paired ratios and need not equal the quotient of displayed medians.

| Shape | Workflow | Before ms | After ms | After/before | 95% interval |
| --- | --- | ---: | ---: | ---: | --- |
| tiny | capture | 0.228871 | 0.228036 | 0.996898 | 0.992916–0.998888 |
| tiny | commit | 0.206986 | 0.206816 | 1.000625 | 0.993813–1.004553 |
| tiny | lifecycle | 1.403908 | 1.413533 | 1.005988 | 1.003832–1.009265 |
| medium | capture | 0.425122 | 0.433557 | 1.018742 | 1.010940–1.024774 |
| medium | commit | 0.291241 | 0.289701 | 0.994115 | 0.991716–0.998095 |
| medium | lifecycle | 1.980175 | 1.998045 | 1.010177 | 1.007406–1.014868 |
| large | capture | 16.031205 | 16.724164 | 1.042703 | 1.037801–1.056069 |
| large | commit | 1.264437 | 1.263137 | 1.000376 | 0.996840–1.003156 |
| large | lifecycle | 25.804666 | 27.168548 | 1.052994 | 1.049081–1.057175 |
| vendor | capture | 0.509672 | 0.518318 | 1.016889 | 1.009502–1.022071 |
| vendor | commit | 0.319211 | 0.318221 | 0.997337 | 0.991784–1.000614 |
| vendor | lifecycle | 2.129891 | 2.158641 | 1.013713 | 1.010134–1.016160 |
| unicode-vendor | capture | 0.513792 | 0.521048 | 1.015524 | 1.011321–1.018577 |
| unicode-vendor | commit | 0.319661 | 0.319117 | 0.997842 | 0.994625–1.001512 |
| unicode-vendor | lifecycle | 2.140772 | 2.169841 | 1.012084 | 1.010577–1.016338 |
| valid-4attr | capture | 0.490362 | 0.494108 | 1.007169 | 1.002278–1.016081 |
| valid-4attr | commit | 0.311521 | 0.311247 | 0.999100 | 0.994633–1.001669 |
| valid-4attr | lifecycle | 2.093296 | 2.116496 | 1.012858 | 1.006684–1.014424 |

All p95/p99 paired changes above 5% are retained below. These are individual
block diagnostics, separate from the p50 adoption veto. Blocks are zero-based.

| Shape/workflow | Metric | Block | Change |
| --- | --- | ---: | ---: |
| tiny/capture | p99 | 4 | +12.337% |
| tiny/commit | p99 | 5 | +32.672% |
| medium/lifecycle | p99 | 0 | +6.202% |
| medium/lifecycle | p99 | 2 | +5.312% |
| large/capture | p95 | 1 | +6.248% |
| large/capture | p95 | 2 | +6.537% |
| large/capture | p95 | 4 | +5.453% |
| large/capture | p99 | 1 | +8.233% |
| large/capture | p99 | 2 | +7.776% |
| large/capture | p99 | 4 | +6.193% |
| large/lifecycle | p95 | 0 | +5.730% |
| large/lifecycle | p95 | 1 | +5.277% |
| large/lifecycle | p95 | 2 | +7.192% |
| large/lifecycle | p95 | 3 | +5.196% |
| large/lifecycle | p95 | 4 | +5.826% |
| large/lifecycle | p95 | 5 | +6.711% |
| large/lifecycle | p99 | 0 | +5.696% |
| large/lifecycle | p99 | 1 | +5.453% |
| large/lifecycle | p99 | 2 | +5.690% |
| large/lifecycle | p99 | 3 | +5.114% |
| large/lifecycle | p99 | 4 | +6.066% |
| large/lifecycle | p99 | 5 | +6.733% |
| valid-4attr/capture | p99 | 5 | +7.245% |
| valid-4attr/commit | p99 | 4 | +6.918% |
| valid-4attr/lifecycle | p99 | 2 | +7.764% |
| valid-4attr/lifecycle | p99 | 3 | +14.463% |

Native spread flags total 23: 1 p95, 11 p99, 11 rss_kib.
Every flagged group is retained in the numerical analysis. No p50 spread
exceeds 5%. Allocation measurements have no regression or spread flags.

Process RSS medians are diagnostic and do not imply allocation savings.

| Shape/workflow | Before RSS KiB | After RSS KiB |
| --- | ---: | ---: |
| tiny/capture | 5170 | 5130 |
| tiny/commit | 5080 | 5146 |
| tiny/lifecycle | 5050 | 5200 |
| medium/capture | 5260 | 5202 |
| medium/commit | 5480 | 5516 |
| medium/lifecycle | 5356 | 5324 |
| large/capture | 18662 | 18732 |
| large/commit | 18668 | 18668 |
| large/lifecycle | 18668 | 18636 |
| vendor/capture | 5516 | 5454 |
| vendor/commit | 5454 | 5436 |
| vendor/lifecycle | 5388 | 5428 |
| unicode-vendor/capture | 5548 | 5452 |
| unicode-vendor/commit | 5418 | 5516 |
| unicode-vendor/lifecycle | 5406 | 5420 |
| valid-4attr/capture | 5482 | 5484 |
| valid-4attr/commit | 5484 | 5420 |
| valid-4attr/lifecycle | 5516 | 5482 |

3 paired RSS increases exceed 5%:

- tiny/commit, block 3: 5048 → 5308 KiB (+5.151%).
- tiny/commit, block 4: 4940 → 5240 KiB (+6.073%).
- medium/lifecycle, block 3: 5132 → 5452 KiB (+6.235%).

All four Callgrind publications pass exact-owner, termination, source, output,
and self-cost conservation checks. Scanner self Ir is 16,724,054 before and
16,713,423 after in both repeats, while scanner inclusive Ir rises from
313,977,875 to 317,662,247 in the first pair and 314,089,381 to 317,664,178
in the second pair. Reader calls remain 282,612 and namespace-push self
Ir remains 13,019,166. The candidate introduces 282,612 scanner calls to the
`Result<Event, Error>` drop function. Copy removal does not eliminate all work:
the changed drop path is visible in this guest profile. This is an attribution
observation, not a causal native-cycle or phase-fraction claim.

Native timing, allocation metrics, static assembly, and guest instructions
are separate evidence. This deterministic PPTX matrix does not establish
cold-cache, remote-source, non-seek output, concurrency, cross-format, or
universal-document behavior. No performance baseline advances after rejection.

The existing preservation, typed-refusal, bounded-resource, immutable snapshot,
atomic edit, reversible patch, and physical ownership contracts remain intact.
The three unrelated workspace files are excluded from this batch.

Replay with `python3 -B docs/performance/results/change-0815/validate.py --final`.
The [packet README](results/change-0815/README.md) records the original build
and capture sequence. The [result review](results/change-0815/results-review.md)
and [profile review](results/change-0815/profile-review.md) independently check
the outcome. Numerical results are immutable; the separate
[decision](results/change-0815/decision.json) and
[disposition](results/change-0815/disposition.json) record rejection and restoration.

This closes the bounded event-arm borrowing hypothesis. The retained 0813
implementation remains the baseline. Further work should return to the broader
unmeasured CRUD, corpus, I/O, and scaling gaps rather than repeat this source
spelling without a new measured mechanism. The broader project goal remains active.

Aggregate validation passes before cleanup, including byte-for-byte replay of
all six numerical, quality, qualification, and code-generation checks. Cleanup
verifies the six executable hashes and restored source, then removes only the
owned target: 11,216 files / 4,153,937,889 logical bytes.
Final sealed replay verifies the post-cleanup record.
