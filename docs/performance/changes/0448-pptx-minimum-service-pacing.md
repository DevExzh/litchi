# 0448: Calibrate minimum-service PPTX range pacing

Retain an opt-in `--transfer-delay-policy minimum-service` in the standalone
provider-lifecycle harness. It credits elapsed wrapped-source work and fixed-wait
overshoot against the combined fixed-plus-transfer target, avoiding an unnecessary
second sleep when that target is already met. `separate-sleeps` remains the
default. Explicit policy selection requires an explicit transfer rate.

One build, CPU 2 and one worker compare both policies over plain/media-rich
managed cross-slide-copy lifecycles: 64 KiB ranges, 200 us fixed delay, 25 MiB/s
nominal transfer rate, 30 samples/3 warmups and two reversed repeats. Eight
reports retain 240 samples, with four additional media-rich profiles.

| Corpus | Separate API p50 ms, R1/R2 | Minimum-service API p50 ms, R1/R2 | Change R1/R2 |
| --- | ---: | ---: | ---: |
| Plain | 152.815 / 152.816 | 123.323 / 123.347 | -19.299% / -19.284% |
| Media-rich | 2,578.539 / 2,572.653 | 2,454.029 / 2,460.901 | -4.829% / -4.344% |

The frozen plain 5% gate passes in both repeats. Every source-work, nominal-counter
and output identity check matches, and every serial API clock meets its combined
service floor. All 30 absolute 5% paired triggers are lower latency: plain
open/plan/publication/API-sum p50/p95/p99 and media-rich open p50/p95/p99 in both
repeats. No positive paired or RSS trigger, and no repeat trigger, exceeds 5%.
See [complete measurements](../results/change-0448/measurements.md) and
[the exact review](../results/change-0448/decision.json).

This compares two explicit delay models; it does not establish a production
speedup, physical bandwidth or shared-link behavior. Nominal counters are targets,
not measured sleeping time. Whole-process RSS is about 19.4–19.7 MiB plain and
784.1–785.3 MiB media-rich. The sink retains full output; operation allocator
attribution is unavailable and boundary gauges do not establish allocation peaks.

All 381 harness tests pass, including three new tests and expanded real PPTX
policy/output equivalence. Feature/workspace checks, warning-denied rustdoc,
formatting and boundaries pass. Strict harness lint retains 29 inherited
diagnostics with none added. Twenty deliberate report corruptions reject.

Both profiles retain all callchains. SHA-256 accounts for 63.2–63.9% of weighted
self period in the lifecycle-frame subset; memory moves follow at 17.6–17.9%.
The subset includes untimed work and CPU samples omit blocked sleep. Source
freshness and exact publication remain mandatory for any later read-work reuse.

Replay the sealed exported evidence without the build or profiler:
`python3 -B docs/performance/results/change-0448/verify.py --sealed --cleanup`.
[Validation notes](../results/change-0448/validation-notes.md) retain custody,
corruption probes and cleanup. Native breadth, cold I/O, bounded existing append,
repackaging and scaling remain incomplete; the full non-iWork goal stays active.
