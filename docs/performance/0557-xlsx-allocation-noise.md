# 0557: XLSX allocation instrumentation and unstable baseline pilot

The XLSX sorted-merge candidate was not admitted to measurement. The fresh baseline-only pilot produced maximum absolute paired p50 drift **9.7872%**, exceeding its preregistered **5%** stability limit. During 0557, no candidate source was applied, no candidate binary was built, and no candidate result or production speedup is claimed. The tested harness instrumentation is retained.

OLE2 and OOXML remain first priority. ODF optimization is deferred until that goal completes; iWork is excluded.

## Harness change

The source-backed XLSX report now separates allocation evidence for edit staging and commit-core while retaining the original combined staging-plus-commit interval. One continuously owned observer region captures the split and final boundaries under its mutex. Counter sums, live endpoints, and the maximum segment peak reconstruct the combined interval exactly. Missing peaks, overflow, observer invalidity, and partial evidence vectors fail closed. Eager workflows continue to omit allocation phase vectors; omission is unavailable evidence, never a measured zero.

The staging checkpoint remains inside the unchanged native commit timer. Fresh matched builds are therefore required; allocator timings are diagnostics. These phases exclude final returned-object destruction, fixtures, oracles, and report work outside their boundaries. See the [design](results/change-0557/alloc-design.md), [final source review](results/change-0557/final-review.md), and [integration test](../../tools/perf-baseline/tests/xlsx_planning_allocations.rs).

## Fresh pilot

The [frozen plan](results/change-0557/plan.json) used CPU 2, a fresh child for each case/shape, 20 warmups and 1,000 samples per run, and two repeats of all eight primary rows: **16 successful children and 16,000 measured samples**. Source, locks, binary, commands, environment, corpus/output identities, and per-child RSS sidecars are retained. The [analyzer](results/change-0557/analyze.py) validated the complete pilot.

| Edit workload | Shape | R1 p50 (ms) | R2 p50 (ms) | Change |
| --- | --- | ---: | ---: | ---: |
| One cell | medium | 5.361466 | 5.348315 | -0.2453% |
| One cell | dense-sparse | 35.068957 | 35.340346 | +0.7739% |
| One cell | noncompact | 6.524370 | 5.885817 | -9.7872% |
| One cell | vendor-extension | 5.354064 | 5.390660 | +0.6835% |
| One percent | medium | 20.229965 | 20.205582 | -0.1205% |
| One percent | dense-sparse | 39.110999 | 39.090805 | -0.0516% |
| One percent | noncompact | 22.275983 | 22.302837 | +0.1206% |
| One percent | vendor-extension | 20.296597 | 20.278280 | -0.0902% |

The noncompact one-cell row determines the stop. Its publication median changed from 2.395266 ms to 1.803568 ms (−24.70%), while commit changed from 2.511902 ms to 2.487496 ms (−0.97%). These are separate marginal medians, not additive quantiles. The cause of the same-binary variability is unestablished. It cannot be attributed to an unapplied candidate.

The [noise analysis](results/change-0557/noise-analysis.json) retains all eight paired p50 comparisons. The supplemental [individual drift review](results/change-0557/noise-review.md) lists all **135 absolute changes above 5% across 464 elapsed/phase-distribution and process-sidecar comparisons**. It was added after capture for descriptive review and does not alter the frozen gate. All five native allocation vectors are explicitly unavailable.

The candidate gate is stopped. No noise rescue, native candidate matrix, allocator performance matrix, candidate profile, or adoption test followed the unstable pilot. The 0556 candidate remains preparation only; this pilot does not establish whether that candidate is faster or slower.

## Validation and custody

The harness library suite passed 498 tests, with one existing opt-in real-producer security test ignored and zero failures. Five allocator-binary tests and the focused XLSX allocation integration test passed. Warning-denied all-target/all-feature Clippy, warning-denied harness rustdoc, formatting, the all-feature workspace check, crate-boundary checks, and all 10 strict registry claims passed. Production XLSX source remains byte-identical to the baseline already tested in 0556. No fresh fuzz, sanitizer, native Office security-corpus, cold-cache, cross-platform, or scaling result is claimed.

The independent [pilot audit](results/change-0557/results-review.md) reproduced the stop and all diagnostic flags.

The [execution record](results/change-0557/execution-notes.md) retains the failed preflight attempts and prospective analyzer corrections. [Frozen inputs](results/change-0557/frozen-inputs.json) precede the release build and pilot. Accessible compiler-process observations do not prove host-wide quiescence. The hardware-counter availability probe succeeded; it is not a profile.

The [cleanup record](results/change-0557/cleanup.json) confirms removal of the owned build directory after checking accessible process references. Binary identity and raw evidence remain retained.

Replay the retained pilot with:

```sh
python3 -B docs/performance/results/change-0557/analyze.py --noise-only --output docs/performance/results/change-0557/noise-analysis.json
python3 -B docs/performance/results/change-0557/review_noise.py
```
