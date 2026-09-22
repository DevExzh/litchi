# 0739 — current PPTX cross-slide-copy lifecycle baseline

Planning remains a material current target: on the generated media-rich deck,
its median within-sample share is **71.06%** of the owned lifecycle; application
is **22.00%**. Median process lifecycle p50 is **403.95 ms**. On the plain deck,
the lifecycle is **8.33 ms**, with planning/application shares of 39.44%/45.44%.
These are fresh unchanged-source observations, not an optimization result or
a comparison with the historical 0656 build.

The next bounded measurement should partition current planning into closure
proof, candidate graph construction, serialization, reopen/capture and patch
capture before proposing proof reuse. The outer planning clock cannot identify
any one of those costs or establish how much can be removed. All live-source,
physical-provenance, revision, budget, refusal and publication proofs remain
required. Any internal timing probe needs its own observer qualification with
matched invocations and unchanged output checks; current outer shares cannot
serve as a cross-build comparator. [Source review](results/change-0739/source-review.md).

## Why this path now

The 0727 queue selected fresh DOC/PPT writer attribution first. 0730 and 0734
retained bounded ownership handoffs; 0735 rejected the next PPT staging change.
0738 established startup sensitivity that must be controlled before further PPT
timing transfer. This batch advances the queue's independent PPTX cross-copy
opportunity rather than reopening that rejected candidate.

Current `apply_plan` still fingerprints, captures, replans and validates. Its
retained archive removes the second serialization, but `build_candidate` still
reconstructs the candidate graph, copies retained archive bytes, reopens and
captures. The older namespace-emission and DOC identity-scan opportunities from
0587 already landed in 0653/0659 and are not fresh queue items. These source
facts motivate measurement, not a shortcut around existing proof obligations.

## Workload and boundary

The existing Rust harness and production are unchanged at `65f82b0f8b`. Both
release binaries were freshly built offline/locked with two Cargo jobs on AMD
EPYC 9R45, Linux, Rust 1.95.0. The packet binds 7,281 crate/harness source and
manifest files, four supplementary workspace inputs, and 34 goal/checklist/ADR
hashes. Native and allocator executables are separate.

Each of two generated corpora has nine fresh native processes, 30 measured
samples after three warmups. Six separate allocation processes take one sample
without warmup. Cases alternate order by repeat; all children run serially on
CPU 12. Four qualification processes precede the frozen 24-process matrix.
All 540 native samples are retained. No timing-based exclusions or retries occur.

The lifecycle includes owned source/destination ingress, opened snapshots,
planning, application and final sequential publication. Corpus cloning and sink
reservation precede it. Post-operation reopen, semantic/preservation checks and
teardown follow it. Production validation inside plan/apply remains timed.
Plan/apply/publication are nested diagnostics; `unassigned` is total minus those
three intervals, including ingress/snapshot work and timing/call overhead.
Reopen is outside the total. Shares below are medians of per-sample ratios,
then medians across processes; separately summarized medians need not add up.

## Native observations

Times are milliseconds. Central values are medians of nine process p50s;
intervals bootstrap those process medians with 10,000 draws, seed 7339. Tails
are medians of per-process nearest-rank p95/p99. At 30 samples p99 equals max.
No pooled-sample confidence claim is made.

| Corpus / phase | p50 [bootstrap 95% CI] | p95 | p99 = max | Median lifecycle share |
|---|---:|---:|---:|---:|
| Plain / lifecycle | 8.3342 [8.3082, 8.3769] | 8.5328 | 8.5656 | 100% |
| Plain / plan | 3.2895 [3.2728, 3.3059] | 3.3978 | 3.4394 | 39.44% |
| Plain / commit | 3.8020 [3.7692, 3.8142] | 3.8827 | 3.9365 | 45.44% |
| Plain / publication | 0.0008 [0.0008, 0.0008] | 0.0008 | 0.0008 | 0.01% |
| Plain / unassigned | 1.2616 [1.2573, 1.2639] | 1.2760 | 1.2876 | 15.11% |
| Plain / reopen | 0.6101 [0.6077, 0.6122] | 0.6191 | 0.6215 | outside |
| Media-rich / lifecycle | 403.9453 [402.7684, 404.2773] | 404.9013 | 405.0190 | 100% |
| Media-rich / plan | 286.9387 [286.4106, 287.2112] | 287.6934 | 287.9770 | 71.06% |
| Media-rich / commit | 88.8279 [88.6876, 88.9331] | 88.9853 | 89.0818 | 22.00% |
| Media-rich / publication | 5.9763 [5.8690, 6.2136] | 6.1083 | 6.1122 | 1.48% |
| Media-rich / unassigned | 22.0252 [21.9069, 22.0653] | 22.1530 | 22.2155 | 5.46% |
| Media-rich / reopen | 21.0040 [20.9482, 21.0222] | 21.1120 | 21.1164 | outside |

Lifecycle process-p50 spans are 1.30% for plain and 2.31% for media-rich.
Publication process-p50 spans are 8.00% and 318.56%, respectively. The media
publication medians range from roughly 1.50 to 6.29 ms. This recurrence of
variable publication cost has no mechanism established here. All 14 metric
spread flags above 5% are listed below and in the raw analysis. No process has
an absolute last-ten/first-ten lifecycle median change above 5%. These are
descriptive stability flags, not candidate acceptance gates.

| Corpus | Phase / statistic | Across-process max/min spread |
|---|---|---:|
| Plain | plan_ns / p99 | 5.07% |
| Plain | plan_ns / maximum | 5.07% |
| Plain | publication_ns / p50 | 8.00% |
| Plain | publication_ns / mean | 21.90% |
| Plain | publication_ns / p95 | 7.59% |
| Plain | publication_ns / p99 | 433.75% |
| Plain | publication_ns / maximum | 433.75% |
| Media-rich | publication_ns / p50 | 318.56% |
| Media-rich | publication_ns / mean | 322.89% |
| Media-rich | publication_ns / p95 | 327.42% |
| Media-rich | publication_ns / p99 | 332.85% |
| Media-rich | publication_ns / maximum | 332.85% |
| Media-rich | unassigned_ns / p99 | 5.80% |
| Media-rich | unassigned_ns / maximum | 5.80% |

![All process medians](results/change-0739/phase-medians.png)

## Allocation observations

The table gives the range across three fresh processes per corpus; every field
is retained, including global high-water snapshots that may predate the owner.
The operation region covers the lifecycle, not individual phases. Entry and
exit live bytes include corpus/setup retention. `region_peak_live_bytes`
includes entry live bytes and excludes allocator-internal realloc overlap;
it is not RSS or the whole-child peak. Native latency is not inferred from the
instrumented binary.

| Allocator field | Plain min–max | Media-rich min–max |
|---|---:|---:|
| allocation_calls | 46,613 | 56,356–56,357 |
| deallocation_calls | 38,273 | 46,067–46,068 |
| reallocation_calls | 5,298 | 6,271 |
| failed_allocation_calls | 0 | 0 |
| allocated_bytes | 16,613,707 | 277,807,472–277,810,624 |
| deallocated_bytes | 16,100,040 | 176,507,837–176,510,989 |
| live_bytes_before | 347,373 | 169,800,135 |
| live_bytes_after | 861,040 | 271,099,770 |
| peak_live_bytes_before | 3,688,443 | 862,425,419 |
| peak_live_bytes_after | 3,688,443 | 862,425,419 |
| region_peak_live_bytes | 1,311,524 | 305,066,939 |

## Verification and scope

All captures match the qualification's exact corpus, metadata, sink and output
projection. Each measured iteration repeats full selected semantic/topology/
closure checks, source-byte immutability and byte equality with the generated
expected archive. Borrowed/stale/foreign refusal controls run once during corpus
construction per process. The expected output and oracle use the same Litchi
stack. The owned helper compares logical member payloads and relationships;
it does not independently verify physical ZIP metadata preservation. No native
Office or real-producer compatibility claim follows.

Both builds and three documentation/coverage structural gates pass. These gates
validate registries and the coverage contract, not a new full CRUD timing run.
Twelve deliberate report corruptions are rejected; four
additional controls reject altered medians, confidence intervals, process counts
and commands. The independent replay recomputes statistics and group summaries
from raw reports while sharing the invocation/report validator. All derived
statistics match within its explicit floating-point tolerance. Production Rust
quality suites are not relabeled as fresh runs for this unchanged-source batch.

The default CRUD index remains unchanged. Its cross-document rows cover
source-backed selectors and a non-lifecycle media selector; these owned
lifecycle observations do not silently promote those different scenarios.
Physical cold-cache, remote/range, concurrency scaling, RSS, producer variation,
retention lifetime sweeps and comprehensive CRUD completion remain unproved.
The non-iWork goal stays active.

Owned build files are removed only after matching both recorded executable
hashes. Post-cleanup report replay and exact artifact inventory verification
remain available through the [packet README](results/change-0739/README.md).
The unrelated API-design draft is preserved.
