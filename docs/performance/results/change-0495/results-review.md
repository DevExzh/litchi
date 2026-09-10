# Change 0495: formal measurement results review

## Result

The formal managed-edit run passed its receipt and scalar verification. It demonstrates successful source-backed one-paragraph editing and publication for the six provider arms, both normal and allocator roles, and both unmanaged and managed APIs in the after phase. It does not authorize a performance claim: `claim_authorized` is `false`, and the managed API has no before baseline. Before/after performance review is therefore limited to matched unmanaged rows.

The review retains all eleven paired unmanaged cells that crossed the five-percent review threshold: eight whole-child RSS cells and three latency cells. These are follow-up flags, not causal attribution or proof that the production change caused a regression. The profile run is whole-child evidence as well, so it can guide investigation but cannot assign the RSS deltas to the measured operation.

Formal evidence is in [`analysis/formal1.json`](analysis/formal1.json) and [`verification/formal1.json`](verification/formal1.json). Verification status is `pass`; the analysis schema is `docx-edit-provider-managed-analysis-v1`, and the protocol SHA-256 is `55d417150a998a992c8e6bf3f724041b52569f17c75126517a8579f80f52186a`. The formal run used two repeats, three warmups, and thirty measured samples per process: 72 child processes and 2,160 measured samples. The pilot used one warmup and three samples per process, with 36 processes and 108 samples, and also passed verification.

## Measurement boundary and custody

Each process opened the source archive, staged one paragraph edit, committed it, published sequentially to the sink, and then dropped the package, returned snapshot, and commit. Provider construction, sink reservation, and execution-context teardown were outside the timed interval. Commit diagnostics and source/candidate XML comparison were inside it. Output hashing, semantic/media verification, and preflight patch oracles were outside it. The `whole_child_rss_kbytes` value covers the complete child, including setup and post-clock work; it is not operation-local RSS.

The formal matrix used roles `normal` and `allocator`, phases `before` and `after`, APIs `unmanaged-api` and `managed-api`, and these arms: `owned`, `instrumented`, `file-warm`, `short`, `delayed`, and `range-zero`. The delayed arm used a 100 MiB/s service rate with a 1,000 microsecond delay; `range-zero` used the same 65,536-byte range limit and service rate with zero fixed delay. Repeat two reversed the phase order as required by the protocol.

The corpus was a 16,793,036-byte, 20-member archive with eight 2 MiB media members and 200 paragraphs. Its SHA-256 was `a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4`. The expected edited output was 16,793,048 bytes with SHA-256 `9af99bf7f63aac1ffc13ff59a5038703c96229b906b3db1602d5af5590171795`. The source text and expected edited text were checked independently by the receipt verifier.

All builds used revision `de8ee88b0727ae59e4d2b4c8b8a6c24349724ae8`. The before source manifest SHA-256 was `96d5c8bbcab21dd2f10fd992a96d90bd9ecb1f101bad1a8f5e9ece2c18f8651f`; the after manifest SHA-256 was `0b52dcabee0edda63f06a1e28e8c3f70a80e7e41eed85e59299a024236953982`. Retained normal and allocator binary hashes were recorded in the build receipts and are reproduced here for custody:

| phase | role | binary bytes | binary SHA-256 |
| --- | --- | ---: | --- |
| before | normal | 476,661,024 | `03e98d09034d597ad563519672c26fa90ffbb826c8e6d54ffee387ddb4b3b4b8` |
| before | allocator | 476,401,240 | `4893309b477d0aaddeb1f8c84a7d998671df732e1e929a886436e6e86021fc54` |
| after | normal | 478,225,064 | `d94d8cac8bc1628148f54de67531f003ec2d1180b5a81bbc1b7c0c47dc13a823` |
| after | allocator | 477,975,216 | `450c89732720d89c2bb53b652e308aec9d335e466fa9e3a5b0900329b52aec60` |

The receipt hashes were `3d5604eafb910ca537bed149e61df9fbded51eb51ef63af7a8a094b636f4dcea` (before normal), `a853370d0edafafd7bbe2eae2baa025aa2e4b750614d556c9b7e953c875f1057` (before allocator), `ab756d5205a3f3dad391128479f63e21eee53083b5c9c833ebfb8a9723df48e6` (after normal), and `6c4e2061b03ffdd347e60a5f61e7659a0b7ad6add74705c09ec20a557a204919` (after allocator).

## Receipt and conservation checks

An independent scalar pass inspected all 72 reports and 2,160 sample rows. It found zero row failures and zero terminal failures. Every child exited successfully without timeout or termination, and every cleanup receipt passed. Every row produced the same 16,793,048-byte output and the same output SHA-256 shown above. Every sink accepted exactly the expected output.

The independent pass recomputed source range bounds, logical and physical requested/returned sums, short-read counts, media overlap, sink accounting, managed budget transitions, and allocator live-byte conservation. All recomputed values passed. The 720 managed rows all reported `memory_released_to_baseline=true`, `objects_released_to_baseline=true`, and zero reservation failures.

The source and sink summaries below are sums over each sample's read or write calls. `owned` source counters are intentionally unavailable; the other arms expose the relevant trace. Physical traces matched the available logical traces where the protocol requested them.

| phase/API | source calls | requested bytes | returned bytes | short reads | media-overlap requested/returned |
| --- | ---: | ---: | ---: | ---: | ---: |
| before/unmanaged, non-short | 377 | 16,799,430 | 16,799,430 | 0 | 16,782,376 / 16,782,376 |
| before/unmanaged, short | 4,217 | 142,628,550 | 16,799,430 | 3,840 | 142,611,496 / 16,782,376 |
| after/unmanaged, non-short | 381 | 16,800,458 | 16,800,458 | 0 | 16,782,376 / 16,782,376 |
| after/unmanaged, short | 4,221 | 142,629,578 | 16,800,458 | 3,840 | 142,611,496 / 16,782,376 |
| after/managed, non-short | 389 | 16,801,140 | 16,801,140 | 0 | 16,782,376 / 16,782,376 |
| after/managed, short | 4,229 | 142,630,260 | 16,801,140 | 3,840 | 142,611,496 / 16,782,376 |

All 2,160 samples wrote 16,793,048 accepted bytes in 339 sink writes; the largest write was 65,536 bytes. The counters include the provider's requested ranges and are not claims about external filesystem or network I/O. For the paced delayed and range-control arms, the adapter reported 389 delayed and paced calls with 160,228,229 ns of modeled transfer delay in managed rows; the corresponding after-unmanaged values were 381 calls and 160,221,721 ns, and before-unmanaged values were 377 calls and 160,211,915 ns.

## Managed capability and retention

The after managed API completed all six arms in both roles and both repeats: 24 processes and 720 samples. Each row produced the expected output, stayed within the finite execution limits, released memory and objects to its baseline, and recorded zero reservation failures. This is capability evidence. There is no before managed row, so it cannot support a before-managed speed, RSS, or allocation comparison.

The following values were identical across the managed rows. The cache snapshots and resource snapshots are different lifecycle observations; they must not be subtracted as if they were the same gauge.

| snapshot | memory | input bytes | output bytes | objects | work | cache entries | cache bytes | cache failures |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `before` resource | 0 | 0 | 0 | 0 | 0 | — | — | — |
| `cache_before` | 0 | 2,480 | 0 | 21 | 0 | 0 | 0 | 0 |
| `cache_live` | 1,454,238 | 3,966 | 0 | 58,066 | 50,305,515 | 1 | 29,027 | 0 |
| `live` resource | 1,454,238 | 16,801,140 | 16,793,048 | 58,045 | 100,494,387 | — | — | 0 |
| `after_drop` resource | 0 | 16,801,140 | 16,793,048 | 0 | 100,494,387 | — | — | 0 |

The managed ceilings were 67,108,864 bytes for memory, input, and output, 1,000,000 objects, depth 1,024, and 2,147,483,648 work units. The cache ceilings were 8 MiB and 128 entries. The largest live values were therefore below every configured ceiling. The live resource memory and objects returned to zero after package/document/commit drops; input, output, and work are cumulative charges and correctly remain nonzero.

`cache_live` is captured after edit and commit staging but before publication. It is a post-edit/pre-publication cache diagnostic, despite its name. The run has no post-publication cache snapshot, so it does not prove that the cache itself was released. The resource `after_drop` gauge does cover package consumption, returned-snapshot release, and commit release. The one retained cache entry and 29,027 retained bytes are bounded and coincide with the successful managed source-backed operation. The retention decision is to keep this necessary ordinary managed-edit ownership enabler for the current implementation; the formal run provides no causal evidence that removing it would preserve capability or improve RSS. A follow-up that needs a cache-retention claim should add a distinct post-publication cache gauge or a controlled cache-lifetime experiment, while preserving the current resource-release checks.

## Paired unmanaged before/after flags

The analysis compares only `before/unmanaged-api` with `after/unmanaged-api`, matched by repeat, role, and arm. A flag means the after-minus-before value is greater than five percent at at least one of p50, p95, or p99. Percentages below are rounded to three decimals; latency is in nanoseconds and RSS is whole-child high-water RSS in KiB.

There are eight RSS flags and three latency flags:

| repeat | role | arm | metric | before p50/p95/p99 | after p50/p95/p99 | after-minus-before p50/p95/p99 |
| ---: | --- | --- | --- | --- | --- | --- |
| 1 | normal | owned | RSS KiB | 133,612 / 133,612 / 133,612 | 149,984 / 149,984 / 149,984 | +12.253% / +12.253% / +12.253% |
| 1 | normal | instrumented | RSS KiB | 134,132 / 134,132 / 134,132 | 150,848 / 150,848 / 150,848 | +12.462% / +12.462% / +12.462% |
| 1 | normal | file-warm | latency ns | 2,267,383 / 2,309,818 / 2,356,348 | 4,394,731 / 4,630,277 / 4,665,016 | +93.824% / +100.461% / +97.977% |
| 1 | normal | delayed | RSS KiB | 135,896 / 135,896 / 135,896 | 149,592 / 149,592 / 149,592 | +10.078% / +10.078% / +10.078% |
| 1 | normal | range-zero | latency ns | 184,778,723 / 185,920,977 / 186,059,818 | 181,521,252 / 185,855,464 / 199,249,811 | -1.763% / -0.035% / +7.089% |
| 1 | allocator | instrumented | RSS KiB | 133,556 / 133,556 / 133,556 | 149,780 / 149,780 / 149,780 | +12.148% / +12.148% / +12.148% |
| 1 | allocator | short | RSS KiB | 136,692 / 136,692 / 136,692 | 150,748 / 150,748 / 150,748 | +10.283% / +10.283% / +10.283% |
| 2 | normal | owned | RSS KiB | 136,476 / 136,476 / 136,476 | 150,840 / 150,840 / 150,840 | +10.525% / +10.525% / +10.525% |
| 2 | normal | short | RSS KiB | 139,440 / 139,440 / 139,440 | 149,936 / 149,936 / 149,936 | +7.527% / +7.527% / +7.527% |
| 2 | normal | range-zero | RSS KiB | 136,108 / 136,108 / 136,108 | 148,184 / 148,184 / 148,184 | +8.872% / +8.872% / +8.872% |
| 2 | allocator | short | latency ns | 3,411,827 / 5,834,202 / 5,863,132 | 5,542,995 / 5,642,101 / 5,669,541 | +62.464% / -3.293% / -3.302% |

The p99-only range-zero latency flag and p50-only allocator-short latency flag should not be summarized as broad latency regressions. The file-warm flag is large in repeat one but absent in repeat two's paired comparison, which is consistent with the repeat-variance findings below. The RSS flags are measurements of complete child high-water usage; setup, allocator startup, provider construction, output verification, and report serialization can contribute.

## Repeat variance flags

The protocol also reports absolute repeat-to-repeat changes over five percent. These are descriptive instability flags within one phase/API and are not before/after optimization claims. The table lists every flagged latency or RSS metric. Percentages are repeat-two minus repeat-one, rounded to three decimals; each quantile triple is p50/p95/p99.

| phase | API | role | arm | metric | repeat 1 p50/p95/p99 | repeat 2 p50/p95/p99 | repeat 2 minus repeat 1 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| before | unmanaged | normal | instrumented | RSS KiB | 134,132 / 134,132 / 134,132 | 151,764 / 151,764 / 151,764 | +13.145% / +13.145% / +13.145% |
| before | unmanaged | normal | file-warm | latency ns | 2,267,383 / 2,309,818 / 2,356,348 | 4,432,102 / 4,666,438 / 4,696,848 | +95.472% / +102.026% / +99.327% |
| before | unmanaged | normal | file-warm | RSS KiB | 153,388 / 153,388 / 153,388 | 136,888 / 136,888 / 136,888 | -10.757% / -10.757% / -10.757% |
| before | unmanaged | normal | delayed | latency ns | 561,908,009 / 565,492,788 / 566,492,471 | 562,328,436 / 566,483,899 / 602,587,804 | +0.075% / +0.175% / +6.372% |
| before | unmanaged | normal | delayed | RSS KiB | 135,896 / 135,896 / 135,896 | 149,412 / 149,412 / 149,412 | +9.946% / +9.946% / +9.946% |
| before | unmanaged | normal | range-zero | latency ns | 184,778,723 / 185,920,977 / 186,059,818 | 185,008,782 / 189,510,925 / 195,970,888 | +0.125% / +1.931% / +5.327% |
| before | unmanaged | normal | range-zero | RSS KiB | 150,116 / 150,116 / 150,116 | 136,108 / 136,108 / 136,108 | -9.331% / -9.331% / -9.331% |
| before | unmanaged | allocator | instrumented | RSS KiB | 133,556 / 133,556 / 133,556 | 147,244 / 147,244 / 147,244 | +10.249% / +10.249% / +10.249% |
| before | unmanaged | allocator | short | latency ns | 5,826,990 / 5,873,520 / 5,883,241 | 3,411,827 / 5,834,202 / 5,863,132 | -41.448% / -0.669% / -0.342% |
| before | unmanaged | allocator | short | RSS KiB | 136,692 / 136,692 / 136,692 | 153,136 / 153,136 / 153,136 | +12.030% / +12.030% / +12.030% |
| before | unmanaged | allocator | delayed | RSS KiB | 149,484 / 149,484 / 149,484 | 133,048 / 133,048 / 133,048 | -10.995% / -10.995% / -10.995% |
| after | unmanaged | normal | owned | latency ns | 4,512,086 / 4,581,256 / 4,597,107 | 2,142,008 / 2,166,978 / 2,177,598 | -52.527% / -52.699% / -52.631% |
| after | unmanaged | normal | file-warm | latency ns | 4,394,731 / 4,630,277 / 4,665,016 | 4,647,592 / 4,719,357 / 4,774,337 | +5.754% / +1.924% / +2.343% |
| after | unmanaged | normal | file-warm | RSS KiB | 148,332 / 148,332 / 148,332 | 133,708 / 133,708 / 133,708 | -9.859% / -9.859% / -9.859% |
| after | unmanaged | normal | short | RSS KiB | 137,732 / 137,732 / 137,732 | 149,936 / 149,936 / 149,936 | +8.861% / +8.861% / +8.861% |
| after | unmanaged | normal | delayed | RSS KiB | 149,592 / 149,592 / 149,592 | 135,232 / 135,232 / 135,232 | -9.599% / -9.599% / -9.599% |
| after | unmanaged | normal | range-zero | latency ns | 181,521,252 / 185,855,464 / 199,249,811 | 184,977,150 / 188,326,048 / 188,535,259 | +1.904% / +1.329% / -5.377% |
| after | unmanaged | allocator | owned | latency ns | 4,513,581 / 4,723,987 / 4,749,847 | 4,741,397 / 4,776,868 / 4,800,688 | +5.047% / +1.119% / +1.070% |
| after | unmanaged | allocator | file-warm | latency ns | 2,473,704 / 2,507,829 / 2,510,659 | 4,554,917 / 4,758,448 / 4,789,278 | +84.133% / +89.744% / +90.758% |
| after | unmanaged | allocator | delayed | RSS KiB | 150,752 / 150,752 / 150,752 | 132,676 / 132,676 / 132,676 | -11.991% / -11.991% / -11.991% |
| after | managed | normal | file-warm | RSS KiB | 134,184 / 134,184 / 134,184 | 148,848 / 148,848 / 148,848 | +10.928% / +10.928% / +10.928% |
| after | managed | allocator | owned | latency ns | 6,022,691 / 6,098,372 / 6,106,531 | 3,694,758 / 6,043,422 / 6,067,522 | -38.653% / -0.901% / -0.639% |
| after | managed | allocator | range-zero | latency ns | 186,994,979 / 196,871,315 / 214,634,689 | 182,334,508 / 183,725,278 / 183,885,708 | -2.492% / -6.677% / -14.326% |
| after | managed | allocator | range-zero | RSS KiB | 133,032 / 133,032 / 133,032 | 149,412 / 149,412 / 149,412 | +12.313% / +12.313% / +12.313% |

The repeat flags leave causality unresolved: several large changes move in opposite directions across repeats, while output, sink, source-return totals, and conservation values remain fixed. The observed adverse changes must not be dismissed as noise; targeted repeats and phase-local attribution are required to distinguish production costs from host and allocator variation.

## Allocation deltas

Allocation counters are available in the allocator role. Each value below was constant across all six arms and both repeats within its phase/API group. The before-to-after comparison is unmanaged only; managed is compared with after-unmanaged as a descriptive API delta and has no before-managed baseline.

| group | allocation calls | realloc calls | allocated/deallocated bytes | peak live increment |
| --- | ---: | ---: | ---: | ---: |
| before, unmanaged | 22,859 | 185 | 5,721,334 | 606,959 |
| after, unmanaged | 9,696 | 228 | 1,623,696 | 609,903 |
| after, managed | 20,535 | 328 | 4,270,196 | 622,454 |

| comparison | allocation calls | realloc calls | allocated/deallocated bytes | peak live increment |
| --- | ---: | ---: | ---: | ---: |
| after unmanaged minus before unmanaged | -13,163 (-57.583%) | +43 (+23.243%) | -4,097,638 (-71.620%) | +2,944 (+0.485%) |
| after managed minus after unmanaged | +10,839 (+111.788%) | +100 (+43.860%) | +2,646,500 (+162.992%) | +12,551 (+2.058%) |

The managed API therefore has a substantial allocation-counter delta relative to the after unmanaged API, while its peak live increment changes by 2.058 percent. These are allocator-counter observations, not a managed-before performance claim. Absolute region peaks include each process's pre-existing allocator state; the corresponding absolute peak changes ranged from +2,941 to +5,725 bytes for before-to-after unmanaged and +12,753 to +18,321 bytes for after managed versus after unmanaged, depending on arm. All rows satisfied allocator live-byte conservation and none reported a leak.

## Profile follow-up

The completed profile attempt is [`profiles/profiles1/profile-summary.json`](profiles/profiles1/profile-summary.json). It covered `owned`, `file-warm`, and `short` in both unmanaged and managed after APIs, with three samples and one warmup per case. All six cases completed cleanup successfully. The full perf event set was degraded because `LLC-loads` and `LLC-load-misses` were unsupported on the host; the core event set and strace artifacts were available.

The profile observer covers the whole child, including startup, setup/preflight, the measured operation, post-clock output hashing/semantic verification, and report serialization. Publication markers are stack subsets, not independent phase durations. Use the profile artifacts to investigate the file-warm and allocator/short instability, and add phase-local instrumentation before attributing any RSS change. The current evidence does not justify removing the managed cache or calling an optimization.

## Acceptance reading

This run supports the following bounded statements:

- The six-provider, two-role matrix produced verified output for 72 processes and 2,160 formal samples; all raw rows passed source, sink, output, budget, and allocator conservation checks.
- The after managed API successfully performed the measured source-backed edit under finite limits and released its resource memory and objects to baseline.
- The managed cache observed one retained entry and 29,027 bytes at the pre-publication cache checkpoint; that bounded retention remains a measured enabler, with its lifetime after publication still unmeasured.
- The paired unmanaged review has eight RSS flags and three latency flags listed above; they remain follow-up items, not causal regression findings.
- There is no before-managed baseline, so this report makes no before-managed speed, RSS, allocation, or optimization claim.

The formal result is descriptive evidence for the current provider and source-backed edit path. A future performance claim needs a controlled phase-local design that separates child setup and verification from the operation, repeats the unstable file-warm and allocator/short cases, and adds a post-publication cache observation if retention cost is part of the claim.

## Evidence links

- [`analysis/formal1.json`](analysis/formal1.json)
- [`verification/formal1.json`](verification/formal1.json)
- [`publication-review.md`](publication-review.md), for the timing and cache lifecycle interpretation
- [`profiles/profiles1/profile-summary.json`](profiles/profiles1/profile-summary.json)
