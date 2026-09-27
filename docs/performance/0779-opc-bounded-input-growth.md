# 0779 — bound OPC input-buffer growth

Status: **candidate rejected; production source restored byte-for-byte** to
`345b81ce8a`. The evidence and current-source queue correction are retained.

Ordinary XLSX open uses the shared OPC owned-input reader. That reader kept
an 8 KiB I/O buffer but requested exact `Vec` growth after each chunk. The
4,226,429-byte generated workbook therefore requested 1,092,697,469 cumulative
capacity bytes before ZIP or workbook parsing. Direct baseline phase counters
confirm 1,093,823,950 requested bytes during open and 1,100,160,826 during the
open/edit/NoSync-save lifecycle. The separate real workbook open requests
33,020,061 bytes. Allocation requests include the entire new size of realloc;
they do not establish physical bytes copied or resident memory.

The rejected candidate kept read requests and the input ceiling unchanged. It grew
capacity only when necessary, by the larger of 8 KiB and one eighth of current
capacity, capped by the existing limit. Reservations remained fallible. It added
no metadata hint, cache, network access, concurrency or unsafe code. The chosen
factor limited retained spare capacity relative to doubling, but the measured
memory cost outweighed the small lifecycle latency differences.

## Current-source correction

The next-step recommendation in the sealed 0778 packet was stale: ZIP first-read
coalescing already landed in 0611, followed by structural prefetch in 0623 and
directory prefill in 0632. The current implementation confirms those changes.
[Queue correction](results/change-0779/queue-correction.md) supersedes that
planning note without rewriting its sealed historical evidence. This batch
investigates a different, currently present shared-ingress cost.

## Method and scope

[The frozen plan](results/change-0779/plan.json) uses six alternating native
process blocks, 30 samples after three warmups, on the generated and real XLSX
sources retained by 0778. It measures open, edit, save and lifecycle separately.
Two separate allocator blocks use three samples without warmups. All owner
destruction and semantic/readback checks are outside the relevant region.
The probe calls public Workbook APIs, retains source/output hashes, reopens
saved outputs and checks the A1 marker. Expected published bytes are also
bound to 0778's independently checked outputs.

Save is explicitly NoSync in both legs to isolate ingestion costs. Ordinary
Full durability is unchanged. This is warm-source, absent-destination evidence;
it does not cover physical cold cache, remote transport, concurrency or large
corpora. Whole-process RSS includes setup and post-clock verification. Region
allocation peaks cannot be added across phases, and native timing is kept
separate from allocator and heaptrack observations.

## Results

The fixed capture completed 120 native processes (3,600 timed samples) and 40 separate allocator processes (120 samples), with no result-driven reruns. The primary matrix has eight case/phase pairs; two supplemental open controls cover 3,555-byte and 8,224-byte producer files.

### Native timing and RSS

Absolute times below are medians of six process p50s, with the middle two averaged. Each process p50 uses nearest-rank over 30 samples. Percent changes are the median of the six **paired block ratios**, so they need not equal the ratio of the displayed absolute medians. RSS uses whole-process peak KiB.

| Corpus / phase | Before p50 ms | Candidate p50 ms | Paired p50 change | Paired RSS change |
|---|---:|---:|---:|---:|
| conditional-formatting/edit | 1.5300 | 1.5201 | -0.707% | +0.798% |
| conditional-formatting/lifecycle | 3.2559 | 3.2438 | -0.354% | +0.837% |
| conditional-formatting/open | 1.1564 | 1.1489 | -0.619% | +0.802% |
| conditional-formatting/save | 0.4891 | 0.4890 | +0.067% | +1.482% |
| generated-xlsx/edit | 2.2909 | 2.2772 | -0.474% | +0.263% |
| generated-xlsx/lifecycle | 4.8060 | 4.7789 | -0.581% | +20.203% |
| generated-xlsx/open | 0.4055 | 0.3989 | -1.501% | +2.937% |
| generated-xlsx/save | 1.9755 | 1.9750 | +0.338% | +20.242% |
| boundary-xlsx/open | 0.0646 | 0.0648 | +0.202% | +2.630% |
| small-xlsx/open | 0.0566 | 0.0568 | +0.503% | +2.290% |

Generated-workbook save RSS medians increase from 17,834 to 21,382 KiB; lifecycle medians increase from 17,746 to 21,408 KiB. Their paired RSS changes are +20.24% and +20.20%. These whole-process gauges include open, post-clock reopen and allocator behavior; they are not attributed entirely to the 323 KB extra source capacity.

All flags remain visible: 11 primary and 12 supplemental native spread flags; nine primary and seven supplemental paired metric-series flags. A paired series is flagged if **any block** regresses over 5%, even when its median improves. Generated-open p50 has one +11.51% block among five lower candidate values. Small controls also have noisy tails and RSS pairs. See [all process summaries](results/change-0779/native-processes.csv), [all metric distributions](results/change-0779/native-summary.csv), [every paired block](results/change-0779/native-pairs.csv), and [all flags](results/change-0779/all-flags.csv). No flagged process was dropped or rerun. Process spread is `(max − min) / min`; six blocks provide descriptive uncertainty, not a precise general-population speedup estimate.

### Allocation tradeoff

The table uses operation-region counters. Peak-above-entry and net-live change subtract each sample's entry gauge; they avoid the one-byte process-argument offset between the before/after executables. Values are repeated p50s from the two allocator processes.

| Corpus / phase | Requested bytes before → candidate | Realloc calls before → candidate | Peak above entry before → candidate | Net live before → candidate |
|---|---:|---:|---:|---:|
| conditional-formatting/lifecycle | 39,278,222 → 18,661,545 | 2,497 → 2,445 | 2,390,754 → 2,427,131 | 1,287,593 → 1,323,970 |
| conditional-formatting/open | 33,020,061 → 12,403,384 | 1,014 → 962 | 1,342,090 → 1,378,467 | 1,264,649 → 1,301,026 |
| generated-xlsx/lifecycle | 1,100,160,814 → 48,113,441 | 5,558 → 5,086 | 6,466,390 → 6,789,392 | 5,031,244 → 5,354,246 |
| generated-xlsx/open | 1,093,823,950 → 41,776,577 | 635 → 163 | 4,615,014 → 4,938,016 | 4,529,825 → 4,852,827 |
| boundary-xlsx/open | 587,981 → 596,141 | 76 → 76 | 114,896 → 123,056 | 34,010 → 42,170 |
| small-xlsx/open | 729,856 → 729,856 | 61 → 61 | 112,857 → 112,857 | 33,268 → 33,268 |

Generated open requests 96.18% fewer cumulative bytes but retains 323,002 more bytes: net live rises about 7.13%, and region peak above entry about 7.00%. Lifecycle net live rises about 6.42%. The 8,224-byte boundary case retains 8,160 more bytes (net live +23.99%), with no reduction in realloc calls. Allocation repeat p50s have zero spread; all 17 primary and six supplemental paired allocation flags are retained in the flag table. [All 13 allocator metrics](results/change-0779/allocation-summary.csv) include edit/save controls and absolute gauges.

The independently parsed heaptrack traces exactly reproduce their histograms. In **each** measured generated-file open, `read_limited` accounts for 1,092,697,469 → 40,650,096 requested bytes and 515 → 43 growth reallocations. Each whole process performs a second open outside the clock for verification, so its target totals are twice those values. Whole-process totals are 2,191,954,137 → 87,859,389 requested bytes. [Profile review](results/change-0779/profile-review.md) distinguishes the two call routes. This confirms the allocation site, not physical copy volume, a latent network bottleneck, or a comparable latency improvement.

### Disposition

Reject the one-eighth growth candidate. The small end-to-end latency differences do not justify the measured retained-memory and RSS costs on these representative cases. No geometric-growth change, new helper or candidate test is retained in production. The exact applied source, patch, six passing candidate gates and both binaries' custody records remain in the packet; the complete production source census is restored to the baseline.

This result closes the interpretation of 0778's large cumulative-allocation number as an unproven latency opportunity. Future work should follow current end-to-end profiles and memory evidence, not revive the already-landed ZIP-2 item or treat realloc request volume as copied bytes. It does not rule out a separately measured design with a better memory/latency tradeoff.


## Correctness and architecture

The [design](results/change-0779/design.md) maps ADRs 0001–0006, 0008, 0010,
0011 and 0024 to the scoped change. All 35 previously read architecture inputs
remain byte-identical to 0778. [Source review](results/change-0779/source-review.md)
and [candidate review](results/change-0779/candidate-review.md) retain the
arithmetic and the requested-capacity versus physical-memory distinction.

The rejected candidate’s focused tests preserved returned bytes and read request sizes through growth,
short reads and Interrupted retries; they preserve late I/O and invalid-count
errors, exact-limit/over-limit sentinel behavior and bounded capacity arithmetic.
Review removed an unrelated sentinel-expression change and corrected the draft
short-reader expectation before application. The candidate changed no reachable typed input refusal
or ZIP validation behavior. More generous capacity requests could fail allocation
at a different point; the same typed allocation error and resource name remained.
The candidate passed all six gates, with 3,079 tests passed and one ignored
across 119 suites. These are candidate results; restoration is verified by exact
source identity, not presented as a new production test run.

No CRUD row is promoted. The non-iWork performance goal remains active.


## Integration and cleanup

Integrated by fast-forward through `005f279fd5`, following the independent
Heaptrack attribution commit `d862e596c2`. All changes are performance
documentation and evidence; production source remains identical to
`345b81ce8a`. The 627-file packet seal matches committed bytes. Offline
validation and all five derived-table checks pass from the main workspace
both before and after removal of the original worktree and executables.

The owned build target (3,926,342,433 file bytes), copied root lock, three
reference symlinks, worktree and temporary branch were removed. Exact binary
path/size/SHA witnesses remain in `cleanup.json`. The candidate owner also
confirmed removal of its `/tmp/litchi-0779-patch-check.*` checkout and temporary
patch/log files. All pre-existing worktrees and the three unrelated main-tree
files retain their original state. Raw patch context whitespace is preserved
as evidence; other changed files pass the whitespace check. The sealed packet
is unchanged by this integration note.
