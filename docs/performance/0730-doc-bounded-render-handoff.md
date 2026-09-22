# 0730 — bounded DOC validated-render handoff

Retain the bounded handoff candidate. NoHeadFoot's ordinary public lifecycle
improves 11.89–13.88% in all six paired median comparisons, exceeding the
prospective 5% gate in every cycle. FloatingPictures is mixed: paired medians
range from 14.86% faster to 1.27% slower. Both fixtures preserve exact established
outputs. This adoption accepts the explicitly measured two-edit memory tradeoff
below; it does not claim a general DOC speedup.

| Fixture | Baseline process p50, μs | Candidate process p50, μs | Paired p50 delta | Paired mean delta |
| --- | ---: | ---: | ---: | ---: |
| FloatingPictures.doc | 983.75–1064.47 | 902.31–998.25 | −14.86% to +1.27% | −14.10% to +0.85% |
| NoHeadFoot.doc | 103.41–107.34 | 90.39–92.44 | −13.88% to −11.89% | −14.62% to −11.64% |

Positive deltas are slower. Process ranges span twelve baseline and six candidate
processes per fixture, including A/A controls; paired changes use the six adjacent
A/B and reversed B/A comparisons. These are descriptive ranges, not confidence
intervals or pooled-sample estimates.

| Fixture / cycle | A/A p50 delta | A/B p50 delta | B/A candidate-versus-baseline p50 delta |
| --- | ---: | ---: | ---: |
| FloatingPictures / 0 | −0.14% | −8.81% | +1.27% |
| FloatingPictures / 1 | −3.23% | −5.38% | −14.86% |
| FloatingPictures / 2 | −1.48% | −7.97% | −11.12% |
| NoHeadFoot / 0 | −0.22% | −13.88% | −11.89% |
| NoHeadFoot / 1 | −0.84% | −12.10% | −13.21% |
| NoHeadFoot / 2 | +0.32% | −13.57% | −13.57% |

All twelve A/A p50/mean checks stay within 5%, and no candidate p50/mean
regresses above 5%. Tail variability remains visible: A/A absolute changes
exceed 5% in FloatingPictures p99/maximum once each, and in NoHeadFoot p95 twice
and p99/maximum three times each. These ten tail-statistic flags all point toward
faster second baseline runs and are retained, not treated as evidence of improved
candidate tails. Every process p50/mean/p95/p99/maximum is in `analysis.json`.

| Fixture | Allocated bytes, baseline → candidate | Allocation calls | Single-edit peak live bytes |
| --- | ---: | ---: | ---: |
| FloatingPictures | 15,512,580 → 14,115,027 (−9.01%) | 15,756 → 15,141 (−3.90%) | 3,128,190 → 3,128,190 |
| NoHeadFoot | 1,061,429 → 936,223 (−11.80%) | 1,907 → 1,658 (−13.06%) | 209,515 → 209,515 |

Both allocator repetitions agree exactly. End-of-region retained output stays
342,528 and 27,648 bytes respectively. The intermediate held render capacities
are larger than output length: 587,776 and 36,864 bytes. The latter are the
values charged to the 8 MiB ceiling.

| Two-edit candidate control | Zero-retention peak | Default-retention peak | Increase |
| --- | ---: | ---: | ---: |
| FloatingPictures | 3,182,172 B | 3,714,077 B | 531,905 B / 16.72% |
| NoHeadFoot | 236,249 B | 273,113 B | 36,864 B / 15.60% |

Both peak increases cross the prospective review threshold. We accept them for
this bounded handoff: one explicitly observable allocation is retained, its
capacity has a finite caller-controlled ceiling, and zero retention or explicit
release restores the recomputation allocation profile. Direct tracked-revision
editors still default to zero retention. Two-edit allocated bytes fall 7.49% and
9.41%, respectively, while end-of-region retained bytes remain equal. Releasing
after each stage matches the zero-retention allocation fields exactly on both
routes. No total process-memory or RSS bound follows from this tradeoff.

The candidate removes one duplicate container render from a supported public
DOC body transaction. The common editor returns the render it already validated;
the DOC editor can retain that allocation and move it into finish. The final
strict-owner and public-reader checks remain mandatory. Reuse remains the save
policy, and exact no-ops retain their existing behavior.

Retention belongs to the private editor. A clone drops the token, successful
mutation replaces it, resource reopen clears it, and explicit release permits
later fresh retention. The ceiling measures `Vec::capacity()`: the
default is 8 MiB, while the existing three-argument `TransactionLimits::new`
keeps retention disabled. Over-budget output takes the recomputation path.
Policies meet by componentwise minimum in patches, composition, three-way
plans, and the receiving side of transfers. Read-only donor limits do not
control receiver retention.

This is a retained-state ceiling, not a process-memory ceiling. During a second
clone-first edit, the old token can overlap transiently with the newly rendered
candidate. A separate allocation lane compares zero retention, default retention,
and release after each stage for one and two edits on both fixtures. Its zero
control is the candidate with retention disabled, not the baseline binary.

The ordinary comparison uses the unchanged 0728 default-feature public probe,
with exact output, semantic, unknown-stream and raw-directory preservation gates.
CPU 12 on AMD EPYC 9R45 and Rust 1.95.0 are fixed. Each of three cycles runs
A/A followed by A/B/B/A for both fixtures, reversing case order in the middle
cycle: 36 native processes, 1,800 measured lifecycles and 108 warmups. Eight
separate allocator processes use A/B/B/A. All samples and tails remain in the
packet. Allocation timing is not used as latency evidence.

The prospective small-fixture gate requires more than 5% median improvement
with consistent direction across cycles. A/A deviations and candidate p50/mean
regressions above 5%, plus allocation/peak increases above 5%, require explicit
review. No selective rerun or outlier removal is permitted.

## Correctness and architectural constraints

| Constraint | Candidate evidence |
| --- | --- |
| ADR 0003: isolated edits, atomic publication and reversible patches | Clone-cleared ownership; failed picture graph installation preserves prior bytes/token; patch and conflict tests. |
| ADR 0005: finite, observable and releasable retained state | Capacity ceiling, zero/under/exact/above controls, explicit release, policy intersections and allocation ownership test. |
| ADR 0006: lossless preservation and independent validation | Exact established outputs, semantic/raw-directory oracle, missing-Data and resource-reopen tests, forced final public-reader failure. |
| Container ownership and topology | Common editor returns its own validated render; no common cache or dependency reversal. |

Final root qualification passes all 13 commands: formatting; all-feature,
all-target DOC/common tests (1,367 passed, two ignored); warning-denied Clippy;
14 doctests (12 ignored); warning-denied rustdoc; both probe checks including
four main-probe unit tests; and the repository crate-boundary scan. Independent
replay agrees on all 44 processes, and 20 actual-analyzer corruption controls
pass. All 24 retention processes pass direct output equality and capacity checks.

All 34 inherited constraint hashes remain checked. The packet retains failed
build and qualification attempts, source snapshots, and the documented early
agent scheduling deviation. These are not substituted for the final root-run
qualification.

This two-fixture experiment cannot establish broad producer compatibility,
peak RSS, cold-cache behavior, remote I/O, concurrent scaling, or a universal
speedup. It does not promote broader CRUD coverage. iWork remains excluded and
the non-iWork performance goal remains active.

[Evidence and reproduction](results/change-0730/README.md).

The next separate investigation is public PPT save attribution under the
existing Reuse-preservation contract. This DOC result does not identify PPT's
internal bottleneck or authorize removal of its validation. Broader DOC corpus
and larger/many-edit memory measurements remain useful follow-up evidence.
