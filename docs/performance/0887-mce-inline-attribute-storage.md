# 0887 MCE inline attribute storage

The inline-four candidate is **rejected**. It reduces allocation calls but
does not improve the primary full-scan latency under the frozen gate, and all
four measured scenarios have large process-RSS review flags. Production is
restored exactly to its pre-candidate state. The patch and complete experiment
remain available for investigation; this batch adopts no optimization.

## Hypothesis and implementation

[0886](0886-producer-full-cell-profile.md) identified MCE processing as the
first area to investigate in the generated producer XLSX full scan. This
candidate removes the temporary attribute-vector allocation for start tags
with up to four attributes. It changes private storage in the shared OOXML MCE
owner; it does not cache parsed elements or skip validation.

An initialized `[Attr; 4]` holds the small case. A larger tag reserves a `Vec`
fallibly before moving attributes into it. Subsequent growth remains fallible.
Slice views preserve source order for namespace collection, duplicate checks,
directives, filtering, and emission. The adapter uses safe Rust and introduces
no dependency or public API change. The retained
[candidate patch](results/change-0887/candidate.patch) includes tests at the
inline/spill boundary, namespace-qualified attributes, normalization, and
expanded-duplicate refusal precedence.

The fifth attribute initially reserves five slots; a sixth can grow that
allocation again. This differs from the original vector's capacity sequence.
The measured scenario is therefore not evidence that every attribute-count
shape improves.

## Matched source and protocol

The before binaries are the exact retained native and allocator release
executables from [0885](0885-producer-full-cell-release-baseline.md). The after
build reconstructs that source archive and overlays only the two files in the
candidate patch. All 10,765 source-file hashes are checked before and after the
build. Compiler and release settings match: Rust 1.95.0, optimization level 3,
debug level 1, fat LTO, one codegen unit, incremental compilation disabled,
two build jobs, and locked offline dependencies.

The [capture helper](results/change-0887/experiment.py) launches one fresh
process for each case and variant, sequentially on CPU 12. Eight native blocks
cover four cases, with three warmups and fifty measured samples per child.
Before/after order alternates by block and case, giving four pairs in each
order per case. Four separate allocator pairs follow for the primary case,
with three warmups and fifteen samples per child. This is 72 children, 3,320
measured iterations, and 216 warmups. Every row is retained; the driver stops
on a failure without retrying or replacing a child.

| Case | Timed scope |
| --- | --- |
| Producer XLSX full-cell scan, primary | Fresh workbook open, enumeration, production freshness checks, and callbacks/digest updates across 4,096 cells |
| Ordinary XLSX medium full-cell scan | Cell iteration; workbook open and worksheet selection precede the timer |
| Producer DOCX selected paragraph | Selected paragraph text access; document open precedes the timer |
| Producer PPTX selected slide | Selected slide text/name access; presentation open and slide selection precede the timer |

The primary retains its ordered-value/count oracle, archive identity,
source-version fence, producer refusal census, and route checks from 0885.
Source/oracle preparation, harness post-timer checks, digest finalization, and
destruction are outside its timer. Controls retain exact corpus identities;
DOCX/PPTX check selected text on every iteration. The controls do not all time
MCE parsing and cannot establish a cross-format MCE speedup.

The [reducer](results/change-0887/reduce.py) forms after/before ratios within
each whole-child pair for p50, p95, p99, and process peak RSS. It reports the
median ratio and a percentile bootstrap interval using 10,000 resamples of
whole pairs, seed 887, and zero-based endpoints 249 and 9,749. It never pools
samples from different children. These intervals describe eight paired
processes, not eight independent hosts or workloads. With fifty samples per
native child, its nearest-rank p99 is the observed maximum; this is limited
tail evidence.

The primary adoption threshold is a p50 median ratio at most 0.98 and a
bootstrap upper endpoint below 1.0. Any control pair above 1.05 for latency or
RSS, any primary tail/RSS pair above 1.05, and any native p50 max/min spread
above 1.20 requires review. A passing reducer status means valid evidence,
not automatic adoption.

External GNU `time %M` measures whole-process maximum RSS in KiB for every
child. It includes source generation, setup, and reporting. Allocator results
are a separate primary-only operation-scoped population; region peak live
bytes includes bytes already live at region entry. Neither measure is a
substitute for the other.

CPU affinity does not provide an exclusive host. Other teams use this machine.
This is a warm-cache generated-corpus experiment, with no cold-cache, native
Office, remote-source, scaling, or universal producer claim.

## Correctness and ADR constraints

| Constraint | Evidence |
| --- | --- |
| ADR 0001/0002/0024 owner and API boundaries | Private storage in `litchi-ooxml-common`; no public API or dependency changes; crate-boundary check passes |
| ADR 0005 bounded fallible allocation and evidence | Four fixed slots; fallible overflow; matched retained binaries and separate native/allocator processes |
| ADR 0006 preservation and refusal behavior | Existing validation/emission paths retained; source-order, normalization, duplicate and directive precedence tests |
| ADR 0008 verified migration state | Shared-owner and format tests, reviewed source patch, and retained gate receipts |
| ADR 0010/0011 physical package ownership | Container/package code unchanged; primary archive and semantic identities checked |

The final candidate passes 445 shared OOXML owner unit/integration tests and
3,462 DOCX/PPTX/XLSX unit tests: **3,907 passed**. One existing PPTX diagnostic
`print_the_golden_transcript` is intentionally ignored. Formatting,
warning-denied owner Clippy, warning-denied owner rustdoc, and the crate-boundary
checker pass. The first owner compile attempt failed on three vector-style
iterator calls left after the storage change; those calls were fixed before
the passing gates and release snapshot. The failed log remains retained.

[Semantic review](results/change-0887/semantic-review.json) and
[methodology review](results/change-0887/methodology-review.json) are separate
from runtime evidence. The semantic review's pending gate fields reflect its
earlier creation time; completed gate receipts are under `quality-v2/`.
Standalone `xml_attribute_bounds` controls were suggested but were not run;
this report does not count them as passing checks.

## Results, decision, replay, and cleanup

All 72 fresh processes, 3,320 measured iterations, and 216 warmups completed
without retries or replacement. Every retained semantic check passes. The
[full reduction](results/change-0887/reduction.json) preserves every pair,
interval, and flag.

| Case | Before child p50 median (µs) | After child p50 median (µs) | Paired p50 ratio [95% interval] |
| --- | ---: | ---: | --- |
| Producer XLSX full scan | 3600.019 | 3648.776 | 1.0025 [0.9966, 1.0316] |
| Ordinary XLSX full scan | 131.998 | 132.318 | 1.0022 [0.9972, 1.0075] |
| Producer DOCX paragraph | 105.787 | 107.335 | 1.0106 [0.9828, 1.0291] |
| Producer PPTX slide | 177.901 | 178.857 | 1.0098 [0.9979, 1.0141] |

Ratios are medians of the eight within-pair after/before ratios. They are not
the ratio of the two displayed child medians. Values below one favor the
candidate. The primary paired p50 change is **+0.25%**, with interval
**−0.34% to +3.16%**. It fails both the practical 2% improvement threshold and
the requirement that the upper interval endpoint be below one.

| Case | Paired p95 ratio [95% interval] | Paired p99 ratio [95% interval] | Paired peak-RSS ratio [95% interval] |
| --- | --- | --- | --- |
| Producer XLSX full scan | 0.9977 [0.9939, 1.0208] | 0.9926 [0.9626, 1.0110] | 1.4021 [1.3751, 1.4477] |
| Ordinary XLSX full scan | 1.0035 [0.9868, 1.0096] | 0.9588 [0.8474, 1.1039] | 1.4353 [1.4211, 1.4681] |
| Producer DOCX paragraph | 1.0145 [0.9895, 1.0366] | 1.0170 [0.9908, 1.0666] | 1.4403 [1.3854, 1.4522] |
| Producer PPTX slide | 0.9989 [0.9849, 1.0167] | 0.9514 [0.9037, 1.0215] | 1.5363 [1.4528, 1.5892] |

Every RSS pair exceeds the 1.05 review threshold in all four scenarios.
Process peak-RSS child medians are 21,626 → 30,272 KiB for primary XLSX,
20,366 → 29,240 KiB for ordinary XLSX, 17,360 → 24,948 KiB for DOCX, and
16,836 → 25,758 KiB for PPTX. These are observed whole-process maxima. Their
cause is unresolved; operation-scoped allocation counters do not explain or
negate them. The candidate is rejected without claiming a causal heap or
stack explanation.

Other review flags retain two ordinary XLSX p99 pairs, one DOCX p95 pair,
three DOCX p99 pairs, and one PPTX p99 pair above 1.05. No native p50 spread
flag fires. Median tail ratios alone do not erase individual flagged pairs.

The separate four-pair allocator population has identical per-child medians
within each variant:

| Primary operation metric | Before | After |
| --- | ---: | ---: |
| Allocation calls | 25,933 | 21,678 |
| Reallocation calls | 366 | 356 |
| Allocated bytes | 37,618,900 | 36,799,636 |
| Region peak live bytes | 9,252,247 | 9,252,243 |
| Failed allocation calls | 0 | 0 |

Allocation calls decrease 16.41% and allocated bytes decrease 2.18%. The
four-byte region-peak difference is negligible. These counters support the
intended mechanism but do not establish an end-to-end latency improvement.

The [packet verifier](results/change-0887/verify_packet.py) replays command,
artifact, binary, corpus, and semantic checks; recomputes the full reduction;
and reconstructs all 10,765 source files from the baseline archive plus the
patch in a temporary directory. Source replay checks bytes and SHA-256;
`git apply` can rewrite touched-file permission bits under the local umask.
It does not rebuild or rerun the workload.

```sh
python3 -B docs/performance/results/change-0887/verify_packet.py \
  --plan docs/performance/results/change-0887/plan.json \
  --experiment-root docs/performance/results/change-0887/capture \
  --reduction docs/performance/results/change-0887/reduction.json \
  --output /tmp/litchi-0887-replay-new.json
```

Replay requires the locally retained 0885 source archive and before/after
binaries; these large artifacts are ignored by Git and bound by hashes.
The verifier uses retained-only plan verification so the rejected patch can
be replayed after production restoration and build-directory cleanup.

[Independent capture review](results/change-0887/capture-review.json) reproduces
all paired statistics, flags, raw allocator counters, and custody checks. The
[decision](results/change-0887/decision.json) records rejection and baseline
source hashes after reversing the candidate patch.

[Cleanup](results/change-0887/cleanup.json) removed only the marked
`/tmp/litchi-0887-source` and `target/litchi-0887-build` directories after
checking all 3,846 native and 3,847 allocator source dependencies against the
manifest. No Git worktrees were created for this experiment. Existing
historical recovery material and other teams' worktrees remain untouched.
The first cleanup attempt stopped before deletion because an import resolved
to the baseline helper; invoking the exact candidate verifier CLI fixed it.
That failed attempt is recorded, and the corrected cleanup passes. Three
byte-identical duplicate JSON exports were also removed while retaining their
canonical files.

[Post-cleanup replay](results/change-0887/post-cleanup-verification.json)
passes, including source reconstruction and the full statistical replay.
The broader performance goal remains active.

Next, qualify representative MCE parsing controls and investigate other
measured MCE work, such as repeated name expansion, before another candidate.
The allocation reduction alone does not justify keeping this storage adapter.
