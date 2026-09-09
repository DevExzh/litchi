# 0490 results review

The controlled rerun does not reproduce a consistent file-store tail
regression, but it also does not establish a reliable tail improvement. Normal
file-store p50 improves by a mean 6.73% across six paired blocks, with a
95% block-bootstrap interval of -9.71% to -4.41%. Normal p95/p99 intervals
span both directions. Allocator file-store p50 also has an interval spanning
zero. Individual adverse blocks remain visible. No production change follows
from this measurement, and the historical 0489 regressions remain recorded.

## Formal evidence and uncertainty

The [complete tables](comparison-tables.md) and [machine-readable summary](variance-summary.json)
retain all 72 processes and 4,320 samples. The exact tiny workload, two retained
binary versions, normal/allocator roles, three routes, six alternating blocks,
60 samples and five warmups per child were frozen before capture. Receipt
chronology verifies the actual declared order with no overlapping children.
All source, authored, candidate archive/XML and semantic/preservation identities
match. Data synchronization remains mandatory for the file-store route.

Mean paired-block latency changes are:

| Binary and route | p50 | p95 | p99 |
| --- | ---: | ---: | ---: |
| Normal deterministic | -18.90% | -20.69% | -20.82% |
| Normal memory store | -21.37% | -19.92% | -17.32% |
| Normal file store | -6.73% | -12.16% | -8.91% |
| Allocator deterministic | -16.11% | -18.93% | -22.04% |
| Allocator memory store | -21.86% | -19.41% | -17.04% |
| Allocator file store | -2.53% | -12.98% | -10.86% |

The file-store tail means are not robust gains. Normal p95's interval is
[-35.44%, +13.33%] and p99's is [-33.75%, +15.40%]. Allocator p95 is
[-25.75%, +1.10%] and p99 is [-28.31%, +10.12%]. The bootstrap resamples
six paired process blocks, never individual in-process samples. Six blocks on
a shared host do not establish independence from host/filesystem state or a
strong general confidence guarantee. The means and all intervals are
scoped to the one recorded workload, machine and executable pair.

## Adverse rows and memory

Normal file-store block 4 p99 increases 38.03%; block 6 p95/p99 increase
42.84/11.47%. Allocator file-store block 2 p50/p95/p99 increase
5.01/14.89/36.47%. These are all six elapsed-time quantiles above +5%.
They demonstrate why an overall mean cannot erase a tail concern.

Operation peak heap improves in every block: deterministic 989,190 →
792,403 bytes (-19.89%), memory store 9,184,134 → 8,987,347 bytes (-2.14%),
and file store 795,721 → 598,933 bytes (-24.73%). The file-store path differs
from 0489's capture pathname, explaining the small absolute path-storage
variation; each new paired block uses the same declared naming policy.
No operation-heap quantile rises. No GNU-time maximum RSS pair exceeds +5%.

Seven additional adverse procfs quantiles are retained in the complete table.
Six are memory-store allocator `process_rss_bytes` quantiles. That field comes
from `process_metrics::Delta`: it is the saturating difference between sampled
RSS values, not total RSS or a peak-memory measurement. Its tiny interpolated
baseline can produce percentages such as +24,100% for 204.8 → 49,561.6 bytes.
The absolute values remain visible. The seventh is file-store allocator block
2 median after-sample VmHWM, 6,258,688 → 6,619,136 bytes (+5.76%). VmHWM is
not a delta, and those correlated per-sample observations are distinct from the
single whole-child GNU-time maximum. None of these fields is a bounded-memory
proof or a reason to dismiss the earlier 0489 RSS observations.

## Synchronization attribution and Amdahl limit

The [separate sync summary](sync-summary.json), [plot](sync-attribution.png)
and [review notes](review-notes.md) retain eight diagnostic children. All four
sync-only children have exactly 34 successful fdatasync events: one preflight,
three warmups and 30 measured operations. Path, ordering, count and duration
containment checks pass. The median fdatasync fractions are 79.52%, 80.73%,
81.61% and 78.99%, with median sync duration about 3 ms. One after repeat-1
sync reaches 8.04 ms inside an approximately 8.8 ms operation.

For these traced samples, if that sync cost remains fixed, removing all other
work would give an idealized Amdahl bound of about 1.23–1.27×. This is a scoped
attribution model, not a hardware limit, untraced speedup promise or license to
weaken durability. Tracing perturbs application timing. The data supports sync
latency as a contributor to observed variation; it cannot retrospectively prove
the cause of the original 0489 outlier.

The four syscall summaries have identical selected call counts between
versions: three fdatasync, 94 statx, 24 pread64 and 3,205 writes per whole
child. Their scope includes fixture setup, preflight, warmup, timed operation,
cleanup and JSON report output. Strace `-c` reports system time unless `-w` is
selected; these summaries do not measure blocked wall time. Their write count
must not be presented as timed replay writes. Per-sample benchmark records
retain one replay write and one sync. The file-store provider and data-sync
policy are unchanged, and no production optimization is justified here.

## Validation, cleanup and next work

Twenty focused Python tests pass. Independent reviews cover the timing/source
contract, exact sync alignment and summary parser; root review strengthened
formal terminal/gate checks, timeouts and actual capture chronology before the
final freeze. Unused pre-review helper/protocol versions remain explicit
historical artifacts. Both formal and diagnostic verifiers recompute the
retained results. The evidence seal checks the complete inventory after cleanup.

There is no production Rust change or rebuild. Current sources match the
0489 manifest, whose existing Rust, fuzz and preservation gates remain the
bound executable evidence. All 24 formal file-store replay directories and
eight diagnostic replay directories are removed after proving them empty.
The batch's remaining empty scratch parents and generated bytecode are removed
before sealing. The protected spec-gap worktree and unrelated local files are
preserved.

The [next implementation contract](next-implementation.md) prioritizes a
verified-cold and provider matrix for the existing end-to-end DOCX full-text
selector. Prepared warm queries remain ineligible for cold claims. Genuine
borrowed-source support is still an API gap; bounded concurrency and distinct
native producer/Office evidence remain separate requirements. The broader
non-iWork goal is active and not complete.
