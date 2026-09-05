# 0422 operation-region resource review

This batch is a current V3 allocator diagnostic for source revision
`40c40ea89591140e15e68d00f3fce2c4c376e445`. It is not a before/after
comparison and does not authorize an optimization, latency, or RSS claim.
The four reports use the same `serialized_region_peak_v3` marker; each
selector's two repeats have matching corpus and output identities. Each report
has 30 samples with three warmups in a fresh process.

The stable descriptive result is:

| Selector | Region peak mean | Lifetime peak-after mean | Live-after mean |
|---|---:|---:|---:|
| `pptx_cross_copy_media_rich_lifecycle` | 272,736,303 B | 812,687,524 B | 237,453,206 B |
| `pptx_cross_copy_plain_lifecycle` | 1,360,003 B | 3,559,492 B | 797,423 B |

The R1 and R2 means agree for these resource fields. In the individual
reports, the media region values range from 272,735,375 to 272,737,231 B and
the plain values from 1,359,075 to 1,360,931 B. Every report has 30 aligned
`sample_indices` and `elapsed_ns.sample_order` entries. For every sample,
`region_peak_live_bytes` is at least both endpoint live totals and no greater
than `peak_live_bytes_after`; the summary's lifetime high-water and boundary
invariants also pass.

The region number is an absolute process live-byte maximum observed in
allocator-callback order between `begin` and `finish`. It starts with the
live total already present at entry and includes callbacks from any process
thread that linearize inside that interval. The mutex makes callback, boundary,
and snapshot order well-defined; it does not turn the value into an
operation-owned allocation total. The callback is observed after the allocator
call returns, so allocator-internal realloc copy overlap and physical heap
state are outside this measure. The observer also excludes RSS.

The lifetime high-water number is a process-wide maximum reached at any time
before the finish snapshot. It can therefore remain far above the region
value because setup or an earlier phase reached that maximum. Likewise,
`live_bytes_after` is the process-wide endpoint at finish, not a post-drop
ownership snapshot. The difference between the media endpoint (237,453,206 B)
and lifetime high water (812,687,524 B), or between the plain endpoint
(797,423 B) and its lifetime high water (3,559,492 B), must not be labeled
retained memory or memory released by this lifecycle. The signed endpoint
change is also only a process net change. Whole-process `time -v` RSS in the
table (803,448 KiB for media and 82,732/82,736 KiB for plain) includes setup
and warmups and is not a reduction result.

These reports are V3 evidence and must not be paired with the older V2
process-counter reports or used to claim a historical delta. They establish a
repeatable current baseline for the observer contract only; elapsed comparisons
were intentionally excluded. Raw reports retain elapsed samples for custody,
while the derived resource table omits them; no allocator timing or scaling
conclusion is available.

The next resource study should add explicit snapshots after planning, commit,
publication, and each relevant object drop, while keeping the measured
boundary stable. If destruction is moved inside the interval, it should use a
separate diagnostic selector. Near-limit and one-short resource-limit cases
should record typed refusals and bind snapshots to concrete owners and release
points. A matched source-backed media lifecycle is also needed before making
retention or budget statements. Region peaks and RSS alone cannot establish
those claims.

See [the derived result table](result-table.md), [the observer design](design.md),
and [the validation record](validation.md) for the exact vectors, invariants,
and scope limits.
