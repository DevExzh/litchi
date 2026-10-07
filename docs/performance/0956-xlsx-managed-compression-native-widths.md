# 0956 — native cold managed XLSX compression widths

The existing explicit managed compression route reduces medium one-percent
publication latency at widths two and four. Paired publication p50 ratios are
0.663 and 0.608 relative to width one; the corresponding timed-route ratios are
0.888 and 0.867. This is a measurement of existing behavior, not a production
optimization or adoption recommendation. Four workers increase whole-process
peak RSS by 6.06%, crossing the 5% review threshold. Both parallel widths also
increase the separately measured commit-drop phase substantially. Allocation
qualification and those regressions remain unresolved.

The [bounded packet](results/change-0956/README.md) contains the frozen plan,
source/build identities, qualifications, complete capture inventory, reduction
summary and reviews. Large raw evidence and binaries remain external.

## Experiment and boundaries

One Rust 1.95.0 release executable serves all widths: opt-level 3, fat LTO,
one codegen unit, debug level 1, incremental disabled, no allocator or
frame-pointer instrumentation. Its SHA-256 is
`38381c594b6dd5924e07d3a27f57bf15729811b8f72d7353aac5aac810741304`.
The source is Git base `c8e80705989183d88950e282ef1852f3e3470b10` plus five
reviewed harness files; production Rust is unchanged. Unrelated checkout work
was excluded from the measurement source. Scoped format and warning-denied
Clippy checks passed for both examples.

Three cases compare widths one versus two and one versus four, each with six
alternating ABBA/BAAB blocks. All 144 fresh processes completed once, without
retries or replacement. Each performs 20 warmups and 200 measured iterations:
28,800 measured iterations in total. The CPU mask is 12–15, four distinct cores
in the exposed topology. Affinity is witnessed on the live `/usr/bin/time`
leader and inherited by the workload; individual thread affinity and OS thread
counts are not sampled. Owned builds and Clippy finished before capture.
The host exposes an AMD EPYC 9R45 under KVM, 32 CPUs and 129,447,072 KiB
of memory, running Linux 7.0.0-1013-aws. Full details are in
[host.json](results/change-0956/host.json). This does not establish exclusive
use of the host.

Every iteration creates a fresh finite managed context, package and bounded
sink. The process is warmed; the managed session/private pool is cold. There
is no cold-filesystem claim. Context creation, immutable source cloning, cache
limit preparation and sink reservation are outside the timers. Ingress starts
at the source-backed editor constructor. Edit/commit, publication, and commit
drop are measured separately. Publication includes the consuming editor call,
its package/private-pool teardown and returned snapshot drop; commit drop is
outside publication but inside total. Total therefore describes this timed
route, not the full process or application request. These boundaries preclude
an absolute timing comparison with 0950.

The fixed corpus, finite budgets, output ceilings and 64 KiB maximum sink write
match the 0955 qualification. Every measured output passes exact byte/hash,
semantic reopen, task/output accounting and zero retained live-resource checks
outside the clocks. Each process verifies its actual managed commit inverse in
an untimed preflight; each timed output is checked with a precomputed fixture
inverse. The latter is not a per-sample actual-commit inverse claim. Medium
admits worker permits 0/2/4 and five tasks; dense-sparse remains inline with
five tasks, and one-edit remains inline with two. Worker permits are admission
evidence, not a measurement of task overlap.

All nine small runtime reports and eight invalid-CLI controls passed. The
qualification wrapper then recorded an existing-directory setup failure; its
separate unique-child recovery verified refusal without modifying the sentinel.
No successful workload was retried. The original qualifier replay passed and
its manifest is byte-identical to 0955. Source-manifest reproduction also
matches the frozen manifest exactly.

## Paired results

Each process produces its own nearest-rank quantiles. Within each block, the
two adjacent treatment/serial ratios are combined geometrically; the reported
point is the geometric mean of six block values. Intervals use 10,000 bootstrap
resamples of those six blocks, seed 956. Samples are not pooled across processes.
Ratios below one indicate lower measured cost. The 1.05 thresholds are point
review triggers, not confidence-bound acceptance gates.

| Case | Width | Publication p50 / p95 / p99 | Total p50 / p95 / p99 | RSS ratio |
| --- | ---: | --- | --- | ---: |
| Medium one-percent | 2 | 0.663 / 0.671 / 0.678 | 0.888 / 0.889 / 0.886 | 1.024 |
| Medium one-percent | 4 | 0.608 / 0.640 / 0.636 | 0.867 / 0.877 / 0.874 | **1.061** |
| Dense-sparse one-percent | 2 | 0.998 / 1.027 / 1.005 | 1.002 / 1.011 / 1.004 | 1.000 |
| Dense-sparse one-percent | 4 | 0.998 / 1.002 / 0.997 | 0.992 / 0.994 / 0.993 | 1.000 |
| Medium one-edit | 2 | 1.003 / 1.007 / 1.022 | 1.000 / 1.002 / 1.008 | 0.996 |
| Medium one-edit | 4 | 1.004 / 1.001 / 1.002 | 1.001 / 1.000 / 1.000 | 0.999 |

For medium, publication p50 95% intervals are [0.629, 0.703] at width two
and [0.571, 0.650] at width four; total p50 intervals are [0.874, 0.906]
and [0.849, 0.885]. RSS intervals are [0.995, 1.052] and [1.026, 1.098].
RSS and CPU time, faults and context-switch counts cover the entire process,
including corpus construction and verification; they are not phase allocation
or CPU-counter measurements. All quantile intervals and secondary metrics are
retained in [the summary](results/change-0956/comparison-summary-v4.json).

The medium commit-drop p50 ratios are **1.738** and **1.851**, with p95/p99
ratios also above 1.70. Across the twelve processes in each role, serial p50
commit drop spans about 167–170 microseconds, versus 291–295 at width two and
306–320 at width four. This penalty is included in total and is not dismissed
because total improves. Its cause has not been attributed; pool teardown must
not be assigned to this phase because the consuming publisher already covers
it. Edit/commit point ratios are about 1.01 for medium. Ingress and smaller
phases retain individual pair/block flags in the summary.

All aggregate publication/total points remain below the 5% regression
trigger, but their tails are not uniformly clean. Dense-sparse publication p95
has individual pair flags at both widths; width two's interval extends to
1.056. Width-two one-edit publication p99 has a pair flag, and width-four
one-edit publication has pair flags at p50, p95 and p99. Medium RSS has pair
flags at both widths, in addition to the width-four aggregate regression.
The all-phase pair checks flag five of six groups at p50, all six at p95,
and all six at p99; they include small ingress/drop phases and are distinct
from aggregate publication/total results. These are retained diagnostic flags,
not erased by aggregate results. The
earlier 0950 one-edit p99 +5.65% review flag remains binding; this different
timing/context contract does not resolve it.

## Custody and disposition

All capture files were inventoried before reduction: 1,154 files, 23,241,616
bytes, manifest SHA-256
`f923ebe1bb5f8a3a58458f6458b0795ea26071ebd6ea4254c624d769b489ad28`.
The first reducer correctly stopped on its own schema assumption: it expected
the deadline in the terminal record, while frozen capture v3 records it only
in the start record. The preserved v4 adapter requires that start deadline and
rejects a contradictory terminal value. Pairing, quantiles and bootstrap are
unchanged. Raw capture, plan and helper v3 are unmodified; no native workload
was rerun. Both the failure and separately reviewed adapter are retained.

Development records also retain the rejected pre-build source revision with a
missing report field, the correctness wrapper's post-workload directory setup
failure, and the pre-capture custody hardening. None supplies a substituted
native measurement. Owned build/temp caches were removed while canonical
source, executables and evidence were retained.

An exploratory invocation of the older 0955 diagnostic executable with
`--help` created a manifest directory because that qualifier accepts a positional
output path. This additional unlogged correctness run is disclosed; its generated
manifest was relocated out of the checkout and matches 0955. It contributes no
0956 timing evidence.

The next step is allocation qualification and attribution of the medium
commit-drop/RSS costs, with width two the narrower measured candidate. There
is no change to default execution policy, threshold, worker count or production
API. The overall Office performance goal remains active; iWork is excluded.
