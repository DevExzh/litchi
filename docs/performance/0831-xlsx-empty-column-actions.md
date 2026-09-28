# 0831 — skip the unused XLSX column-action owner map

Adopt the three-line empty-action guard. On the pinned real XLSX edit it removes
one allocation and 524,288 requested bytes per operation: **2,898,254 → 2,373,966
bytes (18.09% less)**. The native paired p50 ratio is **0.980416**
(bootstrap interval **0.975937–0.983917**), about **1.96% lower**.
The frozen requested-byte benefit gate passes, and no latency, RSS or
allocation/region-peak regression flag is triggered.

The real open/edit/default-save ratio is 0.997905 (0.984528–1.001317), which
includes 1.0. No lifecycle speedup, real-operation peak-memory reduction or
process-RSS reduction is claimed. Twenty native spread flags and ten p99/p50
diagnostics remain visible below.

## Change and preservation boundary

The change adds an empty-action return immediately after the protected-sheet
check in `validate_column_actions`. Cell-only edits do not query this validator's
column-owner map, so its allocation and initialization are unnecessary. The
[0830 exact-owner profile](0830-xlsx-edit-allocation-profile.md) identified one
524,288-byte allocation per real-file edit at this site. The two parser maps
remain intact, including their bounded overlapping-range behavior.

The source scan, cell/row checks and defaults validation keep their order.
Nonempty column actions retain their protected-sheet, implicit-width, sparse
split and style-retargeting behavior. The empty path no longer attempts this
allocation and therefore cannot produce its particular allocation-failure
error; no other refusal is bypassed. Public APIs, immutable ownership,
publication, default save durability, package preservation and concurrency
are unchanged.

| ADR constraint | Evidence |
| --- | --- |
| 0001/0002/0024 API and owner boundaries | Three lines inside the existing XLSX validator; no dependencies or exports |
| 0003 snapshots, atomic commits and patches | Existing transaction and patch paths remain unchanged; full XLSX suite on both legs |
| 0005 bounded resources and performance evidence | Removes one unused fixed map; native and observer measurements separated |
| 0006 preservation and validation | Scan and nonempty action checks retained; pinned full-output byte oracle on both legs |
| 0010/0011 physical-package ownership | No archive, compression, package or save changes |

## Fixed experiment

[The packet](results/change-0831/README.md) retains the before and after source,
input freeze, machine/build identities, raw reports, command receipts, failed
attempts, offline readers and cleanup witness. The baseline revision is
`3bcaee6f418a78ee62bab76aed0d6d9f1f93d024`.

The host is an AMD EPYC 9R45 on Linux x86_64, with Rust 1.95.0 / LLVM
22.1.2 and an ext4 filesystem. Affinity exposes 32 CPUs; the captured cgroup
hierarchy has no finite CPU or memory quota. Release builds use optimization
level 3, thin LTO, one codegen unit, debug info level 1 and panic unwinding.
Full host and build receipts are retained.

Nine cases run serially on CPU 12. Six native blocks alternate AB/BA, with three
warmups; real edit, real lifecycle and no-op use 500 measured samples per
process, and six synthetic scale guards use 30. Two separate observer blocks
alternate AB/BA with three samples and no warmups. Qualification uses one
sample per source/binary/case. The planned total is 180 reports and 20,304
measured samples. Observer elapsed times do not support latency claims.

Within-process analysis uses nearest-rank p50/p95/p99; the harness's separately
serialized midpoint p50 is checked against its own definition. Across-process
values use midpoint medians. Paired ratios use 10,000 bootstrap draws, seed
831831, with sorted endpoints 250 and 9749. The benefit gate is at least 15%
fewer requested allocated bytes for real edit with identical output. Native
p50 lower bounds above 1.05, paired RSS ratios above 1.05, and allocation/region
peak increases require individual review.

The real fixture is the checked-in LibreOffice QA `dateAutofilter.xlsx`
(8,435 bytes, SHA-256 `d7ab3dbb59388d245ee779bf8547748dc6bac70f3c7216e673e0d97dbbbd6bc4`).
The reference output is 8,521 bytes, SHA-256
`0f6902152f1b8c40c47086206023887e87f54ef57f1da417bc91d8494a4d3e68`.
Real edit times the public edit operation and verifies admitted outcomes; it
does not serialize each timed edit. The separate pinned 0830 probe checks the
complete reference bytes on both source legs. Real lifecycle includes open,
edit and default atomic save and checks every published hash. Synthetic
commit/save excludes setup and edit planning, compares complete deterministic
expected bytes and reopens each result. No-op exits before the changed guard.
No nonempty column-action latency is measured.

## Matched results

Native entries are midpoint medians over six processes per leg; ratios and
intervals use paired block ratios. Time units are milliseconds. RSS is the
whole-process high-water value in KiB, including harness setup.

| Case | p50 before → after (ms) | Paired ratio [interval] | RSS before → after (KiB) | Paired RSS ratio |
| --- | ---: | ---: | ---: | ---: |
| real-edit | 0.253691 → 0.248711 | 0.980416 [0.975937, 0.983917] | 137,276 → 137,340 | 1.000277 |
| real-lifecycle | 5.408116 → 5.395071 | 0.997905 [0.984528, 1.001317] | 137,312 → 137,322 | 1.000073 |
| one-cell-tiny | 0.107190 → 0.101195 | 0.942858 [0.940786, 0.947448] | 137,342 → 137,264 | 0.999694 |
| one-cell-medium | 0.888474 → 0.882790 | 0.993495 [0.990319, 0.995486] | 137,254 → 137,330 | 1.000510 |
| one-cell-dense-wide | 62.104056 → 61.972437 | 0.998447 [0.985964, 1.000063] | 137,308 → 137,270 | 0.999505 |
| one-percent-tiny | 0.181211 → 0.169736 | 0.936734 [0.930982, 0.942064] | 137,330 → 137,332 | 0.999956 |
| one-percent-medium | 3.720112 → 3.705948 | 0.996076 [0.991668, 0.999868] | 137,274 → 137,294 | 1.000204 |
| one-percent-dense-wide | 125.612858 → 125.528730 | 0.999925 [0.997917, 1.004960] | 137,308 → 137,262 | 0.999563 |
| noop-medium | 0.000570 → 0.000570 | 1.000000 [0.991228, 1.044173] | 137,258 → 137,292 | 1.000219 |

Tail and mean values are descriptive. Synthetic scale guards have only 30
samples per process, making their p99 the observed maximum.

| Case | p95 before → after (ms) | p99 before → after (ms) | Mean before → after (ms) |
| --- | ---: | ---: | ---: |
| real-edit | 0.417322 → 0.419517 | 0.519953 → 0.518018 | 0.281399 → 0.278295 |
| real-lifecycle | 5.693427 → 5.649633 | 5.943518 → 5.776288 | 5.450358 → 5.415310 |
| one-cell-tiny | 0.121631 → 0.118576 | 0.131406 → 0.123171 | 0.109478 → 0.103267 |
| one-cell-medium | 0.896564 → 0.889845 | 0.905140 → 0.893339 | 0.888947 → 0.882932 |
| one-cell-dense-wide | 62.722395 → 62.243595 | 62.905135 → 62.726482 | 62.203888 → 62.000151 |
| one-percent-tiny | 0.201226 → 0.187771 | 0.209911 → 0.191586 | 0.184443 → 0.172872 |
| one-percent-medium | 3.731088 → 3.720992 | 3.758018 → 3.732547 | 3.722655 → 3.707172 |
| one-percent-dense-wide | 126.721664 → 126.236640 | 127.212099 → 126.368674 | 125.741269 → 125.603657 |
| noop-medium | 0.000730 → 0.000700 | 0.000780 → 0.000770 | 0.000590 → 0.000591 |

Observer values use the median of three samples per process, then the midpoint
median of two processes. Peak above entry subtracts entry live bytes from each
region peak before taking medians; it is independent of process RSS.

| Case | Allocation calls before → after | Requested bytes before → after | Peak above entry before → after (bytes) | Net live bytes (both legs) |
| --- | ---: | ---: | ---: | ---: |
| real-edit | 2,510 → 2,509 | 2,898,254 → 2,373,966 | 1,066,957 → 1,066,957 | 4,642 |
| real-lifecycle | 3,542 → 3,541 | 4,561,313 → 4,037,025 | 1,103,104 → 1,103,104 | 40,789 |
| one-cell-tiny | 1,137 → 1,136 | 1,144,556 → 620,268 | 545,845 → 518,759 | 29,285 |
| one-cell-medium | 5,008 → 5,007 | 2,052,324 → 1,528,036 | 941,014 → 941,014 | 421,715 |
| one-cell-dense-wide | 142,183 → 142,182 | 42,994,886 → 42,470,598 | 21,210,673 → 21,210,673 | 14,348,553 |
| one-percent-tiny | 1,775 → 1,773 | 2,169,665 → 1,121,089 | 573,530 → 552,204 | 56,139 |
| one-percent-medium | 18,893 → 18,889 | 7,978,549 → 5,881,397 | 2,340,425 → 2,340,425 | 1,706,037 |
| one-percent-dense-wide | 303,080 → 303,078 | 89,569,992 → 88,521,416 | 36,443,008 → 36,443,008 | 28,951,315 |
| noop-medium | 0 → 0 | 0 → 0 | 0 → 0 | 0 |

The real edit and lifecycle absolute region peaks are three bytes lower after,
but their entry-live baselines are also three bytes lower. Their peak above
entry is unchanged. The tiny synthetic cases reduce peak above entry by
27,086 and 21,326 bytes; other measured operation peaks remain unchanged.
Absolute region and process-lifetime allocator peaks remain in `analysis.json`;
`decision.json` retains the baseline-adjusted derivation. Lifetime peaks can
include corpus construction outside the timed operation. No observer sample
reports a failed allocation.

Every one-cell case removes one 512 KiB map. The one-percent cases remove two,
four and two maps at tiny, medium and dense-wide scale respectively. The no-op
control allocates zero bytes inside its measured region on both legs.

## Regression and uncertainty review

No frozen guard is triggered: all native paired p50 lower bounds are at most
1.05, all paired RSS estimates are at most 1.05, and neither observer block
increases allocation calls, requested bytes or absolute region peak in any
case. Net live bytes are unchanged. The following descriptive flags are
retained rather than interpreted as general tail improvements.

| Case | Within-leg spread above 5% | p99/p50 above 1.05 |
| --- | --- | --- |
| real-edit | before.p95, before.p99, before.mean, after.p95, after.mean | before, after |
| real-lifecycle | before.p99, after.p99 | before, after |
| one-cell-tiny | before.p95, before.p99, after.p95, after.p99 | before, after |
| one-cell-medium | none | none |
| one-cell-dense-wide | none | none |
| one-percent-tiny | before.p99, after.p99 | before, after |
| one-percent-medium | before.p99 | none |
| one-percent-dense-wide | none | none |
| noop-medium | before.p99, before.mean, after.p50, after.p95, after.p99, after.mean | before, after |

The after no-op p50 spread flag is visible even though its paired median ratio
is 1.0 and its interval does not trigger the frozen regression rule. The
lifecycle and dense-wide intervals include 1.0; their allocation reductions
support no additional latency claim. No geometric mean or cross-format
aggregate is used for these differently scoped workloads.

## Validation and retained exceptions

Both legs require fresh formatting, all-feature/all-target checking, full XLSX
tests, warning-denied Clippy and rustdoc, crate boundaries, pinned byte-oracle
tests and harness library tests. Root and standalone harness lockfiles differ
in existing dependency versions; each is fixed across legs. Root-package
quality uses the root lock, while measurements and harness tests use the
standalone lock. Cross-lock equality is not claimed.

The initial preparation attempt failed an incorrect lock-equality assertion
before any build or measurement; its source and traceback are retained.

The frozen driver's observer identity assertion also required an additive repair:
its selected `allocator-metrics,ordinary-save-process-metrics` features emit
`ordinary_save_procfs_and_system_allocator_operation_scoped`, while the driver
expects the allocator-only string. The unchanged original driver and failure
log are retained. `recovery.py` binds its own source and all preexisting
artifacts before execution, revalidates complete baseline qualifications and
runs only missing cases. It requires the exact decorated identity, preserving
the binaries, feature set, source legs, samples, case order and statistical
rules. The offline analyzer and independent audit remain mandatory gates.

Before comparative capture, a retained baseline-reader preflight exposed two
reader assumptions: the shared allocation validator takes the allocation
object directly, and the native no-op report omits that envelope entirely.
Both failed attempts and their exact source snapshots are retained. The fixed
reader treats that one absent native envelope as not recorded; it does not
invent zero counters. The observer no-op still requires measured allocation
vectors. Both independent readers subsequently accepted all 18 baseline
reports, including their corpus/output identities. No workload was repeated
to repair these reader errors.

Both source legs passed all eight quality gates. Each has 2,087 XLSX tests and
three pinned oracle tests passing, plus 555 harness tests passing and one
ignored. The ignored test is the existing opt-in real-producer security-corpus
test. Both release builds per leg passed. Independent qualification admission
accepted all 36 reports before comparative capture began.

## Disposition and remaining work

The root decision adopts only the archived three-line guard. Independent replay
agrees on all 180 reports / 20,304 samples and all 200 quality/build/capture
command receipts. The complete native and observer vectors, failed reader
attempts and original supervisor failure remain retained. The owned target
cleanup removed 11,308 files (8,430,128,751 logical bytes),
and the marked scratch directory contained only its 38-byte marker. The
analysis, independent audit and root decision all replayed successfully after
cleanup without the build executables.

The two bounded parser maps remain the next allocation representation to
investigate; this experiment does not justify removing either parse or
weakening overlapping-column semantics. Broader non-iWork CRUD coverage,
cold/range-source behavior, parallel scaling and the program-level goal remain
open.
