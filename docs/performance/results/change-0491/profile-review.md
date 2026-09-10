# Whole-child provider diagnostics

These profiles include deterministic corpus construction, preflight, warmups, the 30 timed operations, oracles and report serialization. They are not operation-only CPU or syscall attribution, and do not establish a speedup. The normal binary and canonical capture gate are bound by profile verification.

| Arm | Cycles | Instructions | Whole-child IPC | Branch misses | Page faults | Elapsed ms |
|---|---:|---:|---:|---:|---:|---:|
| file | 2,550,213,183 | 6,743,434,698 | 2.644 | 24,813,942 | 29,561 | 646.61 |
| range-64-0us | 2,545,165,912 | 6,817,929,449 | 2.679 | 25,037,327 | 29,564 | 595.76 |
| range-65536-1000us-104857600bps-minimum-service | 2,559,880,602 | 6,802,545,648 | 2.657 | 25,105,682 | 29,572 | 1275.70 |

PMU events ran for roughly 80–84% of the interval; exact raw CSV multiplex metadata is retained. Cache references report zero despite positive cache misses, so a cache miss rate or cache-locality conclusion is not justified. The delayed child spends wall time in simulated service delay; the whole-child instruction counts also include the much larger fixture-generation work.

The selected file strace reports 633 pread64, 627 statx, 17,572 write, 9 read, and 2 fsync calls. fdatasync has no row. The two setup syncs and report/setup writes are outside the timed read lifecycle. The trace has no per-operation syscall boundary and is not atomic-save evidence.

Recheck all four profiles with `python3 -B docs/performance/results/change-0491/profile_providers.py --verify profiles2`. The verifier checks exact result/header/artifact inventory, build/source/tool identity, timing order, parsed raw counters, benchmark reports and private-directory cleanup.
