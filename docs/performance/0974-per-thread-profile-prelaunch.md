# 0974 — per-thread mmap profile diagnostic stops before native launch

**Status: capability-only diagnostic; native prelaunch attempt failed; repair
review complete.** The
0974 design tested the smallest follow-up to the 0973 loss result: explicit
`-m 16M` together with `--per-thread`, while retaining the event, period,
32 KiB DWARF stack, CPU-31 outer affinity, launcher, fixture, and zero-loss
stop-without-retry gate. No Office workload, target build, production source
edit, or production/runtime change occurred.

## Capability and lifecycle evidence

The explicit `-m 16M` probe using the default per-CPU mapping failed during
perf ring mapping with a permission error before workload execution. Adding
`-C 31` failed earlier at event setup:
the `instructions:u` event was denied under the host's perf permissions. Both
are setup failures, not loss-free measurements.

The `--per-thread -m 16M` probe succeeded in setup for `/usr/bin/true`: perf
reported a 16,781,312-byte mapping, produced a 44,336-byte `perf.data`, and
returned zero. Its metadata had one tracked thread and CPU map `-1`, consistent
with per-thread mapping and disabled inheritance. `true` was too short to
produce a period-1M sample, so this proves setup capability only; it does not
prove sustained ring capacity or an Office capture.

The deliberate lifecycle negative control returned zero and decoded 175 sample
records on the main PID/TID, while exposing two `PERF_RECORD_FORK` records: one
for a worker thread and one for a waited child. It demonstrates that the raw
task stream can expose scope violations that a native per-thread run must
reject. It is not a native or Office workload result.

## Prelaunch attempt and repair proof

The authorized runner attempt stopped at child 01 with return code 2 and
`RecursionError: maximum recursion depth exceeded`. The runner created an empty
child directory, then the v3 adapter's monkeypatch cycle recursed while building
the command. It never constructed native perf argv, reached `subprocess.Popen`,
or launched a process. The terminal records zero packets, children 01–12 not
launched, and no retry. The corrective audit binds this as a
`prelaunch_tooling_failure` with `no_admission`; there is no loss, marker,
stack, lifecycle, timing, or performance result from the attempt.

The v5 repair was checked through the real `run_child` command-generation path
with a `Popen` sentinel. It made exactly one intercepted call with the
canonical `--per-thread -m 16M` argv, launched no process, and verified
idempotent installation plus restoration of the raw-scanner aliases. The v1/v3
recursion failure remains preserved; the earlier v4 focused test remains
historical because it did not exercise this actual boundary. The [independent
offline repair review](results/change-0974/receipts/independent-repair-review-0974-v5-v1.json)
(SHA-256 `b74e7fb3c04f88ea57a4c9b2f6fd73d8ebbc5894c7b80bbaf9c3af93c4826c97`)
passes the repair evidence and confirms that no native process or workload was
started.

## Interpretation and disposition

The probes establish a host-specific capability distinction: explicit large
per-CPU mapping and CPU-filtered recording are unavailable under the observed
permissions, while per-thread mapping can be created for a short process. They
do not establish the cause of the 0973 loss, show that per-thread recording can
capture the retained native workload, or identify a performance hotspot or
causal owner. No 0974 native cohort was admitted, and no
performance, adoption, or production claim follows. Any future native run is a
newly reviewed batch after this prelaunch failure, with the existing lifecycle
and zero-loss gates retained.

## Evidence

| Evidence | SHA-256 |
| --- | --- |
| [v5 plan](results/change-0974/plans/plan-0974-pre-capture-v5.json) | `dddd126895d04cd1738719d99aa00ff1b68dcd551233a5c2391f37f7c30e71ee` |
| [v5 helper seal](results/change-0974/receipts/helper-seal-0974-v5.json) | `4cc5ef36a28c6b254a580abd01b21792daef317b383ac5dc718ea544bb3668a3` |
| [prelaunch root terminal](results/change-0974/receipts/root-terminal-0974-v1.json) | `403d58e6471a5f2afbfa8da4eb019904e7cdc46da3ad5f2677ee43cd9bad73de` |
| [corrective prelaunch audit](results/change-0974/receipts/corrective-audit-0974-prelaunch-no-admission-v1.json) | `078cfefe0694486d3b0100044b37c494862fa27d12e2a5d51636f5491e318ecd` |
| [v1 boundary regression](results/change-0974/receipts/launch-boundary-regression-0974-v1.json) | `7e5da9d71647ee4b8cb3625e291da49a671ebb05330457d7acd551d896efdb78` |
| [v5 boundary repair proof](results/change-0974/receipts/launch-boundary-regression-0974-v3.json) | `68c3c2b1df6027218d7407253a255bf770b5d7d875a9e48125c4134fac89cdd4` |
| [independent v5 repair review](results/change-0974/receipts/independent-repair-review-0974-v5-v1.json) | `b74e7fb3c04f88ea57a4c9b2f6fd73d8ebbc5894c7b80bbaf9c3af93c4826c97` |
| [historical v4 pre-capture review](results/change-0974/receipts/independent-review-0974-v4.json) | `86688588f55d80274b906f092d31b82c01e05f137ea86f536e7ac2b4c6fc0cca` |
| [per-thread capability probe](results/change-0974/receipts/capability-perthread-mmap16m-v1.json) | `503cacc304daf95d8e7254feb1bf074d01cf707bfbeb6c11b049a857522705fd` |
| [CPU-filter capability probe](results/change-0974/receipts/capability-cpu31-mmap16m-v1.json) | `70b64d1f18a13c59492777cb8bca8d5a4e461436a9fed65cea0ff4c4e1021145` |
| [mmap capability failure](results/change-0974/receipts/capability-mmap16m-v2.json) | `eb7e8873cdb7b2a2ead8edd1f476150efad11d79a5142fcc7a0492051cbb4786` |
| [lifecycle negative control](results/change-0974/receipts/lifecycle-negative-control-v1.json) | `852e22aadbebb43d75075f9f1749343b50f4008d6a8c37b6a1f35cf954895ec8` |
| [per-thread read-only review](results/change-0974/analysis/per-thread-mmap-readonly-v1.md) | `1ae15da07a1a7f80e92ae84e338c70d41fd9cd18f80dad57a9f442a05d9acc1a` |

The corrective audit and repair proof are offline/tooling evidence only. The
0974 disposition remains **no native admission, no performance claim, and no
production change**.
