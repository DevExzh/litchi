# 0789 — Linux RSS accounting calibration

The standalone known-page experiment reproduces a substantial difference
between Linux resource-usage high water and observed resident memory. GNU time
prints the kernel's returned value exactly in all four detailed syscall traces.
For a 4 MiB mapping touched by 32 workers across all CPUs, maximum observed RSS
exceeds GNU-time `%M` by 1,784 and 2,040 KiB. Pinning the same case to one CPU
changes that difference to −48 and −52 KiB. This supports the source-backed
per-CPU accounting explanation for counter disagreement; it does not identify
the cause of the historical 0787 candidate's 222 KiB increase.

No library change, adoption, regression-threshold revision, or performance
improvement is claimed. The 0787 rejection remains authoritative. This packet
qualifies measurement interfaces before another Office workload experiment.

## Protocol and scope

Base revision is `5a26710a10af2e44829daa04def3edf0962a80ea`. The host is the
same 32-CPU AMD EPYC 9R45 Linux `7.0.0-1012-aws` environment, with 4 KiB pages,
GNU time package `1.9-0.4`, and libc `2.43-2ubuntu2.4`. Exact host/tool/source
identities, compiler commands, and the pre-build plan are retained in the
[packet](results/change-0789/README.md).

The archive-only C11 probe maps 0/1/4/16/64 MiB of anonymous memory and writes
one volatile byte per page. Main-thread readback verifies every byte and a
deterministic checksum. Cases use no workers, four control workers after main
thread touches, or four payload workers; the 4 MiB cases also use 32 workers.
The two affinities permit CPU 0 only or all CPUs 0–31. Workers are joined
before both `touched` and `workers_joined`; these are two checkpoints of the
same joined state, not measurements of live workers.

Every child acknowledges six phases: startup, mapped, touched, workers joined,
unmapped, and final. It samples self `getrusage` before and after each ACK. The
parent records PID, parent, executable, start time, and task IDs at each phase.
Full observations retain `smaps`, `smaps_rollup`, maps, status, and stat;
identity observations retain stat and identity only. Identity observations
still inspect procfs and are not an unobserved control.

The frozen matrix has 17 cases × two affinities × four launcher/observer
combinations × two repeats = **272 children**. The second repeat reverses the
four launcher/observer positions. All children run serially; this is an
accounting calibration with no timing comparison or confidence interval.
There are 136 full and 136 identity matrix children, each with six checkpoints.

Four additional predeclared GNU-time traces verify the syscall path. Their
original strace command abbreviates rusage fields, so they cannot prove the
numeric correspondence. A separately recorded supplement adds `-v` and
captures four new children with the same zero/64 MiB and affinity controls.
All originals remain retained. Totals including that supplement are
**280 children, 864 full snapshots, 816 identity snapshots, and 1,680
checkpoints**, plus nine separate build qualification checks. The primary
analyzer deliberately reports its original 276-child/840-full-snapshot scope;
the independent summary includes the four supplemental children.

## Results

All 864 full snapshots have equal summed smaps RSS, rollup RSS, and status
`VmRSS`. No discrepancy is normalized away. With no workers, anonymous
mapped-to-touched residency increases by exactly the requested payload in
every full observation, including all 64 MiB controls. Four-worker cases add
36 KiB beyond the payload; 32-worker cases add 44 KiB. Those are observed
process-wide anonymous deltas, not payload-only allocation or peak estimates.

The table gives maximum observed RSS minus GNU-time `%M`, in KiB, for the
**full-observer matrix children**, in repeat order. Positive means observed
residency exceeds the exit resource counter. Negative does not establish an
accounting error: a transient peak may precede a checkpoint.

| Payload and touch route | All CPUs, repeats 0 / 1 | CPU 0, repeats 0 / 1 |
| --- | ---: | ---: |
| 4 MiB, main, no workers | 328 / 332 | 332 / 328 |
| 4 MiB, main, four control workers | 324 / 492 | 192 / 196 |
| 4 MiB, four payload workers | 496 / 496 | 196 / 196 |
| 4 MiB, main, 32 control workers | 504 / 500 | −52 / −52 |
| 4 MiB, 32 payload workers | 1,784 / 2,040 | −48 / −52 |
| 64 MiB, main, no workers | 332 / 332 | 336 / 336 |

64 of 68 full-observer GNU-time matrix children have positive gaps; the four
negative cases are the single-CPU 32-worker controls. Affinity changes these
observations without changing the known touched payload, consistent with
per-CPU residue effects. It does not isolate every scheduling or stack effect.

The detailed traces match all 48 self-rusage samples to their protocol lines,
and all four GNU-time `%M/%R/%F` triples to the exact child `wait4` result:

| Payload | Affinity | GNU time and traced wait4: RSS KiB / minor faults / major faults |
| --- | --- | ---: |
| 0 MiB | CPU 0 | 1780 / 79 / 0 |
| 0 MiB | all CPUs | 1780 / 78 / 0 |
| 64 MiB | CPU 0 | 67016 / 16463 / 0 |
| 64 MiB | all CPUs | 67016 / 16464 / 0 |

This excludes a GNU-time formatting/unit mismatch in those observed children.
The values returned by the parent wait on the GNU-time or strace wrapper have
a different scope and are retained separately.

## Launcher and observer limitations

The direct `posix_spawn` route starts with self maximum RSS of
17,708–19,372 KiB, while GNU-time probe children start at 1,364–1,480 KiB.
The direct route's pre-exec high-water history masks most small mappings.
Its exit counter equals the maximum self sample in all 136 matrix children,
but comparing its small-process absolute high water with the GNU-time route
would be misleading. For the larger 64 MiB main-thread control, direct and
GNU-time routes both have positive observed-minus-exit gaps of 328–336 KiB.
The source review explains why exec history must be distinguished from a
fresh post-exec residency observation.

121 of 136 GNU-time matrix children have exit RSS equal to maximum sampled
self RSS; the other 15 increase by 300 KiB at exit. Thus even a final self
sample is not universally equal to the exit value. Across the 1,632 matrix
ACK intervals, self RSS increases by 288 KiB three times under full observation
and twice under identity observation; all other intervals show zero change.
These are descriptive observations, not proof that proc snapshots have no
observer effect. Before-ACK formatting/output, scheduling, and counter updates
also occur between samples.

The [source review](results/change-0789/source-accounting-review.md) and
[source references](results/change-0789/sources.json) identify Linux's
approximate shared per-CPU RSS counters on the resource-usage path and summed
counters on the proc status path. Installed kernel headers match those relevant
upstream v7.0 definitions. Complete Ubuntu vendor sources were not diffed, and
individual kernel counter residues were not instrumented. The observed
affinity dependence supports this mechanism without proving the origin of a
specific historical Office-process result. No constant offset correction is
justified by these data.

## Verification and custody

Strict C11 normal and ASan/UBSan builds pass with `-Wall -Wextra -Werror`.
Six negative checks reject invalid size, worker count, touch mode, missing
workers, EOF ACK, and wrong ACK. Three sanitized positive runs exercise zero
allocation, 64 MiB worker touches, and 32 control workers. All checks retain
stdout, stderr, exit status, compiler commands, and source/binary hashes.

The probe's pre-build review corrected two missing integer format markers,
made mapped-byte reporting literal at each phase, and made an unexpected
join failure terminate before worker storage can be released. None of these
changes followed measurement. Analyzer corrections and the separate detailed
trace supplement are recorded; raw measurements are never replaced.

The production census remains exactly 9,196 files, and all 35 architecture,
ADR-index, goal, and taxonomy inputs are unchanged. Rust quality gates were
not rerun for this standalone diagnostic: there is no Rust production or test
change. The C probe owns bounded mapping/thread storage outside the library;
it neither changes Rust unsafe policies nor introduces a runtime dependency.
Root alone built and ran children; source and analysis agents worked offline.

| Contract | Disposition |
| --- | --- |
| ADR 0001/0002/0024 ownership and topology | No library/package/dependency changes |
| ADR 0003 snapshots, edits, commits, patches | No document operation or publication changes |
| ADR 0005 explicit resources and performance evidence | Fixed mapping/worker limits; isolated, source-bound observations; scope differences retained |
| ADR 0006 preservation/security and ADR 0008 verification | No preservation/refusal changes; no CRUD or native-producer promotion |
| ADR 0010/0011 physical-package ownership | No archive or OPC ownership changes |
| Remaining accepted ADRs | Inputs unchanged; no production behavior affected |

The packet retains replay scripts, raw captures, source references, and a final
hash seal. Both temporary executables are verified immediately before removal;
cleanup records their identities and permits offline replay after removal.
Unrelated working files and pre-existing worktrees are preserved.

## Next bounded step

The small cached-Part measurement must distinguish exit resource accounting
from observed residency and allocation peaks. First qualify a direct
ordinary-fork/exec control with the same low-memory launcher history as the
GNU-time child; the current direct posix_spawn route is unsuitable for small
absolute high-water comparisons. A future experiment should
predeclare its launcher, affinity controls, phase observations, and decision
policy before capture, and retain the existing native RSS metric alongside
those independent observations. This calibration neither retries adoption nor
relaxes the failed 0787 policy. Whole-child `%M` should not be described as an
exact physical residency peak for this small-process protocol.

This is supporting evidence for category 15, with no new CRUD baseline or
capability promotion. OLE2/OOXML remain active, ODF remains deferred, iWork is
excluded, and the broad performance goal remains open.
