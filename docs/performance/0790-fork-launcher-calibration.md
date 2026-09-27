# 0790 — ordinary-fork launcher calibration

An explicit ordinary-fork launcher removes the large inherited startup
high-water value from all 48 direct matrix children. Their self `ru_maxrss`
starts at 1,360–1,480 KiB while their launcher retains 17,644–18,588 KiB.
The independent GNU-time lane starts at 1,472–1,480 KiB. Both launchers still
show substantial differences between exit resource accounting and observed
resident pages. This narrows the measurement problem without changing any
library code or reconsidering the rejected 0787 candidate.

The complete predeclared acceptance result is **fail**, not pass: three of
four traced controls violate the requirement that every direct child's
startup high water be lower than its launcher's. Their launchers already
have low high water under strace's different process topology. All 48 ordinary
matrix direct children satisfy that requirement. The frozen criterion and all
three failures remain retained; no captures were repeated or replaced.

## Experiment

Base is `99b62729ab`, the committed 0789 calibration. The host remains the
32-CPU EPYC 9R45 Linux `7.0.0-1012-aws` environment with 4 KiB pages. Current
compiler/time/strace/taskset identities, affinity, memory state, original
working-tree inventory, and source hashes are in the
[evidence packet](results/change-0790/README.md).

The known-page C probe is byte-identical to 0789, including its original
`RSS0789`/`ACK0789` marker names. A new archive-only C launcher opens a fresh
usage receipt, samples its own pre-fork resource usage, performs ordinary
`fork`, and executes the probe in the child. It waits for that exact PID using
`wait4`, retaining the child usage and exit status. The parent benchmark driver
still uses `posix_spawn` to execute taskset and the launcher; that outer
process's usage is kept separate from the probe child's usage.

The matrix uses six cases: zero payload/no workers, 4 MiB/main/no workers,
4 MiB/main/four control workers, 4 MiB/four payload workers,
4 MiB/32 payload workers, and 64 MiB/main/no workers. Each uses CPU 0 or all
CPUs, direct launcher or GNU time, identity or full proc observation, and two
repeats with reversed launcher/observer order: **96 matrix children**.
Four separately scoped direct-launcher traces cover zero/64 MiB and both
affinities. Verbose rusage is enabled while execve remains abbreviated to
avoid retaining environment values. Totals are **100 children, 600 acknowledged
checkpoints, 312 full snapshots, and 288 identity snapshots**.

All workloads were run serially by root. Analysis and source review were
separate offline work. This is a descriptive two-repeat calibration, with no
latency estimate, bootstrap interval, physical transient-peak assertion, or
new CRUD coverage.

## Results and acceptance failure

All child identities, exact page-pattern checksums, mapped-byte witnesses,
affinity checks, exits, and retained artifacts validate. Every full snapshot
has equal summed smaps RSS, rollup RSS, and status `VmRSS`. Anonymous
mapped-to-touched residency increases by exactly the payload with no workers,
by payload plus 36 KiB with four workers, and by payload plus 44 KiB with
32 workers. Those extra values describe the whole process, not payload-only
allocation costs.

The direct matrix startup range falls from the prior unwrapped control's
17,708–19,372 KiB to 1,360–1,480 KiB. This is a launcher-scope correction,
not a reduction in the probe's actual memory footprint. No historical timings
or samples are pooled. Direct and GNU-time controls use separate processes;
no claim of exact cross-process equality is required or justified.

The 4 MiB/32-worker full-observer matrix illustrates the remaining accounting
difference. Values below are maximum observed rollup RSS minus the child's
exit resource high water, in KiB, in repeat order:

| Launcher | All CPUs, repeats 0 / 1 | CPU 0, repeats 0 / 1 |
| --- | ---: | ---: |
| Ordinary fork, direct wait4 | 1,528 / 1,528 | −56 / −52 |
| GNU time | 1,528 / 1,268 | −52 / −48 |

Thus a direct kernel wait interface reproduces the counter/residency gap
without GNU-time formatting. Negative differences remain compatible with
unobserved transient peaks; snapshots are not continuous monitoring.
For both launchers, 40 of 48 matrix children have exit RSS equal to their
maximum sampled self RSS; the remaining eight increase by 300 KiB at exit.
The six nonzero matrix ACK changes all occur in direct children: five are
288 KiB and one is 292 KiB. Full/identity observations remain separate, and
these differences are not normalized away or interpreted as observer neutrality.

Detailed traces match all 48 self-rusage triples and all four direct child
wait4 triples exactly. The four exit triples (RSS KiB / minor faults / major
faults) are `1780/80/0`, `1480/77/0`, `67012/16462/0`, and `66944/16462/0`.
These verify the syscall/report correspondence, not identical launcher history.
The trace filter records getrusage/wait4/execve, not process-creation syscalls;
ordinary-fork behavior is supported by exact launcher source/build custody
and observed parent/child identity rather than a traced fork call.

The predeclared acceptance criterion requires every direct child's startup
self high water to be below its launcher's pre-fork self high water. The
following traced children fail it:

| Trace index | Child startup KiB | Launcher pre-fork KiB |
| --- | ---: | ---: |
| 0 | 1480 | 1172 |
| 1 | 1480 | 1172 |
| 2 | 1476 | 1172 |

The outer strace process introduces another fork/exec boundary. Its launcher
child does not retain the same Python pre-exec history as an ordinary matrix
launcher. Consequently “must decrease” is not a valid universal expectation
across those topologies. That interpretation explains why the criterion needs
topology-specific design in future work; it does not retroactively change this
packet's failed predeclared acceptance result. The fourth trace satisfies the
comparison, and its complete values remain in `analysis.json`.

## Verification and architectural scope

Three strict builds pass: the unchanged normal probe, normal launcher, and
ASan/UBSan launcher, with C11, `-Wall -Wextra -Werror`. Seven qualification
checks cover malformed invocation, relative command rejection, failed exec,
nonzero child exit, signal exit propagation, zero payload, and 64 MiB worker
touches. Sanitization covers the launcher; the unchanged probe's separate
0789 sanitizer evidence is referenced by exact source/receipt hashes rather
than claimed as freshly rerun.

The offline analyzer preserves acceptance failures as data. A successful
replay means the packet reproduces those failures; it does not mean the
experiment passed every planned criterion. The independent review checks the
launcher source and raw evidence. No Cargo gates were rerun because no Rust
production or test source changed.

The 9,196 production files and 35 architecture/taxonomy/goal inputs remain
unchanged. ADR 0001/0002/0024 ownership, ADR 0003 edit/publication semantics,
ADR 0005 explicit resource and evidence requirements, ADR 0006 preservation
and security, ADR 0008 verification, and ADR 0010/0011 archive ownership are
unaffected. The C launcher is a bounded external diagnostic with no runtime
dependency or relaxation of Rust unsafe policies. Remaining accepted ADRs have
unchanged inputs and no affected production behavior.

All three owned executables are hash-verified before removal. The final seal
binds scripts, raw captures, derived results, source review, cleanup, and the
six performance documents. Unrelated working files and pre-existing worktrees
remain preserved.

## Next action

The ordinary matrix fork control is usable evidence that pre-exec history can
be isolated, while the traced “must decrease” check is not universally valid.
Future measurements must predeclare qualification separately for ordinary and
traced launch topologies. With those scopes explicit, the next useful Office
experiment can compare allocation lifetimes and acknowledged residency around
the cached-Part operation while preserving the native GNU-time metric and the
historical 0787 rejection. No constant RSS correction or retrospective
threshold relaxation follows.

This supports category-15 measurement work only. OLE2/OOXML remain active,
ODF deferred, iWork excluded, and the broad performance goal remains open.
