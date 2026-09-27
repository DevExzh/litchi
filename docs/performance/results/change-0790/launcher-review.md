# 0790 launcher review

This is an independent read-only review of the 0790 launcher packet. It
reviews the launcher source, the inherited 0789 source custody, and the raw
matrix and verbose-trace receipts. It does not modify the 0789 probe, change a
production crate, or turn the launcher into a runtime dependency.

## Source-backed reason for the launcher

The 0789 direct lane starts `/usr/bin/taskset` with `os.posix_spawn`. On Linux
glibc, `posix_spawn` uses a `clone` path with `CLONE_VM | CLONE_VFORK` until the
child execs. Linux `exec_mmap` carries the old address space's RSS high-water
into the new task's signal high-water. The direct 0789 probe therefore began
with self `ru_maxrss` of 17,708--19,372 KiB while its initial proc RSS was only
about 1,800 KiB. The 0789 GNU-time route inserted an ordinary `fork` before
exec and began near 1,364--1,480 KiB.

The installed 7.0.0-1012 headers show the separate accounting paths: the
resource path uses `setmax_mm_hiwater_rss`, while `get_mm_rss` reads the
approximate per-CPU counters and proc reporting can sum them. The source
references and the exact installed-header hashes are retained in
`../change-0789/source-accounting-review.md` and `../change-0789/sources.json`.
This review uses that source-backed mechanism as a testable explanation; it
does not claim that upstream source alone proves every Ubuntu vendor detail.

The 0790 launcher addresses the inherited high-water path by doing the
ordinary fork after it has started. Its parent records `getrusage(RUSAGE_SELF)`
before `fork`, the child closes the usage receipt and directly `execv`s the
absolute probe path, and the parent waits for that exact child with
`wait4(child_pid, ...)`. The usage receipt records the child PID, status, exit
code, launcher pre-fork maximum, and the wait4 rusage. The launcher source has
no `posix_spawn`, `vfork`, `clone`, shell, `taskset`, or GNU-time call; the
outer taskset process is replaced by the launcher before the launcher forks.

An ordinary fork gives the target a fresh signal accounting structure and a
duplicated mm whose high-water starts from the launcher's current small RSS.
That is the property being qualified. It is not equivalent to claiming that a
phase sample sees every transient peak, or that target self rusage must equal
the later wait4 result.

## Raw acceptance review

The 0789 probe is byte-for-byte preserved. Its SHA-256 is
`fb4359028a1812d1e1d7c9787176b482abf9b7d3ca9293d7f5363a59643fe1f2`, matching
both the inherited 0789 artifact and the 0790 `probe.c`. The normal probe,
normal launcher, and sanitized launcher were built under the frozen source
list. The seven sanitized-launcher qualifications cover invalid invocation,
relative target refusal, failed exec, a nonzero target, a signaled target, a
zero-page probe, and a 64 MiB worker probe; each has retained output, status,
and usage artifacts.

The packet contains 96 matrix children and four separate direct verbose
traces. The matrix has 288 full and 288 identity phase snapshots; the four
traces add 24 full snapshots, giving 312 full snapshots overall. Every
receipt and compressed point artifact passed its recorded size and SHA-256
check, every successful child exited zero, and the raw protocol contains six
phases with the expected ACK exchange. The target's executable, PID, parent,
start time, and task list remain stable across checkpoints. Direct matrix
targets have the launcher as their parent and only the target task in their
task list; the full observations do not see a live worker task.

The matrix high-water bound passes for all 48 direct matrix children:

| observation | range in the direct matrix |
| --- | ---: |
| target startup self `ru_maxrss` | 1,360--1,480 KiB |
| launcher pre-fork self `ru_maxrss` | 17,644--18,588 KiB |
| target startup below the old 0789 direct minimum | 1,480 KiB < 17,708 KiB |

Every direct matrix target startup is strictly below the launcher's recorded
pre-fork high water. This is strong bounded evidence that the high 0789 caller
history did not survive the launcher fork in the matrix lane.

The frozen acceptance wording says “for every direct child,” which also covers
the four separately traced controls. That universal criterion is not fully
met: three traced controls fail the strict startup-versus-launcher comparison.
The trace launcher start is 1,172 KiB; the corresponding target startup
values are 1,480, 1,480, and 1,476 KiB for trace indices 0, 1, and 2. Trace
index 3 is 1,120 KiB and passes. These failures occur with the low-memory
`strace` observer, whose launcher baseline is lower than the target's normal
loader startup. They do not show inherited 0789 high water, but they do mean
the predeclared all-direct bound cannot be reported as passing. The matrix
qualification and the verbose trace qualification must remain separately
labelled; no rerun or post-hoc threshold change is justified.

The four verbose traces do establish the intended wait4 path. Each contains
12 target self `getrusage` calls and one launcher `wait4` of the exact target
PID. The wait4 triples, in trace order, are:

| affinity | payload | wait4 `ru_maxrss / ru_minflt / ru_majflt` |
| --- | ---: | ---: |
| CPU 0 | 0 MiB | 1,780 / 80 / 0 |
| all CPUs | 0 MiB | 1,480 / 77 / 0 |
| CPU 0 | 64 MiB | 67,012 / 16,462 / 0 |
| all CPUs | 64 MiB | 66,944 / 16,462 / 0 |

The launcher usage receipts match those four traced wait4 triples. The 0 MiB
CPU-0 control also demonstrates why no equality acceptance rule is valid:
wait4 reports 1,780 KiB while the largest acknowledged target self sample is
1,480 KiB. That difference is an unobserved exit-time high-water observation,
not a parser error. The 0790 report must retain direct self, launcher wait4,
GNU-time, proc RSS, and smaps as distinct measurements.

The 24 matrix full observations with no workers all show an exact mapped-to-
touched anonymous increase equal to the requested payload, including the
0/4/64 MiB cases under both affinities and both launchers. This validates the
known-page witness while leaving the transient-peak limitation explicit.

The verbose strace filter records `execve`, `getrusage`, and `wait4`, but it
does not include a fork/clone syscall event. The ordinary-fork claim therefore
rests on the frozen launcher source, strict build/source custody, the observed
launcher-to-target parent topology, and the matrix high-water separation. A
future packet seeking syscall-level proof should predeclare a process-syscall
trace such as `trace=process`; that omission is a review limitation, not a
reason to reinterpret the three traced-bound failures as passes.

## Disposition

The 0790 packet qualifies the ordinary-fork launcher for the 48-child matrix
lane and supplies exact wait4 path evidence. It does not satisfy the frozen
“every direct child” startup-bound criterion because of the three traced
controls, so the packet must report that criterion as failed with the raw
indices and values above. The result supports a bounded measurement correction
for future diagnostics: direct 0789-style posix_spawn startup high water and
ordinary-fork target high water are different scopes. It does not revise the
0787 rejection, alter a production threshold, claim a performance gain, or
adopt a library candidate.
