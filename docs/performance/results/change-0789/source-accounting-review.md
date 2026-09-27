# 0789 source accounting review

This review is read-only evidence for the 0789 known-resident-page
calibration. It explains how the installed GNU `time` process obtains `%M`,
how Linux obtains `ru_maxrss`, and why that value can be below the values read
from `/proc/<pid>/status` and `smaps`. It does not change the 0787 rejection or
the 0788 measurements.

## Finding

The source tree identifies a specific, testable mechanism for the 0788
discrepancy. On this 32-CPU SMP host, Linux stores the RSS components in
batched per-CPU counters. The `getrusage` high-water path reads the global
part of each counter through `percpu_counter_read_positive`, while the proc
status path sums the global part and all per-CPU residues through
`percpu_counter_sum_positive`. A local residue can therefore be visible to
`VmRSS`/`VmHWM` or a page-table scan while absent from the value returned at
child exit to GNU `time`.

This is a source-backed hypothesis with a strong magnitude match, not a
completed attribution of the historical 0787 RSS increase. The exact Ubuntu
vendor source for every out-of-tree or packaging change is not installed; the
running kernel's 7.0.0-1012-aws headers match the relevant upstream v7.0
definitions, and the 0789 controls are required to connect those definitions
to the observed children.

The 0788 result being investigated is precise: all 32 handshake-enabled
children had GNU-time `%M` below both the largest `smaps_rollup` RSS and
`VmHWM`, with a gap of 156--1,716 KiB. With 4 KiB pages that is 39--429 pages.
The representative baseline small/primed/width-4 child reported `%M = 4052`
KiB while its final snapshot reported 5312 KiB for smaps RSS and status
`VmHWM`/`VmRSS`; the paired candidate reported `%M = 4396` KiB and reached the
same 5312 KiB snapshot. PID, parent, executable, start time, and task identity
were independently checked in the 0788 packet.

## Verified installed-tool facts

The installed executable is `/usr/bin/time`, package `time` version `1.9-0.4`,
SHA-256
`919efdcc04c1dbc7a1479c2f3310d491c95f028e1762a84c47810ac0ec1c0d88`, and GNU
Time reports its 2018 upstream version. Its dynamic symbol table contains
`wait3@GLIBC_2.2.5`; it does not call a separate `getrusage` symbol from the
executable. This establishes the installed entry point without assuming that
the current GNU Time development branch is the installed binary.

The exact GNU Time v1.9 source at commit
`d0a4b03f564a282da1efa981dc806b47c94871a2` has these relevant facts:

* `src/resuse.c` calls `wait3(&status, 0, &resp->ru)` and waits until the
  returned PID is the child being measured. The rusage structure is supplied
  by that wait operation.
* `src/time.c` formats `%M` directly from `resp->ru.ru_maxrss` through
  `get_rusage_maxrss_kb`; it does not read `/proc` and does not perform a
  second RSS sample after the wait.
* `configure.ac` selects `GETRUSAGE_RETURNS_KB` for Linux, and
  `src/rusage-kb.h` makes the memory conversion a no-op in that mode. GNU
  `time` therefore prints the Linux `ru_maxrss` value in KiB; it is not a
  page-count or byte-count formatting error.

The glibc source for `posix/wait3.c` implements `__wait3` as
`__wait4(WAIT_ANY, ...)`. The Linux glibc `wait4` implementation passes the
usage pointer to the Linux `wait4` system call when that ABI is available.
The installed man page states the same library/kernel relationship. Thus the
observed `%M` chain is:

```text
GNU time v1.9 wait3
    -> glibc wait3 wrapper
    -> Linux wait4(pid = any child, rusage pointer)
    -> kernel child-reap rusage
    -> format ru_maxrss in KiB
```

## Launcher inheritance and 0789 scope

The two 0789 launchers do not create the probe through the same process
topology. The capture script records `os.posix_spawn` of `/usr/bin/taskset`.
The direct command is therefore `capture -> taskset -> probe`, while the
GNU-time command is `capture -> taskset -> GNU time -> probe`; the latter has
an additional child created by GNU Time itself.

**Verified implementation facts.** Linux glibc's `posix_spawn` implementation
uses `clone` with `CLONE_VM | CLONE_VFORK`. Its source says that the child runs
in the same memory space until it calls `execve` or `_exit`, and that the
caller is suspended until then. The Linux `exec_mmap` path calls
`setmax_mm_hiwater_rss(&tsk->signal->maxrss, old_mm)` before releasing the old
address space. An ordinary `fork` creates a new zeroed signal structure for a
non-thread child, and `dup_mm` resets the duplicated `mm` high-water value to
the current RSS. GNU Time v1.9 calls ordinary `fork()` and then `execvp()` for
the command. These facts are recorded as `glibc-posix-spawn`,
`linux-v7-exec-mmap`, `linux-v7-copy-signal`, `linux-v7-dup-mm`, and
`gnu-time-v1.9-fork-exec` in `sources.json`.

**Verified 0789 capture facts.** All 136 direct rows have direct parent
`wait4.ru_maxrss` exactly equal to the largest probe `getrusage(RUSAGE_SELF)`
sample in that row. That equality is internal consistency for the direct
child; it does not make the direct value independent of its launcher. Across
the 68 direct full rows, startup self `ru_maxrss` is 17,708--19,372 KiB while
startup status `VmRSS`/`VmHWM` is only 1,804--1,816 KiB. Across the 68 GNU-time
full rows, startup self `ru_maxrss` is 1,364--1,480 KiB and status remains
1,804--1,816 KiB. The direct command's identity and recorded `spawn_method`
show that the first `exec` is the taskset process; the GNU-time command has
the extra ordinary fork in GNU Time.

The source-backed interpretation is that the direct `posix_spawn` child first
executes taskset while its old address space is still the capture process's
Python `mm`. `exec_mmap` folds that old `mm` high-water into the new child's
`signal->maxrss`; taskset then execs the probe, and subsequent exec high-water
updates retain the maximum. This accounts for a probe self high-water near
the capture process's 17--19 MiB high-water despite a roughly 1.8 MiB current
probe RSS. In the GNU-time path, GNU Time's ordinary `fork` supplies a fresh
signal structure and a duplicated `mm` whose high-water is reset to the
wrapper's current RSS before the probe exec. That accounts for the clean
roughly 1.4 MiB probe startup sample. This is a source-backed inference that
matches the captures; the exact Ubuntu vendor C path was not independently
diffed.

For the 4 MiB, 32-worker, full GNU-time rows, the observed differences are
also topology-qualified rather than adoption evidence:

| Matrix row | Affinity | GNU `%M` (KiB) | Maximum smaps RSS (KiB) | Maximum status `VmHWM` (KiB) | smaps minus `%M` (KiB) |
| --- | --- | ---: | ---: | ---: | ---: |
| `matrix/131` (repeat 0) | all CPUs | 4,296 | 6,080 | 6,080 | +1,784 |
| `matrix/135` (repeat 0) | one CPU | 6,132 | 6,084 | 6,132 | -48 |
| `matrix/264` (repeat 1) | all CPUs | 4,040 | 6,080 | 6,080 | +2,040 |
| `matrix/268` (repeat 1) | one CPU | 6,132 | 6,080 | 6,132 | -52 |

These rows show why the known-page and affinity controls remain useful, but
they do not turn GNU-time `%M` and the current direct `wait4` topology into
interchangeable measurements.

## Verified Linux accounting paths

The upstream Linux v7.0 source and the host's installed 7.0.0-1012-aws
headers show the following paths.

1. `wait4` calls the kernel wait machinery with a `struct rusage`. On child
   reap, the kernel asks `getrusage(p, RUSAGE_BOTH, ...)` for that child. The
   uapi header defines `RUSAGE_BOTH` specifically for `sys_wait4`.
2. At group exit, `kernel/exit.c` records the process high-water value with
   `setmax_mm_hiwater_rss(&tsk->signal->maxrss, tsk->mm)`.
3. `getrusage` obtains the child/group maximum and converts pages to KiB with
   `PAGE_SIZE / 1024`. Linux therefore supplies KiB to GNU Time on this host.
4. `include/linux/mm.h` defines `get_mm_rss()` by calling
   `get_mm_counter()`, and `get_mm_counter()` calls
   `percpu_counter_read_positive()`. On SMP that accessor reads the shared
   `fbc->count` and deliberately does not sum the per-CPU slots.
5. The same header defines `get_mm_rss_sum()` using
   `get_mm_counter_sum()`, which calls `percpu_counter_sum_positive()`.
   Linux's proc `task_mem()` uses this summed accessor for `RssAnon`,
   `RssFile`, `RssShmem`, `VmRSS`, and the current component of `VmHWM`.
6. `fs/proc/task_mmu.c` explicitly explains why proc collection reads both
   the cached high-water field and current RSS: the kernel updates high-water
   values only when it is about to lower RSS, and snapshots can be
   inconsistent. It computes `hiwater_rss = max(mm->hiwater_rss, total_rss)`.
   The `smaps` and `smaps_rollup` paths instead walk the page tables and report
   their own point-in-time resident totals.

The per-CPU counter implementation gives the scale of the possible
difference. `percpu_counter_add_batch()` accumulates local changes until the
configured batch threshold and then folds them into the shared count.
`__percpu_counter_sum()` takes the counter lock and adds every online and
dying CPU's local value. In v7.0 the default batch is
`max(32, nr_online_cpus() * 2)`. The host has 32 online CPUs, so the default
threshold is 64 pages. The observed 39--429-page gaps are compatible with
residues across RSS categories and CPUs, but compatibility alone does not show
which residues occurred in a particular child.

This also explains why a normal proc snapshot should not be treated as the
same measurement as GNU-time `%M`:

| Observation | Source path | Scope |
| --- | --- | --- |
| GNU-time `%M` | child `wait4` rusage, `ru_maxrss` | kernel child high-water accounting at reap |
| direct parent `wait4` rusage | same Linux wait4 rusage path | child high-water accounting; current direct topology carries pre-exec inherited high water |
| status `VmRSS` | summed per-CPU RSS components | point-in-time proc counter snapshot |
| status `VmHWM` | cached high-water or summed current RSS, whichever is larger | proc high-water view |
| `smaps`/`smaps_rollup` RSS | page-table walk | point-in-time resident-page snapshot |
| probe `getrusage(RUSAGE_SELF)` | live process getrusage path | process-local sample, with current approximate RSS folded into high-water |

The first two rows should agree when they observe the same child lifecycle and
the launcher has not added inherited high water. The current direct and
GNU-time rows do not satisfy that condition because GNU Time inserts an
ordinary `fork` after the taskset exec. The remaining rows have different
sampling and counter scopes, so their disagreement is expected to be
diagnostic rather than an automatic data-quality failure.

## Decisive controls

The frozen 0789 plan contains the controls needed to distinguish the source
paths without running another office workload.

* The standalone probe samples `getrusage(RUSAGE_SELF)` before and after every
  ACK. The final phase ACK keeps the child alive while the parent captures its
  final proc state. A self value that remains below summed proc RSS while the
  child is blocked supports the counter-path hypothesis; a self value that
  catches up only after an additional phase points to timing or a final
  high-water update.
* The direct launcher uses the parent process's `wait4` result on the probe
  child, and the GNU-time launcher records `%M` from GNU Time's ordinary
  forked child's `wait3`/`wait4` result. In this packet the direct `wait4`
  equals direct self in all 136 rows, but both carry the direct launcher's
  inherited high water. A decisive follow-up must launch the probe through an
  ordinary `fork`/`exec` child or a fresh low-high-water helper, then compare
  that child's self sample, its parent's direct `wait4`, and GNU Time's `%M`.
  The parent `wait4` on the GNU-time wrapper is retained separately and must
  not be mistaken for the probe child's `%M`.
* The probe maps and touches a known number of anonymous pages at 0, 1, 4, 16,
  and 64 MiB. Its checksum and mapped-byte fields verify that the intended
  pages were touched. Comparing the expected page count with self rusage,
  direct wait4, GNU-time `%M`, status, and smaps makes the accounting error
  measurable without allocator or office-format noise.
* Each case runs on one CPU and on all 32 CPUs. A per-CPU residue mechanism
  should change when page faults and worker touches are constrained to one
  CPU; a difference that is invariant under affinity needs another explanation.
  Main-thread and worker-thread touch routes separate page faults from thread
  creation and joining.
* The 0789 phase protocol captures identity-only and full proc snapshots in a
  fixed order. Workers join before the relevant post-touch observation, and
  the parent waits on the same process whose PID, executable, parent, and
  start time it checked. This prevents a launcher or PID-reuse explanation
  from being silently mixed into the accounting comparison.

No one control is sufficient. In particular, a final proc snapshot cannot
prove that no transient peak occurred, and a two-repeat result cannot support
a confidence interval. The purpose of 0789 is to reconcile the accounting
interfaces before any future adoption experiment is designed.

## Source limits and disposition

The host package exposes the exact `mm.h`, `percpu_counter.h`, and uapi
resource definitions through `linux-aws-headers-7.0.0-1012`. The complete
Ubuntu kernel C sources are not installed, so the `kernel/sys.c`,
`kernel/exit.c`, `fs/proc/task_mmu.c`, `fs/exec.c`, `kernel/fork.c`, and
`lib/percpu_counter.c` statements are checked against upstream v7.0. The
glibc source is likewise checked against upstream glibc 2.43 rather than a
vendor source package. The running kernel is an Ubuntu AWS build
(`7.0.0-1012-aws`); vendor patches were not independently diffed. This is why
the source review calls the launcher inheritance a source-backed inference
and the accounting mechanism testable rather than declaring the 0787 guard
invalid.

The historical 0787 rejection remains authoritative. This review recommends
using the 0789 direct/self/known-page/affinity results to qualify the memory
measurement, with no production change and no threshold revision.
