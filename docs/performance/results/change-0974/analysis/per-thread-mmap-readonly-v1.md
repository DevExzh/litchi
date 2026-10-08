# 0974 per-thread mmap read-only review

This note records installed perf documentation, retained upstream perf source
references, and the one authorized harmless probe. It does not authorize an
Office workload, a build, a sysctl/rlimit change, or a fallback.

## Installed behavior

The installed perf 7.0.14 help advertises --per-thread as "use per-thread
mmaps." The captured installed man page says that it overrides the default
per-CPU mmaps, automatically disables inheritance, and is ignored with -a or
-C (perf-record-man.txt lines 502-506). The same man page describes -C as a
selected-CPU mode that creates a system-wide dummy event for sideband
tracking (lines 366-374).

The direct captures are:

- /home/zhuhe/litchi-performance-0974/analysis/cpu-filter-mmap-readonly-v1/perf-record-help.txt
  SHA-256 7f9d5101c37496de4bd6c0af78869ae399eb36d9c51a7b78e3bcd4e75c3b3200
- /home/zhuhe/litchi-performance-0974/analysis/cpu-filter-mmap-readonly-v1/perf-record-man.txt
  SHA-256 38ad97a73c5df417d171bf3ef58358011346f85d9ec89cac77428da70e81afb1
- /usr/bin/perf version 7.0.14
  SHA-256 174229fb48b8557aafa821dd828d883dce47c35c31067e1f31e4a01a91c6c98b

## Source mechanism

Read-only upstream Linux v7.0 perf sources are retained under
/home/zhuhe/litchi-performance-0974/analysis/linux-v7-perf-source. Their
hashes are recorded in the earlier CPU-map analysis. Relevant code paths are:

- target.c lines 59-64: per-thread is incompatible with an explicit CPU list
  or system-wide target.
- evlist.c lines 1047-1086: per-thread-only record setup keeps a thread map
  and an any-CPU map; it does not enumerate all CPUs for the target.
- lib-perf-evlist.c lines 707-750: an any/empty CPU map selects the
  per-thread mmap path, and the mmap count is one per tracked thread.
  mmap_per_thread lines 634-673 maps each tracked thread and only separately
  maps system-wide CPUs if such events exist.
- evlist.c lines 913-917: the data mapping length is (pages + one metadata
  page) times page size.
- builtin-record.c lines 936-970: the selected-CPU sideband dummy is added
  when a CPU list is in use; no CPU list avoids that particular requirement.

The installed binary and man page are authoritative for this host; the
upstream source is corroborating mechanism evidence and may differ in
downstream details.

## Empirical result

The authorized per-thread probe is:

  /usr/bin/taskset --cpu-list 31 /usr/bin/perf record -v --per-thread
  -e instructions:u -c 1000000 -m 16M --call-graph dwarf,32768
  --period --timestamp --sample-cpu --buildid-all --no-buildid-cache
  --user-callchains -o .../true.perf.data -- /usr/bin/true

Receipt:
 /home/zhuhe/litchi-performance-0974/capability-perthread-mmap16m-v1/capability-receipt-v1.json
SHA-256 503cacc304daf95d8e7254feb1bf074d01cf707bfbeb6c11b049a857522705fd

It returned 0 and produced a 44,336-byte perf.data. Verbose perf logged
mmap size 16781312B, exactly 4097 pages at 4096 bytes each. The header
contains one instructions event, one thread-map entry, CPU map -1, and no
dummy event. It retains the same sample payload as the prior mode:
IP/TID/TIME/CALLCHAIN/CPU/PERIOD/REGS_USER/STACK_USER, 32,768-byte user
stack, and read format ID/LOST. The decoded file has metadata only and zero
sample/loss records because true is too short. This is setup evidence, not a
zero-loss workload result.

The prior -C31 probe returned 255 before mmap and verbose perf reported
failure opening both instructions:u and dummy:u under perf_event_paranoid=1.
That result is preserved separately and is evidence against using -C31 as a
resource workaround.

## Recommendation and boundary

The smallest currently demonstrated operational route is per-thread mode
with the existing outer taskset CPU31 and no -C. It avoids the selected-CPU
dummy path and creates one ring for one tracked thread in the harmless probe.
The mode automatically disables inheritance, so the native workload must stay
in the launcher's PID after its execve and must not create child processes or
additional threads. A future design must retain launcher PID/exec PID,
raw task/FORK/EXIT evidence where emitted, and bounded thread snapshots; any
descendant or extra thread makes the capture unsupported rather than hiding
that work.

No conclusion about sustained 16 MiB capacity, sample loss, instruction
variation, native ownership, latency, or speedup follows from this probe.
