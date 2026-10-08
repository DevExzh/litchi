# 0975 fresh per-thread instruction diagnostic

Status: prospective only. No build, native workload, or capture is authorized by
this document. 0975 is a fresh cohort after the retained 0974 pre-launch
failure; it is not a retry of the occupied 0974 output root.

The batch reuses the repaired, hash-bound 0974 v5 runner and the retained v4
reducer. The only new implementation layer is a small 0975 path/custody wrapper
that binds those immutable helpers to this batch. The wrapper must not alter
native argv, decoder arithmetic, task-scope rules, or loss accounting.

The fixed native command remains serial and pinned to CPU 31: `instructions:u`,
period 1,000,000, DWARF callchains capped at 32,768 bytes, `--per-thread`, and
`-m 16M` immediately before `--call-graph`. Each of the 12 children uses the
same retained baseline executable, fixture, 5 warmups, and 80 measured samples.
The runner requires an explicit root GO, refuses an existing capture root, and
stops on the first failure without retrying or selectively rerunning a child.

Admission retains the exact command, cwd, approved environment, executable and
fixture bindings, launch PID/CPU, task/header/lifecycle stream, raw STACK_USER
summary, report, terminal receipt, and decoder artifacts. Every child must
pass the inherited correctness/report gates, have at least 100 marker samples,
positive marker-period share, zero summed `PERF_RECORD_LOST`, and zero terminal
`PERF_RECORD_LOST_SAMPLES`. The raw and terminal loss values remain separate
checks and are never added. `--per-thread` task metadata, same-PID/TID records,
absence of FORK/CLONE or descendant EXIT, and the recorded post-exit `/proc`
observations are fail-closed. The decoder and reducer revalidate immutable raw
and header files without reopening `/proc`.

All outputs use `/home/zhuhe/litchi-performance-0975` exclusively. The plan,
helper seal, reused 0974 v5 files, reducer-v4/adapter-v3/adapter-v1 lineage,
source/fixture/binary hashes, and this design are bound before GO. The result is
diagnostic evidence only; it makes no qualification, speed, or exact timed
region claim.
