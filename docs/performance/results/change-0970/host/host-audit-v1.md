# Batch 0970 host and harness measurement audit

- **Audit:** `0970`
- **Generated:** `2026-10-08T02:08:32.506913+00:00`
- **Scope:** read-only host and retained-0969 evidence audit
- **Workload/build status:** no Cargo invocation, build, benchmark, or native capture was started; the two explicitly authorized harmless capability probes ran `/usr/bin/true` and a Python affinity child only
- **Decision status:** this document proposes a bounded diagnostic; it does not retry the 0969 qualification and makes no speedup, slowdown, or host-cause claim

## Scope and custody

The current checkout is `/home/zhuhe/code/litchi`, at HEAD
`68b17e83b0290da5ad9dd2669a6a466ec7c69dca`. At the audit time, the harness and
runtime paths inspected below had no working-tree diff. Existing changes in
other paths were left in place and untouched. The retained 0969 decision names
source revision `34fc85e23ccefcd8ad6ffd4d57d9e2671ef1a5a1`; that historical
revision label is not substituted for the current checkout HEAD.

Current inspected file hashes are:

| file | SHA-256 |
| --- | --- |
| `tools/perf-cfb-profile/src/lib.rs` | `b4893321ac9a9321a9055823a7f05ecb35e59319531913affc49342b8ccd73dc` |
| `tools/perf-cfb-profile/src/main.rs` | `13206a810887dfcb4897a265167edc8fa3706ae736a2785a27354c7617eb5ca9` |
| `tools/perf-cfb-profile/src/bin/alloc.rs` | `2542257a9ab68347d3e8133b5f65a2254cc1e439fef665e327e093470f365f5f` |
| `tools/perf-cfb-profile/Cargo.toml` | `41dee12e4eb7d1e3853a33a7743afdf6b5d6519d5e7456e61ce4978aa9177945` |
| `tools/perf-cfb-profile/Cargo.lock` | `343c5464ae1144cbb72284cbb3e9b52b5dff2b798d09462d3c0d6a74b317bac7` |
| `crates/litchi-doc/src/body_text.rs` | `0e8cb6351be3327322882ae4035230cd241aed0b6208cb8589e7d76692d3d786` |
| `crates/litchi-cfb/src/writer/core.rs` | `a5ead3dded13baa8bdaa350742bfe79185e39cc2f16f86e433701e939fb2b19c` |

## Current host snapshot

The topology, cgroup, and load observations below were taken on
`2026-10-08T02:05:48Z` unless a more specific time is shown.

| observation | result | implication/limit |
| --- | --- | --- |
| host | `ip-172-31-39-190`, Amazon EC2 `m8a.8xlarge`; Ubuntu `26.04.1 LTS`; kernel `7.0.0-1013-aws` (`#13-Ubuntu SMP PREEMPT Fri Sep 4 17:02:11 UTC 2026`) | Identifies this host only; it does not reconstruct 0969's conditions during each child. |
| CPU topology | x86-64 AMD EPYC 9R45, family 26 model 2 stepping 1; one socket, 32 cores, SMT 1; one NUMA node; L3 128 MiB (4 instances) | There are 32 distinct logical CPUs, `0-31`; no SMT sibling sharing is present in this snapshot. |
| CPU online/present/possible | all `0-31`; `nproc` 32 | The host currently exposes all 32 CPUs. |
| shell affinity | `/proc/self/status`: `Cpus_allowed_list: 0-31`, `Mems_allowed_list: 0`; one thread | A child that inherits this scope can migrate across all 32 CPUs unless its launcher pins it. This is the shell, not a historical 0969 child. |
| cgroup path | `0::/user.slice/user-1001.slice/user@1001.service/tmux-spawn-3a22b971-8fd0-4c28-9367-732dd8523549.scope` | This is the current shell's cgroup; 0969 did not preserve a child cgroup path in its quiescence receipt. |
| cgroup2 mount and scope | `/sys/fs/cgroup` is a cgroup2 mount (`rw,nosuid,nodev,noexec,relatime,nsdelegate`); the exact scope directory exists at `/sys/fs/cgroup/user.slice/user-1001.slice/user@1001.service/tmux-spawn-3a22b971-8fd0-4c28-9367-732dd8523549.scope` | This resolves the v2 path; the missing root `/sys/fs/cgroup/cpu.max` is expected because the root has no local file. The scope file is authoritative for this shell and its inherited children. |
| nested cgroup controllers | `cpu memory pids`; child has no local cpuset files; ancestor `/user.slice` has `cpuset.cpus.effective=0-31` | Effective current cpuset is broad. Lack of child cpuset files means inheritance, not an additional restriction. |
| nested CPU quota/weight | `cpu.max=max 100000`; `cpu.weight=100`; `cpu.weight.nice=0` | No current cgroup CPU quota is configured for this scope. |
| nested cgroup counters | at the one-second read interval, `usage_usec` increased by `800246` (`user +758995`, `system +41252`), while `nr_periods`, `nr_throttled`, and `throttled_usec` all increased by zero | No current-scope throttling was observed in this one-second diagnostic interval, which had no benchmark child. Cumulative counters cannot prove that 0969 was never throttled. |
| memory/pids | `memory.max=max`; `pids.max=141220` | No memory CPU-like quota is visible; this audit did not exercise memory or PID limits. |
| load at snapshot | `/proc/loadavg`: `0.52 0.94 1.85`; one runnable process at read time | The machine was lightly loaded by ordinary standards, but this is a point-in-time observation. |
| CPU pressure | `/proc/pressure/cpu`: `some avg10/60/300=0.00/0.00/0.00`; `full avg10/60/300=0.00/0.00/0.00` | No recent CPU-pressure average was visible at this snapshot; no historical 0969 pressure series exists. |
| one-second `vmstat` interval | second sample: `us=0 sy=0 id=100 wa=0 st=0` | No steal or wait was observed in this one-second interval. It does not establish frequency or steal behavior across 0969's four-minute A/A wave. |
| frequency driver/policy | `/sys/devices/system/cpu/cpu0/cpufreq` is absent; `cpupower frequency-info` reports “no or unknown cpufreq driver”, no policy/governor/current frequency | Frequency was not fixed or observed by this audit. There is no basis for treating wall-clock samples as constant-MHz measurements. |
| idle/thermal visibility | cpuidle driver `acpi_idle`, governor `menu`; no readable `amd_pstate`, `thermal_zone*`, or `cpu*/thermal_throttle/*` files | Thermal throttling and actual hardware frequency cannot be ruled in or out from this interface. |
| perf availability | `/usr/bin/perf`, version `7.0.14`; `kernel.perf_event_paranoid=1`, `kptr_restrict=1`, `nmi_watchdog=0`; current `CapEff=0` | The authorized `/usr/bin/true` probe enabled the fixed software and hardware event list successfully. `taskset` 2.41.3, `turbostat`, `cpupower`, `numactl`, and `chrt` are installed. |
| build/workload processes | no Cargo/rustc/native probe process was present when the final targeted process check ran | The audit itself did not start one; this is not a historical assertion about 0969. |

The current snapshot is consistent with an unrestricted, broad 32-CPU cgroup
and no current throttling, but it cannot certify the host state during the
retained A/A wave. In particular, the absence of a cpufreq driver and thermal
files means that frequency behavior is an unresolved variable.

## What 0969 actually recorded

The authoritative retained plan is
`/home/zhuhe/litchi-performance-0969/plans/plan-0969-aa-v6.json` (SHA-256
`e853030180d7464c06cf1ed5d525957ce2d50514b8337a1b168392374d00b5df`). Its
baseline-only A/A schedule was:

- 24 cells: `docnohf` and `docfloat`, edits `1/2/8/32`, retention
  `zero/default/release`.
- 12 same-binary pairs per cell, with six `AB` and six `BA` pairs.
- Each child: five warmups and 80 native samples, serially launched; 576 child
  packets in total.
- The executable was the retained 0967 baseline native binary
  `/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile`, SHA-256
  `0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb`.
- The runner requires no Cargo and declares `serial=true`, `retry=false`, and a
  900-second child timeout.

The retained runner implementation is
`/home/zhuhe/litchi-performance-0969/tools/run_0969_v4.py` (bound in the plan
as SHA-256
`c7315d5c5311c5643aebca16e8cf2b931ee1b74663eb5a8a31e8185f504ddd20`). In
`launch_packet`, it builds a command of the form:

```text
/usr/bin/time -f %M -o <packet>.rss <binary> --case <case> --input <fixture> \
  --edits <edits> --retention <retention> --warmup 5 --samples 80
```

and calls `subprocess.Popen(..., start_new_session=True)`. There is no
`taskset`, `sched_setaffinity`, CPU-list, governor, cgroup, or priority setting
in this launcher. `start_new_session` changes the process session; it does not
pin a CPU. The runner records command, binary/fixture/source hashes, packet
creation time, child terminal status, output paths, RSS from `/usr/bin/time`,
and an excluded wall duration. It does not record the child PID, effective
CPU affinity, CPU migrations, per-process context switches, cgroup path or
`cpu.stat` deltas, frequency, thermal state, steal, task-clock, or hardware
cycles.

The 576 packet files span `2026-10-08T01:43:05.967633Z` through
`2026-10-08T01:47:55.364881Z` (about 289.4 seconds). Their excluded wall
Durations have a median of about 614 ms, minimum 114 ms, and maximum 1.416 s.
Those are runner durations, not the harness's reported native metric. Packet
objects have no `pid` field. The retained quiescence receipt at
`2026-10-08T01:42:08Z` records only CPU count 32, node/kernel identity,
load averages `[0.935546875, 2.3408203125, 2.39697265625]`, and an empty
`blocking_processes` list. It does not record cpuset, cgroup quota or counter
state, frequency/governor, thermal state, perf permission, affinity, or
scheduler counters. Root GO ended its 57.85-second cooldown at `01:43:05Z`.

The A/A result is
`/home/zhuhe/litchi-performance-0969/capture/aa-v4/aa-readiness-v4.json`
(SHA-256
`6bb110145d0583be3ff12c3c6e1ada6611be6b5396f55f24698bb0358bb51efe`). It is
`readiness_fail` against upper-CI limit `1.05`, with 12 pairs per cell and
10,000 bootstrap resamples, so the runner stopped before candidate formal
lanes. For example, `docfloat/e1/default` has A/B p50 ratios
`[1.0149, 0.9138, 1.1884, 1.0594, 1.2230, 0.8775]` with symmetric upper
`1.2057`, and B/A p50 ratios
`[0.8891, 0.7906, 0.9923, 0.9789, 1.1784, 0.8063]` with symmetric upper
`1.2525`. These establish same-binary noise under the retained schedule; they
do not identify migration, frequency, cgroup, background scheduling, or hash
seed as the cause. The retained decision therefore correctly says
`uninterpretable`, `formal_not_admitted`, and `not_measured`.

## Harness boundary and source-path inspection

The current `tools/perf-cfb-profile` source makes the following boundaries
explicit:

1. `run` reads and verifies the fixture, then completes `prepare` before any
   warmup or measured sample (`src/lib.rs:937-945`). `prepare` runs the three
   retention routes and all correctness/oracle controls (`src/lib.rs:686-796`).
2. The protocol overrides the README's native defaults (3 warmups/30 samples)
   with 5 warmups/80 samples through the CLI. The loop is sequential: warmups
   at `src/lib.rs:950-953`, measured samples at `src/lib.rs:955-960`.
3. The native clock starts at `timed_case` line 354, calls the unique
   `#[inline(never)]` `profiled_measured_case` wrapper at line 355, and stops at
   line 356. The wrapper calls `run_case` (`src/lib.rs:332-345`).
4. `run_case` (`src/lib.rs:290-330`) includes the owned source clone,
   `Snapshot::open_bounded`, one `Edit` with the requested repeated paragraph
   replacements, optional retained-render releases, `commit`, and the output
   `Vec` copy. Per-sample correctness checks, hashes, public-reader reopen, and
   JSON serialization occur after the measured result returns; they are outside
   the native elapsed interval.
5. The allocator binary uses the same operation-local wrapper under
   `alloc_metrics::region` and intentionally reports no timing claim
   (`src/bin/alloc.rs:43-49`, `src/lib.rs:388-410`).

There is no explicit `rand`, `Rng`, `thread_rng`, or random input generation in
the profile harness. The baseline path nevertheless constructs and uses normal
Rust hash tables. `litchi-cfb::writer::OleWriter` stores `storages` in a
`HashSet` and CLSID/metadata tables in `HashMap`s (`crates/litchi-cfb/src/writer/core.rs:320-341`), and the writer performs hash-table membership/insertion work during
serialization and path handling. The writer converts storage paths to a
`BTreeSet` before effective layout planning (`core.rs:525-539`) and imposes a
canonical storage order (`core.rs:2294-2307`), so output bytes are intended to
be deterministic. That canonicalization does not make the cost of the default
process-seeded hash operations identical between fresh child processes. The
repository's determinism test documents that each `OleWriter::new()` obtains a
per-process/thread-seeded `RandomState` (`crates/litchi-cfb/tests/writer_storage_determinism.rs:9-19`).

The timed package path reaches `RevisionEditor::finish_unchecked`, which either
moves a retained render or calls `Package::finish`
(`crates/litchi-doc/src/tracked_revision/package.rs:2654-2659`). Therefore fresh
0969 children can differ because of process-seeded hash state even though their
final bytes pass the deterministic-output gates. This is a plausible
process-specific component, not a demonstrated explanation of the A/A failure.

## Authorized capability probes

Two harmless capability probes were run after the host snapshot, with no native
benchmark, Cargo process, or production process involved. The exact command,
exit status, and raw stderr for the fixed telemetry mode are in
`/home/zhuhe/litchi-performance-0970/telemetry-probe-v1.json` (SHA-256
`8cb97d377c6162c57e1ae092c7c8e33f9b2b23c38aa2d44f188d3831e20614dc`). The
command was:

```text
/usr/bin/perf stat -x, -e task-clock,cycles,instructions,context-switches,cpu-migrations,page-faults -- /usr/bin/true
```

It exited 0 and enabled all six requested events. The CSV stderr values were
`27.51 msec task-clock`, `961846 cycles`, `741538 instructions`, `0
context-switches`, `0 cpu-migrations`, and `59 page-faults`. This freezes one
telemetry mode for any later diagnostic; a later child must not silently fall
back to a different event set if one event is unavailable.

The child-only affinity probe is in
`/home/zhuhe/litchi-performance-0970/affinity-probe-v1.json` (SHA-256
`0f2e97080d2ba9898a3c0c74660f3280a4cbb66e4546997fa93749d476ce6a17`). It used
Python `os.sched_setaffinity` in a child, selected the highest currently allowed
CPU (31), observed `[31]`, then restored the child's original `0-31` mask before
exit. The parent mask was `0-31` both before and after. No global affinity or
host setting changed.

## Evidence gap by possible cause

| possible contributor | retained 0969 evidence | what can be said now | missing discriminator |
| --- | --- | --- | --- |
| CPU migration | launcher has no affinity setting; no child PID or migration counter | migration was allowed by the command shape, but its actual rate is unknown | child affinity plus `cpu-migrations` or `se.nr_migrations` per run |
| cgroup throttling | only `cargo_processes=0` and a point-in-time load snapshot were retained | current scope has `cpu.max=max` and zero current throttle delta; 0969 history is unknown | cgroup path and before/after `cpu.stat` for every wave, including ancestors |
| preemption/background scheduling | no per-process CPU time, context-switch, runnable, or scheduler series | current CPU pressure is zero at the snapshot; that does not cover 0969 | `task-clock`, context switches, migrations, `/proc/<pid>/sched`, and interval `/proc/stat` |
| frequency/thermal/steal | no frequency, APERF/MPERF, thermal, or steal evidence | current cpufreq/thermal interfaces are unavailable; one current `vmstat` interval had zero steal | `cycles/ref-cycles` or turbostat/APerf-MPerf evidence plus interval steal/thermal counters |
| process/hash-seed variation | fresh children; no explicit RNG; default `HashMap`/`HashSet` in the writer path | process-specific variation is plausible while bytes remain deterministic | compare same-process 80-sample spread with fresh-process medians under fixed affinity |
| I/O/cache effects | RSS and excluded wall duration only | no I/O pressure was assigned to any packet | keep the existing elapsed metric, but capture CPU/task counters and page/cache context around the child |

## Minimal discriminating diagnostic (proposal only)

Do not rerun formal qualification while the A/A gate is failing. After a new
root-approved A/A pass, the smallest useful diagnostic is a **single-cell,
baseline-only** comparison using the already retained native binary and the
`docfloat/e1/default` fixture, which was visibly noisy in the retained A/A
receipt. It should remain exploratory and must not be merged into the formal
qualification sample.

Use 3 serial `AB` pairs and 3 serial `BA` pairs in each of two launcher
conditions, for 12 fresh children total:

- inherited affinity (the current unpinned command shape), and
- `taskset -c 31` around the exact same child command.

Keep the existing `--warmup 5 --samples 80` settings and the same fixture,
retained binary, and output validation. Wrap each child with the fixed, already-probed event list:

```text
perf stat -x, -e task-clock,cycles,instructions,context-switches,cpu-migrations,page-faults -- <child>
```

The pinned form places `taskset -c 31` before `perf stat`; the unpinned form
omits it. This diagnostic's profiler overhead makes its elapsed values
non-qualification evidence. Before and after each child, record the exact
command, PID, `/proc/<pid>/status` affinity, `/proc/<pid>/sched` migration and
context-switch fields when available, cgroup path, cgroup `cpu.stat`,
`/proc/loadavg`, `/proc/pressure/cpu`, and `/proc/stat`. The authorized capability probe above established this exact event set on the
current host; preserve the same event list and CSV format for each child. No
per-child fallback event set is allowed.

Interpretation should be limited to patterns:

- Lower spread together with zero or near-zero migrations in the pinned lane
  supports migration/scheduling as a contributor.
- If pinned runs have no migrations but fresh-process medians still differ while
  each child's 80-sample within-process distribution is comparatively tight,
  process-specific state such as default hash seeds becomes a leading
  explanation. It remains an attribution, not a causal proof.
- A positive delta in `nr_throttled` or `throttled_usec` identifies cgroup
  throttling for that interval; the current `cpu.max=max` makes this less
  likely but does not substitute for a run-time delta.
- Changes in cycles/ref-cycles or steal/frequency observations with stable
  migrations and context-switch counts point toward frequency/virtualization
  variation.
- If none of these patterns separates the lanes, retain the A/A failure as
  unresolved environmental/process noise and do not infer a host cause.

A future formal receipt should add affinity, child PID, cgroup path and quota,
start/end cgroup counters, interval load/pressure, frequency/thermal
visibility, migration/context-switch evidence, and perf preflight outcome to
make a later A/A result diagnosable. No production source, harness code, or
retained evidence was changed for this audit.
