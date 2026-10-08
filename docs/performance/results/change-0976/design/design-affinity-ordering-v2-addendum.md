# 0976 affinity-ordering v2 corrective addendum

This addendum corrects one overbroad statement in design-affinity-ordering-v1.md. The v1 file is preserved unchanged. This is a read-only source and evidence correction; no native workload, build, replay, probe, or repository edit was run for this addendum.

Authored at 2026-10-08T06:10:00Z UTC.

## Correction

The v1 sentence that cmd_record unconditionally creates a reader pthread before the workload is incorrect. The installed perf may create parallel mmap reader pthreads only when its parallel-streaming thread-spec option is enabled. With the command used by the v6 control, only one perf TID was expected and observed.

The installed perf is /usr/bin/perf version 7.0.14, SHA-256 174229fb48b8557aafa821dd828d883dce47c35c31067e1f31e4a01a91c6c98b.

In the installed ELF disassembly of cmd_record:

- At 0x8188d, the code tests record plus offset 0x24c; the conditional branch at 0x8189b reaches 0x82de1 when that value is zero.
- The nonzero path at 0x81908 initializes pthread attributes, calls pthread_attr_setaffinity_np at 0x8198e, calls pthread_create with record__thread at 0x819a7, and waits for the new thread's readiness pipe.
- The zero path bypasses that pthread creation loop and continues through the normal workload-start path.

The exact v7.0 perf source cross-check identifies this guard as record__threads_enabled(rec), which returns rec->opts.threads_spec. record__start_threads() returns immediately when that guard is false; only the parallel-streaming path creates t=1 through nr_threads-1 reader threads and then adjusts the main reader thread affinity. Source snapshot URL: https://raw.githubusercontent.com/torvalds/linux/v7.0/tools/perf/builtin-record.c. Retrieved source SHA-256: a8e8f945db9111c6c379751c55e50e842321108b214353f5d64ff270f2b87af8, 121429 bytes.

The installed help distinguishes --per-thread, which selects per-thread mmaps, from --threads[=<spec>], which writes trace data with parallel threads. The v6 command included --per-thread but did not include --threads.

## Observed v6 control

The v6 command receipt is /home/zhuhe/litchi-performance-0976/control/affinity-control-v6/command-receipt.json, SHA-256 c8c0b1babcc11b04ce401573f420975071ec3146047e75adaa9c660e3976a000. Its argv contains taskset CPU31, perf record, --per-thread, instructions:u, period 1000000, -m 16M, DWARF 32768, and the affinity bridge; it contains neither --threads nor --control nor -D -1.

The bridge receipt is /home/zhuhe/litchi-performance-0976/control/affinity-control-v6/bridge-receipt.json, SHA-256 7fc8af88bbc5f99f4bcebf9ed9ca89e99917255abb071264b69a55c69e1a821d. It observed exactly one perf task at every recorded before, after, final, and periodic check:

- perf PID/TID 3786017, starttime 30518336, initially [31], then [30].
- target PID/TID 3786018, starttime 30518338, [31] throughout.
- No additional perf TID was observed during the 0.25-second control target.

The target receipt is /home/zhuhe/litchi-performance-0976/control/affinity-control-v6/target-receipt.json, SHA-256 0dcac8748bb1e89fa6a3f012b65fa5d028ed9b85b7a286fb7b9c89bfb573abc5. It reports control_only=true and native_workload_started=false. The terminal receipt is /home/zhuhe/litchi-performance-0976/control/affinity-control-v6/terminal-receipt.json, SHA-256 14c93c60bbc165188c7726d5e4cf52ca98a0f314b9a1007427f93f31d650b747, returncode 0 and no_retry=true.

This control proves the single-TID handoff for this exact v6 command and host. It does not prove that a future command with --threads or another perf mode will remain single-threaded, nor does it prove the proposed disabled-event/control-FIFO ordering: v6 did not use -D -1 or --control.

## Updated design interpretation

The per-TID handoff remains the safe general procedure, but its wording is now conditional:

1. Enumerate the actual perf task set using /proc/PERF_PID/task after the perf command is verified by PID, starttime, executable, and UID.
2. If the set contains only the leader, move and verify that TID to CPU30.
3. If parallel streaming has created additional reader TIDs, move and verify every observed TID to CPU30. A leader-only move remains invalid in that case.
4. After enabling or beginning the capture, fail closed if a new perf TID appears, a known TID changes affinity, or any task identity/starttime changes. The capture must record the task-set snapshots rather than assume a fixed reader count.
5. Keep the target bridge, launcher, and native process on CPU31 and preserve the existing period, DWARF, ring-size, loss, marker, and scope gates.

Replace the v1 proof obligation “no reader TID is created after the verified handoff” with “no unverified perf TID may appear after the verified handoff; any newly created TID must be observed and moved before enabled sampling, otherwise abort.” The source shows that the parallel reader set is conditional; the control receipt shows only one TID in the selected non-parallel mode.

