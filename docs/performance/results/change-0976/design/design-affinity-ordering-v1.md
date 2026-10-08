# 0976 affinity-ordering design

Status: prospective design only. No native workload, build, capture, probe, or repository edit was run for this note. Authored at 2026-10-08T05:50:21Z UTC.

The objective is to keep the target on CPU31 while moving perf's active recording work to CPU30 before events are enabled. The existing 0975 design exposed a startup hazard: putting outer perf on CPU30 and relying on an inner taskset CPU31 can record the short launcher handoff on CPU30. This design preserves the 0973 gates and treats the recorder's initial CPU31 interval as an explicit handoff interval, with events disabled during the move.

## Evidence and bindings

The accepted ADR and GOAL revalidation is external at /home/zhuhe/litchi-performance-0976/receipts/adr-goal-revalidation-0976-v1.json, SHA-256 364dc685aa3c1f91471602a89fddcc48f17e92aec2f6824461608ad2f0c4b3e2.

The installed profiler is /usr/bin/perf, version 7.0.14, SHA-256 174229fb48b8557aafa821dd828d883dce47c35c31067e1f31e4a01a91c6c98b. Its local perf-record manual is /usr/share/man/man1/perf-record.1.gz, SHA-256 82b366662e48210cd62ff60e556f0fd673217a08768030a7d6eb3b6dafa5b73d.

The 0975 prior feasibility note is /home/zhuhe/litchi-performance-0975/analysis/next-capture-feasibility-v3.md, SHA-256 a49309d77ef3e051b2166fffba42d3a679ffe3b6af0f8047b37f117de2b4d586. The retained runner named there is /home/zhuhe/litchi-performance-0973/tools-v2/run_0973_v2.py, SHA-256 f4f7af4ec8b8bb1818f3ce25dec046c8b8ad7c23d0a2611c76e24bf10b6aef5d.

Local perf help documents --per-thread as per-thread mmaps with inheritance disabled, -D -1 as starting with events disabled, and --control=fifo or fd as an enable/disable control channel with optional acknowledgement. The installed perf ELF is unstripped and contains symbols and DWARF. Disassembly shows cmd_record creates the record reader pthread before it starts the workload; record__thread reads the mmap and reports its CPU. Therefore the affinity unit is every perf TID, not only the perf leader PID.

Linux sched_setaffinity documentation says affinity is per-thread: getpid() addresses only the main thread, each thread has an independently adjustable mask, and a child inherits its mask across exec. This is the basis for rejecting a leader-only PID move.

## Recommended topology

Use an external same-PID bridge as the command launched by perf. Start outer perf on CPU31, with the target initially inheriting CPU31. The bridge verifies its own mask is exactly [31], identifies its parent using PID, /proc starttime, /proc/exe, and comm, and rejects any parent that is not the expected /usr/bin/perf instance.

Before enabling events, enumerate /proc/PERF_PID/task and bind every verified perf TID, including the leader and the reader pthread, to CPU30 with sched_setaffinity. Re-enumerate and reread each mask. Require all expected TIDs to be present with the same starttimes and exactly [30]; reject EPERM, ESRCH, EINVAL, a missing TID, a new TID, or any mask mismatch. Moving only the leader is invalid because the reader pthread has its own affinity.

The bridge remains on CPU31 and then execs the retained launcher. The exec preserves the target PID and target TID while changing comm/exe. The launcher must recheck [31] immediately before the native exec, and the native process must retain [31]. Do not use an inner taskset that starts under CPU30. Omit --affinity=cpu for this design: installed help describes it as affinity for the trace reading thread, but the observed per-thread CPU_MAP is -1 and the local evidence does not establish a dedicated CPU30 placement or reset behavior for this mode.

Use the retained native settings unchanged: instructions:u, period 1,000,000, --per-thread, -m 16M, DWARF stack size 32768, serial children, and the existing 5 warmups plus 80 measured iterations. Preserve the zero raw PERF_RECORD_LOST gate, zero terminal PERF_RECORD_LOST_SAMPLES gate, at least 100 marker samples, positive marker-period coverage, exact target PID/TID scope, and no unexpected fork/clone/descendant activity.

The normal command should use -D -1 and a control FIFO. The intended sequence is: start perf with events disabled; launch the bridge; verify the target mask and the complete perf TID set; move and verify all perf TIDs on CPU30; send enable through the control FIFO; require the acknowledgement; only then exec the retained launcher/native target. This makes enabled samples begin after the affinity handoff. Any pre-enable record must be absent or have CPU31; any enabled sample with CPU other than the target's verified CPU31 fails the scope gate. The recorder may execute briefly on CPU31 before the move, which is recorded as startup handoff overhead rather than silently treated as steady-state recorder placement.

The bridge/controller must bind and retain the exact perf command, bridge and launcher hashes, environment, cwd, control FIFO paths, enable and acknowledgement transcript, perf PID/TID starttimes, before/after affinity masks, target PID/TID/comm/exe transitions, perf header, raw CPU set, fork/clone/exit stream, and both loss counters.

## Proof obligations before an Office capture

A fresh harmless external control must exercise the exact bridge and control ordering without the Office workload. It should use the same outer CPU31, perf version, --per-thread, -D -1, control FIFO, event, period, DWARF size, and -m 16M settings, but a bounded stdlib-only target that stays on CPU31. It must demonstrate all of the following:

1. The bridge can identify the correct same-user perf parent by PID plus starttime and executable identity.
2. Every perf TID can be moved to CPU30 and remains there through the enabled interval; a leader-only change is insufficient.
3. The target bridge and its exec chain stay on CPU31, with one target PID and no unaccounted descendants.
4. The control FIFO enable and acknowledgement establish the exact event-enabled boundary, and no sample is admitted from the disabled handoff.
5. The raw sample CPU field, PERF_RECORD_LOST records, terminal PERF_RECORD_LOST_SAMPLES, and task lifecycle records satisfy the future gates.
6. No reader TID is created after the verified handoff, or such a TID is detected and causes fail-closed abort.

The control must be preserved as a capability result, not used as Office evidence. If a TID cannot be moved, the affinity changes during the run, the control acknowledgement is missing, or any loss/scope gate fails, stop without a native retry.

## Limits

The local manual and disassembly establish per-thread affinity, reader-thread creation, and control syntax. They do not prove that the installed perf never creates a later reader TID, that its mmap reader does not reset affinity in this per-thread mode, or that same-user TID movement succeeds under the current host policy. The harmless control is therefore mandatory. The design also does not claim that recorder contention during the short disabled handoff has no timing effect; it only ensures that enabled target samples and steady-state target execution remain CPU31 while verified perf reader TIDs are CPU30.

