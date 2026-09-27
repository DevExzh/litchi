# 0788 memory attribution design

The previous experiment rejected the exact cached-Part candidate because
small/primed/floor-0/width-4 whole-child peak RSS increased 5.70%, with all
six pairs adverse. This batch investigates that outcome and always restores
production. It is not a retry of the adoption decision.

The same-width floor-65536 control is mandatory: the candidate hint is
unreachable there. Width 4 uses reused private workers in the baseline;
width 32 uses one wave. The candidate runs cached waves on the calling
thread. Its binary was 19,720 bytes larger in 0787, but binary size alone
cannot explain resident pages or establish the cause of a 222 KiB RSS delta.

Ten native cases have six fresh paired blocks. Four diagnostic cases cover
fresh width 4, primed width 4, primed serial-floor width 4, and primed width 32.
One-sample/no-warmup and thirty-sample/three-warmup protocols separate early
state from repeated lifecycle behavior. Handshake-enabled and disabled
runs use the same diagnostic binary and balanced orders. Native original
tool timings and RSS are never pooled with diagnostic runs.

The child emits a fixed phase marker and waits for an ACK. The parent
verifies the actual child executable/PID, reads smaps, rollup, maps, status,
stat and task IDs, then acknowledges. Only first/last measured samples are
observed; no warmup marker is emitted. Snapshots occur after construction,
preload, operation-before-verification, batch drop and package/context drop,
plus global lifecycle boundaries. Protocol I/O itself can allocate/fault
pages; disabled controls and that limitation remain explicit. No SIGSTOP
or allocator configuration change is part of this initial experiment.

A file-backed clean-page increase supports a code/rodata/unwind residency
hypothesis; anonymous/private-dirty growth and allocation call stacks support
an allocator hypothesis. Neither classification alone proves a cause. Stack
or TLS conclusions require mapping and thread evidence. If only high-water
counters differ, classify the evidence as unresolved transient/accounting
behavior, not noise. Heaptrack is a separate intercepted-allocation profile;
its global peak is not operation allocation or native peak RSS. Merged
backtrace peaks are disabled because the tool warns they are inaccurate.

Primary documentation motivates checking smaps independently of scalable
RSS counters; it does not establish that 0787's result was spurious. Exact
URLs and narrow claims are retained in references.json. Root alone runs
all builds, native children and profilers serially.
