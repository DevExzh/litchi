# 0498: bounded source-backed multi-Part reads

The before revision is `56e8ddbca`. This work follows the measured independent
large-Part and delayed-source opportunity identified in 0497's
`next-implementation.md`. It adds an explicit production batch read operation;
it does not reinterpret harness worker teams as production scaling evidence.
The serial selected-Part loop is captured before production edits. The retained
before executable can supply additional matched controls after implementation.

The API must retain returned collection memory charges for the lifetime of its
owner, preserve full typed errors, use source-bound immutable metadata, and
route each payload through the existing cache and source checks. Scheduling
must respect caller worker/task/byte limits and a minimum-work threshold.
Bounded deterministic waves are preferred if they provide useful measured
scaling without the draft's shared queue and condition-variable complexity.
All admitted workers must join on success or failure. No partial collection
is returned. Source freshness and cancellation remain final fences.

The earlier draft's budget rollback language is corrected: cumulative Work
and InputBytes represent actual work and cannot be rolled back to pretend a
parallel failure performed only serial work. Successful clean cache entries
may remain after a failed batch. Temporary scheduling charges must release,
while returned payload and collection charges remain with their owners.
No claim of schedule-independent resource exhaustion or identical provider
side effects is made. Error selection must be deterministic over input order
within the admitted work, preserving original typed error fields.

## Measurement and checks

Use deterministic few-large and many-small packages, with fresh selective
payload loads after package setup. Cover owned, warm positional file,
instrumented short-read, and fixed-delay caller sources. Compare the serial
control with the production batch at 1, 2, 4, and 8 workers. Keep source calls,
accepted bytes, overlap maxima, budget/cache release, and semantic byte oracles
alongside latency samples. Capture three warmups and thirty measured samples
in two repeats for accepted rows, and report every greater-than-five-percent
latency/RSS flag. Report speedup, efficiency, and an Amdahl estimate only for
fixed-work rows with actual measured scaling. Setup and whole-child profiling
must not be mislabeled operation-local evidence.

Compile with `RUSTUP_TOOLCHAIN=1.98.1`, `CARGO_INCREMENTAL=0`,
`CARGO_PROFILE_RELEASE_DEBUG=0`, and `CARGO_BUILD_JOBS=2`, using the owned
`/home/zhuhe/.cache/litchi-goal-0498/target`. Disk is shared and constrained;
do not remove unrelated targets or protected work. Preserve the original
ODG/iWork edits. Keep the harness separate from production dependencies.

Required focused checks include order/duplicates, gated overlap and task/byte
bounds, sequential fallback, typed missing/corrupt errors, cancellation,
source mutation, budget/flight/owner release, and source short reads. Run OPC
feature/default tests, applicable downstream checks, formatting, warning-denied
lint and rustdoc, and crate-boundary checks. Broader workspace and native
producer requirements remain part of the full goal; this batch alone cannot
close them. No SIMD or new global executor is introduced.
