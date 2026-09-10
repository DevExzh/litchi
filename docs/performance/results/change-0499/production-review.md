# 0499 production review

Read-only review of the current `source_backed` loader guard, the ordered-batch
worker reuse patch, and `source_backed_batch_reuse.rs` at the working-tree
revision. This review does not replace the focused build, lint, or benchmark
gates.

## Findings

The loader guard is correctly scoped to `CacheAccess::Loader`. It retains the
existing `Arc<LoadFlight>`, calls `complete_failure_with_observer` with the
no-op observer only during unwind, and is disarmed after the ordinary result
path. Normal success, failed loads, uncached publication, and pending
publication cleanup therefore do not acquire an extra cache lock. An unwind
before flight completion releases the flight reservations, removes matching
pending state, and wakes waiters.

The reused path creates one operation-local scoped worker per admitted worker
and one bounded command/reply channel per lane. It sends one request per lane,
receives every reply before admitting the next wave, and preserves input-order
result/error selection. A task panic becomes the existing typed worker-panic
error; that worker exits, the current wave is drained, no later wave is
admitted, and all remaining workers are sent shutdown and joined. The
single-wave path intentionally keeps the original direct scoped-spawn route,
so it creates no channel pool when all requests fit in one wave.

Scheduler accounting now covers the exact reserved vector element types for
command senders, reply receivers, worker handles, and active ordinals. It also
charges one bounded command slot, one bounded reply slot, channel endpoints,
and a named 4096-byte-per-worker channel-control allowance. The latter is an
explicit conservative admission envelope for `std::sync::mpsc` internals; it
is not a portable exact allocator-size or OS-stack measurement. The memory
reservation should therefore be described as bounded scheduling admission,
not as a proof of exact resident memory. The four channel endpoint object units
per worker match the two channels' endpoints; payload bytes remain owned by
the cache/flight reservations.

After worker startup, `run_reused_waves` returns a structured result and the
caller always performs shutdown followed by explicit joins, including fence,
wave-planning, dispatch, reply-disconnect, task-error, and cancellation exits.
The outer source and cancellation fence remains the final precedence decision.

## Evidence

The reuse test's `ThreadId` evidence is valid for this claim: the Rust standard
library documents that `ThreadId` values are unique for every thread created
during a process lifetime and are not reused after termination. Therefore its
bounded ID set distinguishes the operation-local worker pool from the former
per-wave spawning route, and the disjoint sets distinguish the two separate
operations. The IDs are Rust identities rather than OS thread IDs; the
independent clone trace supplies the OS-level creation count.

No production correctness blocker was found in this read-only pass. The
channel-control and OS-stack charges remain deliberately conservative and
should stay scoped as such in the performance report.
