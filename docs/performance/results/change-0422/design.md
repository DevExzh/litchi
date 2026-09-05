# Operation-region high water

The 0421 corrected process high-water snapshots can remain above every live
value in an operation because setup reached an earlier maximum. They cannot
answer the required operation-local memory question.

The new observer serializes each entire allocator callback, region begin,
finish/drop and counter snapshot with one mutex. Begin initializes an active
peak from the absolute entry live total. Successful growth updates raise that
peak. Finish captures all counters and the region maximum, then releases the
region under the same lock. Absolute process counters never reset. Overlapping
region acquisition remains unavailable. Poisoning invalidates measurements
without unwinding or allocating inside an allocator callback.

Only numeric state and synchronization occur under the observer guard. It must
never span System allocator calls, allocation, formatting, logging, user code,
assertions, worker joins, or benchmark work. The normal binary has no global
allocator wrapper. The instrumentation is not a production synchronization
change and is not used to claim latency or scaling.

The exact boundary is observer callback order after System returns, not physical
heap mutation time. Every process callback between begin and finish contributes,
including background threads. Operation workers must finish before the caller
ends the interval if their complete work is to be included. The maximum includes
pre-existing live bytes; it is neither the operation's owned allocation peak nor
RSS. Realloc accounting uses requested old/new sizes and excludes the allocator's
hidden copy overlap. A signed exit-minus-entry live total is net process change,
not retained ownership.

`serialized_region_peak_v3` identifies this observer contract. The new
`region_peak_live_bytes` field must be at least both endpoint live totals and
no greater than lifetime high water at finish. It can decrease across consecutive
operations. All numeric vectors disappear on overflow/unavailable samples.
V2 remains valid historical process-counter evidence but cannot be paired with
V3 under one policy. These captures establish a current baseline only.

A read-only independent design review recommended the mutex over separate
active/max atomics: without a synchronized boundary, initializing/resetting the
local maximum races a callback or finish. The mutex is a deliberately simple
measurement enabler. Per-callback serialization can change thread scheduling;
normal binary measurements and separate contention work remain necessary for
production performance decisions.

## ADR fit

| Contract | Application in this batch |
| --- | --- |
| ADR 0001/0003 | Correctness and immutable document APIs remain unchanged; observer failures do not publish partial numeric evidence. |
| ADR 0002/0010/0011/0024 | All code changes stay in the isolated benchmark tool and its validator; no production dependency or archive ownership changes. |
| ADR 0005 | Explicit diagnostic instrumentation, bounded scope, reproducible source/corpus/output identities, no allocator latency/scaling claim. |
| ADR 0006/0008 | Existing lifecycle preservation/refusal/output gates are retained; focused validation and remaining lint debt are stated explicitly. |
