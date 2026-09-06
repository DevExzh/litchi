# Range pacing inside the managed PPTX cross-copy lifecycle

The provider-lifecycle command constructs deterministic source/destination PPTX
fixtures, validates producer refusal gates, creates fresh caller budgets and
prepares a full-retaining bounded CountingSink outside the API clocks. The
adapter owns an explicit OwnedSource; production APIs see only caller ReadAt.
Open-source and open-destination, plan_cross_slide_copy, and consuming sequential
publication are timed separately. Diagnostics, source/cache/budget snapshots,
exact full-output comparison, readback gates, and final owner drops occur outside
those API intervals. Unlike a discard sink, CountingSink retains the entire
output, so this does not prove bounded total memory.

Each nonempty delegated read first requests the configured fixed delay, then
calls the underlying source with the configured maximum returned length. Each
successful nonempty return additionally requests ceil(returned_bytes * 1e9 /
transfer_bytes_per_second) nanoseconds of sleep. Fixed-delay and transfer sleeps
are separate calls; kernel scheduling and timer granularity can add substantial
oversleep. transfer_paced_calls and transfer_delay_ns record requested pacing,
not time measured sleeping, achieved bandwidth, network I/O or system calls.

The pacing model is per read, without a shared link queue or a global bandwidth
budget. Concurrency can exceed the nominal aggregate rate; this batch uses one
worker and serial public API calls. No physical remote, disk-cold or scaling
claim follows. Range counter and ceiling arithmetic must be checked; unavailable
owners have null counters, including the new pacing fields. Final caller memory,
object and depth reservations must return to zero. Operation allocator-region
attribution is unavailable for this custom journal; process RSS and managed
budget snapshots are separate, incomplete memory observations.
