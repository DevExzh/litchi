# 0787 cached-Part scheduling source review

This is an independent source review of the final archived candidate against
base `6b47617020dd08bca196cbf3ab0b87a69aefbcb1`. I read the candidate design,
implementation notes, model patch, complete before/after source copies, and
the focused test diff, together with ADRs 0001, 0002, 0003, 0005, 0006, 0010,
0011, 0024, 0030, 0031, and 0032. I did not edit the live production source,
run Cargo, compile, run native workloads, run a profiler, or change HEAD.

## Source result

The archived source has the intended narrow shape: one private cache-state
predicate, one private admitted-path serial helper, and focused tests. I found
no source-level correctness blocker in the scheduling implementation.

The archive receipt records archive-only preparation and no compiler, Cargo,
native, profiler, or commit activity. The separate root-owned
`compile-preflight-0` record remains preserved; this review does not treat it
as a runtime test result.

`PartCache::all_entries_ready` takes the existing state mutex and requires an
entry with neither an active flight nor a provisional publication. It only
performs map membership checks; it does not advance LRU state, change hit or
waiter counters, clone payloads, pin entries, reserve budget, or alter cache
retention. Poison recovery follows the existing private cache methods. A
duplicate `EntryId` is harmless, and a missing, pending, or flight-backed ID
rejects the hint.

The hint is reached only after the existing initial `read_parallel` fence, the
existing scheduler admission and `CpuTasks` charge, and the explicit
`ScopedWorkers` branch. The scheduler reservations remain held through the
hint path. Thus the candidate retains the existing preflight, output
admission, `Workers`/`IoConcurrency`/scheduler `Memory` and `Objects`
refusals, and caller-facility observability. The candidate does intentionally
skip post-admission private thread/channel/vector construction when the hint
is true; an allocation or `thread::Builder` failure that existed only during
that construction is consequently no longer observable on this path. This is
an implementation consequence of the optimization, rather than an admission
or budget bypass, and should remain disclosed in the final report.

`read_cached_serial` matches the existing wave structure. A one-wave request
does not add a fence relative to `read_one_wave`; a multi-wave request fences
before each wave as `read_reused_workers` does. Each request still goes through
`read_part_prepared`, so source freshness, cancellation, cache admission,
eviction, flights, ZIP checks, and payload reservations remain authoritative
after the hint. A stale hint can therefore turn into an ordinary hit, wait, or
cold load without returning stale bytes. The helper catches each request
unwind, maps it to `SourceBackedBatchWorkerPanic { ordinal }`, drains every
request in the current wave, selects the lowest ordinal error, and stops
before a later wave. Output remains in prepared input order. The caller-owned
`ScopedWorkers` route is checked first and is unchanged.

These boundaries satisfy the applicable ADRs: explicit execution-context
admission and caller scheduling remain in the low-level OPC owner (0001,
0002, 0031); source identity, cache invisibility, bounded resources and
measured evidence remain authoritative (0005, 0006, 0030); immutable ordered
handles and no new facade or runtime dependency are preserved (0003, 0010,
0011, 0024); and the existing byte cache remains the sole payload authority,
with no derived-value memo (0032).

## Coverage result

The archived tests cover the all-hit byte/source-read result, cumulative
`CpuTasks` and final release, explicit caller-facility wave/task dispatch, the
helper's completed/flight/pending/eviction membership and no-LRU/no-hit
mutation, and both primed panic shapes. The multi-wave test proves typed panic
translation, current-wave draining, and stopping before the later wave. The
one-wave test injects two panics and proves lowest-ordinal selection after the
second request still runs. The internal stale-hint test observes an eligible
entry, evicts it, observes the false hint, and verifies the authoritative
prepared read returns exact bytes. The primed source-change test changes the
source at the first cached request's freshness check and verifies the final
source error with no in-flight load or retained budget leak.

The focused tests and the full OPC test command are recorded as passing in
[quality-0/02.log](quality-0/02.log): 623 library tests, 19
`source_backed_batch` integration tests, and the remaining OPC targets all
report zero failures. The focused names in that record are
`cached_multiwave_batch_keeps_admission_and_skips_physical_reads`,
`cached_batch_still_invokes_an_explicit_caller_worker_facility`,
`cached_worker_panic_is_typed_and_drains_only_the_current_wave`,
`cached_one_wave_panic_selects_lowest_ordinal_and_drains_the_wave`, and
`cached_source_change_after_helper_entry_fence_returns_source_changed`, in
addition to the internal
`cache_ready_hint_rejects_flights_and_pending_without_touching_state` and
`stale_ready_hint_after_eviction_uses_authoritative_prepared_read` tests.

The existing
`batch_reports_low_ordinal_failure_and_does_not_start_a_later_wave` test is a
cold private-worker test and remains useful as differential coverage; it is
not being counted as cached-helper panic coverage. Existing cancellation,
refusal, and retention tests remain the differential gates for the unchanged
ordinary paths. A partial-cache/miss comparison is still a root-owned runtime
and replay check; it is not established by the all-hit and fresh fixtures in
this archive.

The direct stale-helper test's `cold_loads == 2` assertion is grounded in a
fresh unmanaged package: package opening and metadata lookup do not enter the
Part payload cache, the first explicit prepared read contributes one cold
load, `evict_oldest` removes that sole unpinned entry, and the direct helper
miss contributes exactly one more. This is a valid focused invariant; the
root-owned runnable suite executed it under the actual crate configuration and
recorded it as passing in the quality log above.

The added test imports are all consumed by the new source wrappers, panic
guards, thread-identity check, and caller-worker counter. The corrected
`Arc<dyn ReadAt>` coercions and refreshed receipt/model-patch hashes are
consistent in the archive; Clippy and compilation remain root-owned checks.

## Review disposition

**Suitable for root-owned quality gates and paired measurement.** The focused
archive coverage now exercises the requested helper panic, stale-hint, and
post-entry-fence source-change seams, and the runnable test/check records pass.
The receipt's current file hashes match the archived source copies and model
patch. The separate boundary checker and final root-owned quality seal must
still be recorded before claiming the whole gate set, while preserving exact
outputs, resource refusals, source counts, caller-facility behavior, and
fresh/partial-cache/RSS evidence. The intentional post-admission
allocation-failure scope remains part of the candidate's disclosed behavior;
this does not support a broader Office CRUD or whole-program scaling claim.
