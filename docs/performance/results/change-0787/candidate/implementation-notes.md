# 0787 cached prepared-read candidate

This is an archive-only candidate against base commit `6b47617020`. The archive
author did not edit live crates or `HEAD`; root-owned preflight and quality
work are recorded separately. `candidate/model.patch` is the
baseline-relative patch, `candidate/before/` contains complete pre-edit copies,
and `candidate/files/` contains complete replacement files for every path in
the patch. No Cargo command, compiler, native producer, profiler, benchmark,
or commit was run while preparing this archive.

## Candidate shape

The candidate keeps the existing `PreparedBatch` preflight, source-version
fences, output admission, scheduler admission, `CpuTasks` charge, and ordinary
serial fallback unchanged. It does not change the `parallel` eligibility
predicate. The hint is evaluated only inside the already admitted private
parallel path, after its existing fence and after the explicit caller-worker
branch. The `SchedulerAdmission` remains held while the hint path runs, so
`Workers`, `IoConcurrency`, scheduler `Memory`, scheduler `Objects`, and the
conservative reservation envelope retain their current refusal and release
behavior.

`PartCache::all_entries_ready` takes the existing cache mutex and checks each
prepared `EntryId` against `flights`, `pending`, and `entries`. It is a
scheduling hint only: it does not call cache admission, advance the LRU clock,
increment hit or waiter counters, reserve memory or objects, clone a payload,
or create a pin/query/memo. A flight or provisional publication conservatively
rejects the hint even when an entry is also present. Eviction after the hint is
safe because `read_part_prepared` remains the authority for cache admission,
source freshness, cancellation, ZIP validation, and reservation ownership.

For a true hint, `read_cached_serial` uses the existing `wave_end` boundaries.
A one-wave request runs in input order. A multi-wave request takes the same
fence before every wave as the reused-worker path. Each request is wrapped in
the same panic boundary used by private parallel workers and maps an unwind to
`SourceBackedBatchWorkerPanic { ordinal }`; all requests in the current wave
run after an error, the lowest ordinal error wins, and later waves stop. The
ordinary prepared read still performs its per-request source and execution
fences, so a stale hint can become a normal cold read without changing the
read contract.

The explicit `ScopedWorkers` route is checked before the hint and still calls
the caller facility for every wave and request. This preserves the observable
facility callback contract. The private all-ready route may avoid private
thread, channel, and worker-handle construction after admission; consequently
an internal post-admission thread-creation or scheduler-vector allocation
failure that the old worker path could expose is not reachable on this route.
Those are the only intentional refusal differences. Admission-time resource
refusals, preflight refusals, source-version errors, cancellation errors,
`CpuTasks` exhaustion, typed zero-worker/zero-I/O refusals, output ownership,
cache eviction/pinning, and final source fences remain in their existing
locations.

## Focused coverage

The archived integration tests add:

- `cached_multiwave_batch_keeps_admission_and_skips_physical_reads`, which
  warms four Parts, proves the two-wave and one-wave all-hit calls return exact
  bytes with zero physical reads and only the caller thread observing source
  versions, checks cumulative `CpuTasks`, and checks worker/I/O and final
  memory/object release;
- `cached_batch_still_invokes_an_explicit_caller_worker_facility`, which warms
  the same request set with a caller-provided facility and proves the warm
  request still dispatches its two caller waves and four callbacks;
- `cached_worker_panic_is_typed_and_drains_only_the_current_wave`, which
  primes all entries, injects a source-version panic at ordinal zero's cached
  fence, proves the panic is translated, proves the other request in the first
  wave still enters the cache, and proves the later wave is not started;
- `cached_one_wave_panic_selects_lowest_ordinal_and_drains_the_wave`, which
  primes two entries, injects panics at both cached freshness checks, and
  proves the one-wave helper drains both requests while selecting ordinal zero;
- `cached_source_change_after_helper_entry_fence_returns_source_changed`,
  which primes two entries, changes the source at the first cached request's
  freshness check, and proves the source-version error and reservations remain
  clean;
- `stale_ready_hint_after_eviction_uses_authoritative_prepared_read`, an
  internal batch test that observes a true hint, evicts that entry, observes a
  false hint, and proves the helper's ordinary prepared read returns exact
  bytes and performs a second cold load;
- `cache_ready_hint_rejects_flights_and_pending_without_touching_state`, an
  internal cache test that proves a completed entry is eligible, an active
  flight or pending publication is not, a final eviction removes eligibility,
  and the hint does not advance the LRU clock or hit counter.

Existing batch coverage remains the differential gate for partial-cache/miss
parallel overlap, source-version mutation, cancellation and worker joining,
`CpuTasks`/I/O/resource refusals, low-ordinal cold-path errors, and cache
retention/release. The two warm panic tests cover the cached helper's own
panic translation, drain, lowest-ordinal selection, and wave stop boundary;
the stale-hint and source-change tests cover its ordinary-read and freshness
fences. Root-owned quality and full tests must run after applying the archive;
this candidate carries no compile or runtime result.

## Archive controls

`candidate/changed-files.json` is the sorted complete source/test allowlist.
`candidate/model.patch` was checked with `git apply --check --whitespace=error`
against the unchanged worktree. The candidate Rust files were formatted and
checked with `rustfmt`; the parent source file required
`--config skip_children=true` because its archived sibling test module is not
copied into this candidate. The machine-readable receipt records the exact
file hashes, byte counts, patch hash, base revision, and these checks.
