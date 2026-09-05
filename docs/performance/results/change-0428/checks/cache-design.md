# 0428 managed PPTX source-cache design

This is a source and protocol review for the next bounded cache/retention
probe. It does not report a measurement and does not authorize a cache or
latency optimization. The probe should reuse the frozen 0427 `plain` and
`media-rich` PPTX builders and their existing source, preservation, output,
and refusal gates.

The relevant production owners are `SourceBackedPresentation`,
`SourceBackedPresentationEditor`, and their one underlying
`SourceBackedPackage` each. The PPTX owner now forwards the fail-closed
`try_cache_diagnostics()` result. The existing infallible
`cache_diagnostics()` method is retained for compatibility, but it is not an
acceptable evidence source: a poisoned cache lock or diagnostic-counter
overflow must fail the row rather than be recovered into a report. The
forwarding seam and its no-payload-read tests are recorded in
[`forwarding.md`](forwarding.md).

## Ownership and phase boundaries

The source-backed cross-copy selectors remain the production-shaped cases:

```text
pptx_source_backed_cross_copy_plain_lifecycle
pptx_source_backed_cross_copy_media_rich_lifecycle
```

The matching owned selectors remain correctness controls. The media-rich
corpus has eight direct image leaves, each declared as the existing 2 MiB
payload. The direct source-image API is useful for near-limit rows because
`SourceSlide::images()` and `SourceSlide::image(position)` establish the
metadata boundary without reading the embedded payload, while
`SourceSlide::read_image(position)` returns a `SourceImage` holding the
managed `PartData`. Every row must gate the descriptor count, selected part
URI, declared size, payload length, payload digest, and source archive hash
against the same 0427 corpus manifest before it is admitted.

The source-backed editor has an intentional one-way lifetime:

```rust,ignore
let source_view = SourceBackedPresentation::from_read_at_...(...)?;
let destination = SourceBackedPresentationEditor::from_read_at_...(...)?;
let plan = destination.plan_cross_slide_copy(&source_view, ...)?;
let result = destination.publish_cross_slide_copy_to_stream(sink, &plan)?;
```

`publish_cross_slide_copy_to_stream` consumes `destination`. The destination
editor therefore has a final diagnostic point immediately before publication;
there is no valid destination-cache snapshot after publication. The journal
must record `destination_editor_consumed_during_publish` and omit a fake
post-publication destination value. Immediately after publication, the
source view can still be sampled, and both caller-owned budgets can be
sampled. Those are the valid observations of what survived consumption.

Use two independent caller budgets and contexts for a row: one source budget
and context for `SourceBackedPresentation`, and one destination budget and
context for `SourceBackedPresentationEditor`. Cloning within one owner shares
that owner’s budget node; it does not create a cache owner. The protocol must
not put the two roots under a hidden shared parent, because separate roots
allow source-cache and destination-plan reservations to be attributed without
double counting. Do not open a second package, view, or editor only to obtain
diagnostics. The two packages required by cross-copy are the only managed
cache owners. Keep both caller `Budget` handles outside all package scopes so
their current usage can be sampled after the destination has been consumed
and after each explicit drop.

Recommended phase journal for the source-backed lifecycle:

1. After corpus/gate setup, before source or destination package creation,
   record both caller-budget baselines and both source read-counter baselines.
2. After source view and destination editor open, record one fallible cache
   snapshot for each owner and both budgets. Opening is allowed to have loaded
   mandatory root/catalog parts; those reads are the pre-payload baseline.
3. After selected metadata and plan preparation, record both owners again.
   The source view remains available; the editor snapshot is the last one
   that can be obtained if publication is next.
4. Record the destination editor snapshot immediately before the consuming
   publish call. Record the source view and both budgets immediately after
   publication, while the returned result, plan, source view, sink, and
   caller source owners are still held.
5. Drop the returned result, plan, source view, caller source `Arc`s, and sink
   in named scopes, recording the corresponding budget after each scope. A
   source view or `SourceImage` intentionally retained to test pinning must be
   named in the ownership set and dropped before the final row.

Each cache snapshot must retain both the monotonic counters and point-in-time
gauges. Use `SourceCacheDiagnostics::checked_counter_delta` only for intervals
whose before/after snapshots belong to the same owner. Keep source and
destination budget fields in separate namespaces; do not use one owner’s
diagnostic field as the other owner’s total. A process-level total may be
derived from the two independent roots only as an explicitly labelled sum.

The cache gauges mean:

| Field | Interpretation |
| --- | --- |
| `retained_entries`, `retained_bytes` | Current clean cache entries only. A returned `PartData` that survived eviction or bypass is outside these gauges. |
| `in_flight_loads` | Same-part loads currently coordinated by the cache. A completed row must end at zero. |
| `budget_cache_reserved_bytes`, `budget_cache_reserved_objects` | Current reservations owned by cache entries and flights, counted once. |
| `budget_catalog_reserved_objects` | Current package-catalog reservation. It is separate from payload cache objects. |
| `budget_memory_used`, `budget_objects_used` | Current local usage of that owner’s managed budget; it may include that owner’s package and returned handles, but not the other independent root. |

The resource dimensions have different lifetimes. `Memory` and `Objects` are
RAII reservations and should fall when their package, clean cache entry, or
returned handle is dropped. `InputBytes`, `Work`, and managed publication
`OutputBytes` are cumulative `Budget::consume` charges; they do not fall on
cache eviction or package drop. The corresponding diagnostics fields
`budget_input_bytes_used`, `budget_work_used`, and
`budget_output_bytes_used` are therefore monotonic totals, not retained-byte
gauges. A repeated row must report both classes rather than treating growth
of cumulative input/work/output as a retention failure.

## Frozen near-limit rows

The rows below use fresh processes and one worker. They are resource and
correctness rows, not timed comparisons. Cache limits and budget limits must
be present in every report header.

### Exact and one-byte-under managed admission

Use one selected image from the media-rich source slide for the direct cache
admission row. First load only metadata (`images`/`image`) and obtain the
declared uncompressed size `P` before taking the payload-read baseline. The
prevalidated manifest binds the archive, part declarations, and metadata
oracle, but does not pretend to contain a memory reservation. The row's
exact `Resource::Memory` ceiling is derived from the same-policy pinned-root
reservation immediately after opening, plus the selected `P`. Record primed
metadata usage separately. OPC's `reserve_for_load` retries a failed memory
reservation after evicting all unpinned clean entries. Including evictable
metadata in the admission floor would let the nominal one-byte-under request
succeed after eviction. Do not call the total PPTX lifecycle budget simply
`P`: the permanently pinned presentation root remains charged.

Keep the protocol's normal cache policy large enough for the mandatory root,
metadata, and selected payload. The exact row must return the expected bytes,
retain the target entry, and report the matching cache reservation. Its
`Resource::Memory` ceiling is derived in a pre-observation calibration using
the same corpus, cache policy, metadata calls, and source role;
the calibration records the declared payload `P`, the pinned root reservation,
and the primed metadata usage. The exact request may first fail reservation,
evict metadata, and succeed on retry; record those counters as observed.
Check that the post-eviction reservation floor matches the calibrated root.
It is not a claim that the corpus manifest contains a memory
reservation, and the calibration owner is gone before formal observations.
The one-byte-under row repeats the identical metadata boundary with the
calibrated ceiling minus one. It must fail with a typed `ResourceLimit` for
`Memory` before the target payload `ReadAt`, leave no target cache entry or
payload reservation, and write no output. Compare source read calls/bytes
after the metadata baseline, rather than comparing against process-open
reads. The row must also retain the catalog/object values so a surviving
catalog reservation is not mistaken for a payload leak.

Do not silently substitute a cache byte limit of `P - 1` for the one-byte-under
budget row. That is a different, successful oversized-bypass behavior and is
specified separately below. If the full cross-copy plan is also tested under
an exact budget, predeclare the full plan's fixed staging reservation and
validate it independently; the direct selected-image row is the bounded
admission oracle. The `read_image` call can inspect the selected slide XML
again, so the formal source-read baseline must be taken after the metadata
selection boundary, not immediately after package open.

### Pinning and clean-entry eviction

The media-rich source provides equal-size image leaves, so use two distinct
image positions `A` and `B`, `SourceCacheLimits::new(64 * 1024 * 1024, 2)`,
and adequately large independent source and destination budgets. The
two-entry policy is intentional: the source presentation's mandatory
`_presentation` payload is pinned by `SourceInner`, leaving one effective
clean-entry slot for a selected payload.

For the pinning row, load `A` and keep its `SourceImage` alive while asking for
`B`. The pinned root plus pinned `A` leave no evictable cache entry. The
selected-slide XML lookup and `B` therefore take the normal pinned-entry
bypass path even though both fit under 64 MiB; `oversized_bypasses` must not
be used to explain this row. `B` still must return the expected bytes, and
its returned handle owns any temporary managed reservation. Drop `A`, then
load `B` again in a named step. The root plus clean `B` state must now be
possible, with the expected eviction/read transitions.

For the clean-entry eviction row, use the same 64 MiB/2-entry policy but let
`A`'s `SourceImage` drop before requesting `B`. The selected-slide metadata
and `A` can then be evicted as needed; require the retained-entry/byte gauges,
eviction delta, and source-read delta to match the actual phase sequence
rather than assuming a fixed single eviction. Re-reading `A` after the
eviction must produce another cold read. Check both payload digests; an
eviction counter without a read/oracle check is not enough. This follows the
existing OPC invariants `managed_cache_eviction_releases_unpinned_reservation`
and `managed_cache_does_not_evict_externally_pinned_entry_and_bypasses_retention`.

The row must distinguish `budget_cache_reserved_bytes` from that owner's
`budget_memory_used`: a bypassed returned handle can consume managed memory
while the cache reservation remains unchanged.

### Oversized bypass

Use the same selected payload `P` with `SourceCacheLimits::new(P - 1, 128)`
and a memory budget that can hold the temporary returned payload. The generous
entry limit keeps the pinned root and metadata from becoming the reason for
the result. A valid payload larger than the cache byte capacity must succeed
but never enter the cache:
`bypasses` and `oversized_bypasses` increase, the target entry/bytes and cache
reservation do not increase, and a second access performs another cold source
read. Do not require package-wide retained bytes to be zero because mandatory
root/slide entries can fit under this limit. Do not label this row an
admission refusal, an allocation failure, or a managed-budget failure.

### Repeated publication with bounded reservations

Keep one source view and its source cache alive across three complete
publications. Create a fresh destination editor for each iteration with the
same destination context; publication consumes that editor. Reuse the same
source corpus, destination corpus, cache policy, source context, and
destination context, and set the destination `OutputBytes` limit to three
times the prevalidated expected output length. Capture source-view and
destination-editor diagnostics at their valid phase boundaries, publish, then
drop the result, plan, destination editor (inside publication), and sink
before the next iteration. No diagnostic-only package may be opened between
iterations.

The source budget is allowed to retain its source-view/cache baseline between
publications and must be checked again after the final source-view and caller
source-owner drop. The destination `Memory` and `Objects` usage must return to
its pre-iteration baseline after each plan/result/editor/sink drop; its
cumulative `OutputBytes` usage is expected to rise and must remain within the
three-publication cap. The destination editor's last snapshot remains
pre-publication because publication consumes it.

Require stable output/preservation/source-immutability oracles on every
iteration, `in_flight_loads == 0` at each completed owner, and retained source
entries/bytes within the declared cache limits. Require cumulative source
`InputBytes`/`Work` and destination `OutputBytes` to be monotonic within their
configured limits, but make no release assertion for them. Source read
counters and cache counters may increase once per cold load; they are event
totals, not evidence of unbounded retention. An optional fourth publication
must be a separate typed output-budget refusal with zero sink bytes; it must
not be folded into the three successful rows.

## Observer and acceptance guardrails

The fallible diagnostics calls are content-free and do not load payloads.
If the optional cache lock observer is enabled, allocate its fixed event
storage before the measured region. The `Finished` callback runs while the
cache-state mutex is held; it must not allocate, block, call back into the
package, or acquire another cache/budget lock. A failed observer or poisoned
state invalidates the row.

For every row retain source revision, binary, corpus manifest, cache limits,
all six budget limits, context sharing, source read counters, cache snapshots,
typed result/error, output bytes/hash, and ownership labels. Refusal rows
must prove no target payload read and no sink bytes. Successful rows must
prove exact payload bytes and the existing 0427 semantic/preservation gates.

This evidence can establish bounded cache behavior, typed admission/refusal,
pinning, eviction, bypass, and current budget release for the named owner
scopes. It cannot establish RSS release, allocator-page return, physical I/O
behavior for a non-memory source, latency improvement, or an optimization
speedup. In particular, the destination editor's post-publication cache
state is unavailable by API design; the separate caller budgets and surviving
source-view snapshot are the complete valid post-publication evidence.

The design follows the managed cache invariants in the 0086/0088 records and
the accepted ownership/resource contracts in ADRs 0003, 0005, and 0006. It is
intentionally narrower than a general cache benchmark: one source budget and
one destination budget, one source-backed owner per required package, fixed
source-read oracles, and explicit drop scopes keep the interpretation
reviewable.
