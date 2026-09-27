# 0798 OPC checked-iterator census hook

This directory is a source-only diagnostic archive for commit
`5f6ec5ca31df84ea7a103e5ba610a1d59e8f3c60`.
The `before` module is the exact OPC helper at that commit. The `after` module
keeps its iterator/parser behavior and adds a thread-local census whose only
purpose is to explain checked-iterator ownership during the 0798 probe. No
production adoption, latency claim, allocation claim, or public-workflow claim
follows from this archive.

## Public probe API

The instrumented module exposes these functions and plain result types:

```rust
pub fn begin_attribute_census_0798();
pub fn finish_attribute_census_0798() -> AttributeCensusReport0798;
```

`begin` replaces any unfinished session on the calling thread. `finish`
disables the registry before taking it and returns
`AttributeCensusReport0798 { rows, iterator_starts, iterator_clones,
iterator_drops, live_instances_at_finish, counter_saturated }`. Each
`AttributeCensusRow0798` carries the exact raw source bytes, raw element-name
bytes, instance and lineage IDs, local `next` outcome counts, the inherited
successful prefix for clones, complete unchecked lexical counts, lifecycle
flags, termination state, and a per-row saturation flag. The types contain
only `Vec<u8>`, integer, boolean, and enum fields so the standalone probe does
not need a serialization dependency from `litchi-opc`.

The intended call boundary is:

1. call `begin_attribute_census_0798` on the probe thread;
2. construct and use the public owner operation;
3. stop the clock;
4. drop the operation and all iterator owners;
5. call `finish_attribute_census_0798` and serialize the returned plain data
   outside the measured region.

The registry is thread-local by design. A token also carries the originating
`ThreadId`, so moving an iterator to another thread cannot accidentally mutate
that thread's session. Such a move leaves the originating row live and is
reported as a failed lifecycle/conservation case. A stale token from a prior
session is ignored after `begin` replaces the registry.

## Events and conservation

`CheckedAttributes::next` records one call and classifies the exact returned
item as successful attribute, error, or `None`. The existing phase transitions,
error ordering, duplicate positions, and fused behavior remain in the same
code paths. `Clone` creates a new row in the same lineage and resets its local
event counters; `starting_successful_yields` records the successful prefix
already traversed by the source iterator. This avoids treating a clone made
after one successful item as if it had consumed the entire source.

`Drop` performs a separate `unchecked_attributes` scan over the immutable tag.
That scan records successful lexical attributes, yielded lexical items, and
the first lexical error (if present), then stops at that error so the reported
counts have a deterministic lexical prefix. It marks whether the checked
iterator was dropped before a terminal result and whether the observed
successful prefix was incomplete.
Rows still live at `finish` receive the same lexical scan and
`live_at_finish = true`; they are reported as a qualification failure rather
than being inferred as dropped.

For a well-scoped session the probe checks:

```text
iterator_starts + iterator_clones
    = iterator_drops + live_instances_at_finish
```

All event and aggregate counters use checked increments with explicit
saturation at `u64::MAX`. Saturation is retained in the report and fails the
probe's diagnostic qualification. IDs use the same checked/saturating policy;
the generated 0798 inputs are bounded and cannot approach it.

The disabled path checks only the thread-local enabled bit and performs no
registry allocation, source copy, lexical scan, or event bookkeeping. The
enabled hooks add source copies, thread-local borrows, event branches, and
drop-time scans to the diagnostic build. Those costs are intentionally part of
the instrumented diagnostic environment and are excluded from every 0798
timing interpretation. The before module has no census symbols or fields.

## Archived tests

`after/census_tests_0798.rs` covers disabled/no-session behavior, full
consumption, first-only and no-consumption drops, duplicate errors followed by
a valid lexical tail, clone prefix accounting, stale sessions, live-at-finish
qualification, foreign-thread isolation, and saturation policy.
`after/canonical_tests.rs` is the existing OPC helper test source. Its
filesystem test that compares all production copies is intentionally loaded by
the root-owned mirror harness only where the copied source paths exist; it is
not part of the diagnostic hook's adoption decision.
