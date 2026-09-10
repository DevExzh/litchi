# 0492 production next patch: managed forward-start source read-ahead

This is the implementation plan for the first production-compatible 0492
patch. It follows the accepted 0492 budget review and keeps the 0491 exact,
cold, and splice evidence on its existing constructors. It is a plan only;
the production crates and the benchmark helpers are unchanged by this
document.

The private pilot showed a useful locality question: the successful replay has
19 nonempty delayed requests in the open and document-preparation phases. That
number does not by itself authorize a cache or a performance claim. The
production seam must charge the bytes actually fetched from the source, fence
the source on every cache use, and keep an exact mode for every operation that
preserves or republishes the source.

## Ownership and API

The cache belongs in `litchi-opc`, alongside the existing `SourceReader` in
`crates/litchi-opc/src/source_backed.rs`. `litchi-opc` already owns the
`SourceSnapshot`, `IndexedArchive<SourceReader>`, source-version mapping, ZIP
preservation paths, and the `ExecutionContext` passed to lazy reads. This is
the only layer that can see both the logical ZIP request and the physical
source read while charging the same managed context.

Add an OPC-owned policy and re-export it from `crates/litchi-opc/src/lib.rs`:

```rust
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum SourceReadPolicy {
    Exact,
    ForwardStart { window_bytes: NonZeroUsize },
}

impl SourceReadPolicy {
    pub const fn exact() -> Self;
    pub fn forward_start(window_bytes: usize) -> Result<Self>;
}
```

`forward_start` validates a positive window and a hard finite ceiling, for
example `MAX_SOURCE_READ_AHEAD_BYTES = 64 * 1024`. The package constructor
validates the public value again so a caller that constructs the
`NonZeroUsize` variant directly cannot bypass that ceiling. An invalid value
returns a typed `OpcError::InvalidSourceReadPolicy` carrying the requested and
maximum values. Four KiB is the first measured policy value; it is not a
default.

Keep all existing `SourceBackedPackage::from_*` signatures and behavior. They
pass `SourceReadPolicy::Exact` to one expanded private constructor:

```rust
fn from_read_at_inner(
    source: Arc<dyn ReadAt>,
    limits: ReadLimits,
    cache_limits: SourceCacheLimits,
    context: Option<ExecutionContext>,
    source_read_policy: SourceReadPolicy,
) -> Result<Self>;
```

Add only the fully explicit opt-in constructors first, so the public API does
not acquire a large matrix of unmeasured convenience combinations:

```rust
pub fn from_read_at_with_limits_and_cache_limits_and_source_read_policy(
    source: Arc<dyn ReadAt>,
    limits: ReadLimits,
    cache_limits: SourceCacheLimits,
    source_read_policy: SourceReadPolicy,
) -> Result<Self>;

pub fn from_read_at_with_limits_and_cache_limits_and_source_read_policy_and_execution_context(
    source: Arc<dyn ReadAt>,
    limits: ReadLimits,
    cache_limits: SourceCacheLimits,
    source_read_policy: SourceReadPolicy,
    context: ExecutionContext,
) -> Result<Self>;
```

The validation-only constructor
`SourceBackedPackage::from_read_at_for_validation` passes `Exact` explicitly.
The existing `from_path`, `from_vec`, and sequential-reader constructors also
remain exact unless a later owner-specific API deliberately adds the same
policy argument.

The DOCX owner in `crates/litchi-docx/src/source_backed.rs` should add the
matching fully explicit managed forwarding constructor. It constructs the
existing `Package { package, execution }` after passing the policy to OPC; it
does not implement a second cache in the format crate. The first caller should
be the warm semantic DOCX owner used by the benchmark. Validation in
`crates/litchi-docx/src/validation.rs`, filesystem cold reads, and splice
owners continue to call the existing exact constructors.

This placement is compatible with ADRs 0005, 0006, 0008, 0010, and 0011. It
does not change `litchi_core::ReadAt`, add an execution dependency to
`soapberry-zip`, expose an archive implementation type through a facade, or
make an implicit global policy choice.

## Cache state and lifetime

Extend the private adapter as follows:

```rust
#[derive(Clone)]
struct SourceReader {
    snapshot: SourceSnapshot,
    read_ahead: Option<Arc<ArchiveReadAhead>>,
}

struct ArchiveReadAhead {
    mode: Mutex<ArchiveReadModeState>,
    forward_reads_done: Condvar,
    state: Mutex<ArchiveReadAheadState>,
}

struct ArchiveReadModeState {
    mode: ArchiveReadMode,
    forward_reads_in_flight: usize,
}

struct ArchiveReadAheadState {
    bytes: Vec<u8>,
    start: u64,
    valid_len: usize,
    memory_reservation: Option<Arc<Reservation>>,
}

enum ArchiveReadMode {
    Forward,
    Exact,
}

struct ForwardReadLease {
    owner: Arc<ArchiveReadAhead>,
}
```

The actual field names may follow local style, but the ownership rules are
required. `ArchiveReadAhead` is constructed after `SourceSnapshot` captures
the source version and length and before the first
`IndexedArchive::from_reader_with_limits` call. Its one `Arc` is cloned by
`SourceReader` and retained by the `IndexedArchive`; the same state therefore
covers ZIP tail/EOCD/ZIP64/central-directory reads during open and member reads
later through `read_entry`, `read_entry_with_accounting`, and read sessions.
No change is needed in `crates/soapberry-zip/src/office.rs` or
`crates/soapberry-zip/src/locator.rs`.

Allocate one buffer with capacity
`min(window_bytes, snapshot.length as usize)` after checked conversion. For a
managed context, reserve that capacity as `Resource::Memory` before the
allocation and retain the reservation in `ArchiveReadAhead` until it is
dropped. Use `try_reserve_exact`, require the resulting capacity to equal the
reserved amount, then resize the buffer to that capacity. If allocation or
capacity admission fails, drop the reservation and fail the open; do not fall
back to an uncharged buffer. An empty source has no buffer and no memory
reservation. An unmanaged policy still has the same finite byte ceiling but
has no budget token, matching existing compatibility constructors.

The read-ahead reservation is separate from `PartCache`'s retained
`PartData` reservations. It is outstanding `Resource::Memory`, not cumulative
`Resource::InputBytes`; it remains charged while the window can satisfy a hit.
The cache is not placed on `SourceSnapshot`: that snapshot is cloned into
`SourceArtifact` and exact publication helpers. Keeping the window on the
archive-owned `SourceReader` prevents an artifact clone from retaining or
using a read-ahead buffer.

There is no separate metadata `ReadAt` adapter. `SourceReader` is already the
sole adapter used by ZIP locator probes, central-directory indexing, strict
payload reads, and read sessions. `SourceSnapshot` already owns the captured
`SourceVersion` and length. A helper on the snapshot may provide the common
physical reservation/retry code, but it must receive the archive-owned policy
and state explicitly rather than making every snapshot clone a cached source.

## Forward-start read algorithm

`SourceReader::read_at` keeps the current exact branch byte-for-byte and sends
the forward policy through an archive-only helper such as
`ArchiveReadAhead::read_at`. The helper must not call the current exact helper
on the caller's output after it has filled the window: that would charge the
logical output and leave the overfetch uncharged.

For every read-ahead call, including a cache hit, perform the following
sequence:

1. Check `ExecutionContext` when present and call an unconditional source
   version fence. Do not use only
   `SourceSnapshot::ensure_current_io_if_monitored`: `monitor_reads` starts
   false for an ordinary open, while cached bytes must never be served after a
   version change. A source-version error invalidates the positive window and
   returns the existing typed `OpcError::SourceChanged` path.
2. If `output` is empty, return zero after the same context/freshness checks.
   Do not count a hit, miss, fill, or physical request, and do not invalidate a
   current positive window.
3. Under the read-ahead state lock, test whether the request offset is in the
   retained half-open range `[start, start + valid_len)`. On a hit, copy at
   most `output.len()` and the bytes through the retained end. A request that
   crosses the end returns its positive prefix; it does not trigger a second
   source call in the same `ReadAt` operation. Check context and source
   freshness again after the copy, invalidate on failure, and return the
   typed error.
4. On a miss, compute with checked arithmetic:

   ```text
   fill_start = request_offset
   available  = captured_length - fill_start, or zero when at/after EOF
   fill_len   = min(window_bytes, available)
   ```

   An offset at or beyond the captured end returns zero without a physical
   read. Clear `valid_len` before a fill so a failure cannot leave stale bytes
   visible.
5. Reserve exactly `fill_len` as cumulative `Resource::InputBytes` before
   calling the underlying `snapshot.source.read_at`. Use the existing
   `read_source_at_with_context` reservation loop, or factor its reservation,
   `Interrupted`, count-validation, and commit core into
   `read_source_fill_with_context`. The physical output passed to that helper
   is the window buffer, not the caller's output. Preserve its useful dynamic
   shrink behavior: when the remaining input budget is smaller than the
   configured window, retry with the remaining positive amount; when no input
   bytes remain, record the existing reservation failure and make no provider
   call.
6. Preserve the current interruption semantics. An `Interrupted` result
   accepts no bytes, retains the current physical reservation, checks
   cancellation and source freshness, consumes one unit of `Resource::Work`,
   and retries. A transport error drops the reservation. A returned count
   larger than the physical buffer is invalid. A successful call commits the
   reservation with the actual returned count, including a short return.
7. Check the source version and execution context after the provider returns
   and before publishing the window. A positive return is retained only as
   `[fill_start, fill_start + returned)`, with checked end arithmetic. A zero
   return is discarded and passed through as ordinary zero progress. A source
   change or cancellation discards the buffer, even when the provider
   returned bytes; accepted physical bytes remain charged to
   `InputBytes`.
8. Copy the retained prefix into the caller output, then perform the final
   context and unconditional source fence before returning. Map
   `SourceChangedIoError` through the existing `map_io_error` path when the
   ZIP `ReaderAt` boundary requires an `io::Error`.

The fill lock may be held while the synchronous provider call runs in the
first implementation. This gives one fill owner and prevents duplicate
windows with the smallest state machine. Context is checked before the lock,
after acquiring it, before the provider call, after every `Interrupted`
retry, and after the provider call. A provider that blocks inside one
`ReadAt` remains subject to the same cooperative cancellation boundary as the
current exact helper; cancellation is observed before any cache publication
or successful return. If later profiling shows lock wait matters, an
in-flight state and timed condition-variable wait can move the provider call
outside the lock without changing the physical reservation rules.

Forward-start is deliberate. An aligned window can return a short positive
prefix that ends before the caller's requested offset; treating that as a
zero result creates a false EOF. Retrying from the requested offset would need
another fill and another budget decision. A forward-start window always begins
at the logical request, so a positive short return is useful and its actual
prefix can be retained. Aligned windows remain a separate experiment requiring
coverage metadata and dedicated short-read proofs.

## Exact publication mode

Keeping old constructors exact is necessary but insufficient: an opt-in
`SourceBackedPackage` can still be passed to a publisher later, and the same
retained archive currently feeds raw-copy preservation. The first production
patch uses a monotonic exact mode on the shared archive adapter.

`ArchiveReadAhead::disable_for_exact_publication()` is a monotonic transition,
rather than an operation guard. It acquires the mode lock, checks the source
and execution context, changes the mode to `Exact`, waits on
`forward_reads_done` until the count of forward calls already admitted is
zero, then locks `ArchiveReadAheadState` to invalidate the positive window and
take both its `Vec` and `memory_reservation` out of the state. The taken
allocation and reservation are dropped after the state lock is released. The
mode lock is then released before the publisher makes any ZIP read.
Every `SourceReader::read_at` takes the mode lock long enough to select its
mode and, for a forward call, returns a `ForwardReadLease` whose `Drop`
implementation decrements `forward_reads_in_flight` and signals the condition
variable. The lease is held around every return and unwind path of the
physical read, so the exact transition cannot wait forever after an I/O error
or panic. A call admitted after the transition selects `Exact`, so it does not
create a lease. Exact dispatch calls the existing `read_source_at_with_context` path,
including its dynamic input-budget shrink, without consulting the window.
There is deliberately no restore step: once a package has entered an exact
publication/materialization path, all later reads on that package are exact.
This avoids a restore race and makes the transition safe for borrowed APIs
such as `to_opc_package(&self)` as well as consuming publishers. The loss of
acceleration is explicit, and dropping the buffer releases its managed memory
reservation.

Perform this transition before any archive read in the shared low-level paths:

* `into_opc_package_inner` and `to_opc_package_inner`;
* `write_single_part_overlay_to_stream` and `write_part_overlays_impl`;
* `write_topology_to_stream` and the common
  `write_changed_overlays_with_appended_inner` preservation path;
* precompressed authorization and source-XML capture used by splice or
  source-preserving publication.

An empty overlay/topology plan may continue directly to
`write_exact_source`, because that helper reads `SourceSnapshot` through the
exact source-artifact path and does not touch `IndexedArchive`. Any nonempty
operation that can inspect or copy archive members performs the monotonic
exact transition before reading the selected part or building a preservation
plan. Semantic
`PartView::data`, verified decoded reads, and ordinary decoded part streaming
may use forward mode because they return semantic payloads rather than
claiming an exact source ZIP traversal.

`SourceArtifact::fingerprint` and `SourceArtifact::write_to_stream` remain
direct `SourceSnapshot` operations and therefore stay exact regardless of the
package's policy. Source-part splice replay must still perform this exact
transition or use an exact constructor before it consumes parser ranges. A
source change, cancellation, or archive error after the transition leaves the
package in exact mode and emits no stale cache result. The transition also
serializes the policy decision with concurrent semantic reads: an already
running forward read drains, then publication gets an exact view; reads
arriving after the transition are exact as well. This is the proposed
concurrency design and needs the focused race test below before production
code can be considered proven.

This transition preserves the existing exact-range and raw-copy proof without
duplicating an `IndexedArchive` and its catalog. A future implementation could
build a separate exact archive view, but that would repeat ZIP tail/central
reads and retain a second index; it is not needed for the first patch.

## Focused tests

Add tests beside `SourceReader` and `SourceBackedPackage` in
`crates/litchi-opc/src/source_backed.rs` before enabling a benchmark arm.
They should use an instrumented immutable `ReadAt` with controlled short
returns, `Interrupted`, version changes, cancellation, and a physical call
log.

The adapter tests must cover:

* Exact and policy-constructor validation, including zero, one, the hard
  ceiling, and a direct over-ceiling enum value rejected by package open.
* Differential random overlapping, nested, split, empty, EOF, and near
  `u64::MAX` requests against `OwnedSource`; verify returned counts, bytes,
  checked offsets, and untouched caller sentinels.
* Forward-start physical fills no larger than the configured window, one
  same-window hit with no provider call, a miss with one fill, a request that
  crosses the cached end returning a positive prefix, and a new fill after the
  caller advances.
* Empty output with no hit/miss/fill and no input or memory charge change;
  zero returns with no retained entry; positive short prefixes retained only
  through their actual end; and provider caps smaller than the window.
* Source-version changes before a hit, during a fill, and after the copy while
  `monitor_reads` is still false. Each returns typed `SourceChanged` and
  leaves no stale positive window.
* `Interrupted` retries, cancellation before a hit, cancellation before a
  fill, and cancellation after physical bytes but before publication. Verify
  `Work`, reservation release, and no cache publication.
* Managed budget accounting with a child and ancestor budget: a 4 KiB request
  that returns 1 KiB consumes 1 KiB of cumulative `InputBytes`; a hit consumes
  none; a budget with only a smaller remainder uses the existing dynamic
  shrink; a zero remainder makes no provider call; and dropping the package
  releases the retained `Memory` reservation while input bytes remain
  cumulative.
* Memory allocation admission: the actual vector capacity equals the managed
  reservation, an allocation/capacity failure drops the token, and an empty
  source allocates neither buffer nor reservation.
* Concurrent readers share one state without mixed bytes or duplicate fills.
  A gated source should prove that the monotonic exact transition waits for
  an in-flight forward fill and that no forward operation starts after the
  mode changes. A later read on the same borrowed package must remain exact,
  and the transition must release the read-ahead memory reservation.
* A forward-policy package's materialization, overlay, topology, precompressed
  authorization, and splice entry points use exact source requests after the
  transition. A source-artifact fingerprint/copy remains exact. A transition
  error leaves the package in exact mode once the transition has begun.
* Existing exact constructor, filesystem cold, and splice tests retain their
  original physical ranges and source-version behavior. The new policy has no
  effect unless explicitly supplied.

At the DOCX boundary, add one managed smoke test through the new
`litchi-docx::source_backed::Package` constructor. Main-document text and
part semantics must match the exact-policy oracle; a later publication must
perform the monotonic exact transition. Do not use the unmanaged tool wrapper
as evidence for managed behavior.

## Benchmark hookup and acceptance

After the focused tests pass, add one production arm to the next performance
bundle (0493 or its replacement), rather than changing the sealed 0492
producer. It should call the DOCX managed constructor with
`SourceReadPolicy::forward_start(4096)` and the same finite
`ExecutionContext` shape as the exact arm. The instrumented underlying source
records delegated physical ranges; a separate SourceReader-boundary
diagnostic records logical ZIP requests. Keep these vectors separate.

The report must include the selected policy, `budget_managed = true`, source
version before/after, logical request count and bytes, physical fill ranges,
actual accepted `InputBytes`, retained read-ahead capacity, cache hit/fill
counts, `Interrupted`/cancellation outcomes, and peak RSS/allocation scope.
Physical fills may overlap media compressed spans; preserve the existing media
range oracle and report the overlap rather than inheriting a zero-overlap
claim from 0491. Exact logical ranges, physical/logical identity where the
policy requires it, and byte-conservation checks remain mandatory for each
arm.

Run the formal before/after protocol with the same four arms, 16 children, and
2 repeats used by the current gate. Baseline exact, production forward-start,
and any 16 KiB comparison must be separate labeled arms. The
private unmanaged pilot can explain the hypothesis but cannot supply the
managed budget or production timing result. A gate should require the
physical-call reduction demonstrated by the source trace together with
conserved logical bytes, exact source fences, bounded memory, and no semantic
or publication regression; latency alone is insufficient under ADR 0005 and
ADR 0008.

The production change is ready for integration only when the old constructors
remain exact, focused OPC/DOCX tests pass, the exact-mode concurrency test
passes, managed `InputBytes` and `Memory` reservations are observable in the
receipt, and the fresh production 0493 capture—not the sealed 0491/0492
artifacts—supports the forward-start claim.
