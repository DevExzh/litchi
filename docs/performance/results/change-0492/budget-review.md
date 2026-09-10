# 0492 bounded range read-ahead budget review

This review covers the proposed DOCX range read-ahead experiment and the
smallest shape that could later be made production-compatible. It does not
authorize a change to the sealed 0491 evidence, the filesystem cold protocol,
or the source-splice protocol.

The 0492 ADR refresh records all 30 accepted ADR/index hashes as unchanged
from 0491. The relevant hard constraints are the positional, immutable source
contract and hierarchical execution budgets in ADR 0005; preservation,
fail-closed validation, and source-version requirements in ADR 0006; the
dependency and ownership rules in ADRs 0002, 0010, and 0011; and the measured,
continuously verifiable migration policy in ADR 0008.

## Evidence and current ownership

The successful retained diagnostic replay is the strongest available locality
evidence. In
`validation/source-replay-diagnostic-run2.stdout`, the warm source-backed
DOCX lifecycle makes 19 nonempty requests (18 calls during open and 5 during
document preparation; four of those 23 calls have an empty output buffer):

| phase | all calls | empty calls | nonempty calls | nonempty bytes |
| --- | ---: | ---: | ---: | ---: |
| open | 18 | 3 | 15 | 2,480 |
| document preparation | 5 | 1 | 4 | 1,486 |
| text query | 0 | 0 | 0 | 0 |

The delayed arm therefore pays the fixed service delay 19 times for 3,966
requested/returned bytes. The largest request is 1,424 bytes. The replay is
diagnostic provenance rather than a formal provider trace: it uses its own
`ReadAt` recorder and does not include the lifecycle's range transport adapter.
The matching 19-call/3,966-byte aggregate is a strong locality lead, but a
fresh 0492 capture must serialize the logical and physical range vectors
before claiming provider-level identity.

The existing layers are intentionally separate:

* `litchi_core::ReadAt` (`crates/litchi-core/src/source.rs:62`) is immutable
  positional access. Its `read_at` may return a short prefix, and its
  delegated `SourceVersion` is the source identity/revision contract.
* `SourceReader` (`crates/litchi-opc/src/source_backed.rs:2632` and
  `:2662`) adapts that trait to soapberry-zip's output-first `ReaderAt`.
* `read_source_at_with_context` (`crates/litchi-opc/src/source_backed.rs:11030`)
  is the current context-aware physical read boundary. It checks
  cancellation, reserves `Resource::InputBytes` for the delegated output
  window, retries `Interrupted` while charging `Work`, validates the returned
  count, and commits only accepted bytes.
* `IndexedArchive` (`crates/soapberry-zip/src/office.rs:874`) owns ZIP layout
  and payload descriptors. It cannot depend on `litchi-core` to charge a
  Litchi execution budget.
* `ZipLocator` (`crates/soapberry-zip/src/locator.rs:1304`) scans the tail in
  bounded chunks. `locate_in_reader_impl` and `finish_locate_in_reader`
  (`:732` and `:692`) issue the EOCD, ZIP64, and first-central-header probes.

Consequently, a generic wrapper placed above `SourceReader` can make the
benchmark fast while silently overcharging or undercharging a managed open.
The present 0492 tool-layer pilot is valid only as an unmanaged, explicitly
named experiment. It must not be described as production budget coverage.

## Decision: use a forward-start window for the first pilot

The first bounded policy should start its one window at the caller's requested
offset:

```text
fill_start = request_offset
fill_len   = min(configured_window, captured_length - request_offset)
```

The window stores only the returned half-open range
`[fill_start, fill_start + returned)`. It has one preallocated buffer, one
mutex-protected state, no LRU, no cross-sample state, and no background task.
The 4 KiB window is the appropriate first value for the retained trace. A
16 KiB comparison needs a separate measured arm; it must not become the
default based on the 3,966-byte total.

Forward-start is preferable to an aligned window for the first correctness
gate. Suppose a 4 KiB aligned fill begins at zero, the caller asks at offset
3,000, and the provider returns a permitted short prefix of 1,000 bytes. The
cache contains no byte at the caller's offset. Returning zero would make
`ReaderAt::read_exact_at` report EOF even though bytes exist at 3,000. Retrying
from the aligned fill's end would turn one logical request into an unbounded
or hidden sequence of physical requests and would need a second budget
reservation policy.

A forward-start fill can return only the accepted prefix beginning at the
requested offset. A short return is therefore either useful to the caller or
an ordinary zero-progress result from the underlying source; it cannot create
an EOF solely because the window was aligned before the requested offset. A
positive short prefix is retained with its actual end, while a zero-length
return is discarded and passed through as ordinary zero progress so the
existing caller can apply its normal EOF/error handling.

For this trace, a 4 KiB forward policy is expected to make three physical
fills if the provider returns the requested bytes: the 22-byte EOCD tail, the
1,327-byte central-directory tail through EOF, and `0..4096`. This is a
diagnostic expectation, not a result. The final fill crosses the first media
member's compressed range by about 69 bytes, so the candidate must report
that physical overlap. It cannot inherit 0491's zero-media-overlap statement.
An aligned 4 KiB policy could reduce the trace to two fills, but its tail fill
would begin at `16789504` and fetch unrequested XML payload ranges as well;
that is a separate tradeoff and a separate claim.

## Budget-compatible production shape

The benchmark wrapper may remain private to
`tools/perf-baseline/src/docx_provider_lifecycle.rs`. If the measured result
justifies production work, the first production owner should be
`litchi-opc`, at the `SourceReader`/`SourceSnapshot` boundary, rather than a
public `ReadAt` implementation in a format crate or a new dependency from
soapberry-zip to `litchi-core`.

The smallest complete production seam is a private `ArchiveReadAhead` handle
on `SourceReader`, constructed beside the `SourceSnapshot` immediately before
`IndexedArchive::from_reader_with_limits` begins. `SourceReader` is cloneable
and the ZIP ownership path retains it inside the archive, so the handle should
be an `Arc` around one
mutex-protected window; `IndexedArchive<SourceReader>` retains that same reader
for both locator/index reads and lazy member reads. The package's member-load
dispatch at `source_backed.rs:9639-9811` always calls the retained archive
(`read_entry`, `read_entry_with_accounting`, or a read session), and the ZIP
payload readers at `office.rs:3830`, `:4139`, and `:4262` therefore cross this
same adapter. No change to `soapberry-zip::IndexedArchive` is needed.

Keep the mutable window out of the general `SourceSnapshot` state if possible.
That snapshot is cloned into `SourceArtifact` and exact publication paths
(`source_backed.rs:2754-2806` and `:10759-10875`); sharing a retained read-ahead
buffer there would extend its `Memory` charge into those owners or invite an
exact-copy path to consume a cache. A private `SourceSnapshot` helper may own
the common physical-read/accounting primitive, but only the archive-owned
`SourceReader` should select the read-ahead policy. This keeps the cache's
lifetime and semantics at the ZIP archive boundary.

Because preservation writers also traverse `self.archive` for copy actions,
the policy's scope must be explicit. The first production caller should pass
it only from the warm DOCX semantic-read owner; all existing cold, exact-copy,
topology, and splice constructors continue to pass `Exact`. If a future public
package can both opt into read-ahead and publish from the same archive, its
publication path needs an explicit exact-read guard before it calls the ZIP
writer. A package-wide flag silently changing raw-copy ranges would invalidate
the existing exact-range evidence.

The state must retain a finite configured capacity and, on a managed path, a
`Resource::Memory` reservation for that capacity for as long as it can hold
the buffer. Use `min(window_bytes, captured_length)` for the allocation (and
zero for an empty source), preallocate it before the timed/open phase, require
the actual vector capacity to equal the reservation, and release the
reservation when the archive-owned handle is dropped. The no-read-ahead path
must remain byte-for-byte and allocation-shape compatible with the current
path.

An additive owner API is ADR-compatible when it lives in `litchi-opc`, keeps
the existing constructors at `Exact`/disabled behavior, and accepts only a
finite validated policy such as `SourceReadPolicy::Forward { window_bytes }`.
It must not extend `litchi_core::ReadAt`, put execution context into
`soapberry-zip`, or make a global/default policy decision. A private
`from_read_at_inner(..., source_read_policy)` parameter is the smallest first
patch. If a production caller later needs opt-in, add one additive
`SourceBackedPackage` constructor (and the corresponding managed constructor)
that takes the OPC-owned policy; do not multiply policy into the core trait or
the ZIP generic API. The benchmark wrapper remains tool-private until that
owner API has managed-budget tests.

The intended ownership split is therefore:

```text
SourceBackedPackage::from_read_at(..., SourceReadPolicy)
  -> SourceReader { snapshot, read_ahead: Arc<ArchiveReadAhead> }
  -> IndexedArchive<SourceReader>
```

`SourceReader` remains the ZIP `ReaderAt` adapter and continues to receive the
captured length from the package constructor. Its `read_at` dispatches either
to the existing exact helper or to an archive-only helper that performs the
physical fill. There is no need for a second metadata `ReadAt` adapter: the
existing `SourceReader` is already the sole generic seam used by
`ZipLocator`, central-directory indexing, strict payload readers, and read
sessions. If a helper is placed on `SourceSnapshot`, pass the archive policy
and state into it explicitly; do not make every clone of the snapshot an
implicit read-ahead source.

An empty caller buffer remains a zero-byte operation after the same context
and monitored-source fences used by the exact helper. It must not count as a
hit or miss, invoke a physical fill, or invalidate the current positive
window; this preserves the four empty calls in the replay shape.

For a nonempty request the source-reader operation should perform this
sequence while the one fill mutex is held:

1. Check the execution context and the captured source version. An existing
   cached range is usable only when its version equals the captured version
   and it contains the requested start. Copy at most the requested output
   length; a request crossing the cached end returns the prefix and lets the
   caller issue its next positional request.
2. On a miss, calculate the forward-start fill length with checked subtraction
   from the captured source length. Reserve exactly that finite fill window as
   `Resource::InputBytes` before calling the source. The reservation is the
   physical fill reservation, not the caller's logical output length. The
   helper must be called on the underlying `snapshot.source` with the physical
   fill buffer; placing a cache above the current helper would reserve only
   the logical caller buffer and undercharge the overfetch.
3. Use the existing `read_source_at_with_context` machinery, or factor its
   retry/count-validation core so that it can serve this fill, preserving its
   `Interrupted`/`Work`/cancellation behavior. Do not add a second retry loop
   in the cache. Commit the reservation with the number of bytes actually
   accepted by the source. A failed transport read releases the reservation;
   bytes accepted before a later version refusal remain accounted as consumed
   input.
4. Check the source version again. Publish the returned prefix only when it
   still equals the pre-fill version. On a version change, discard the window
   and return the existing typed `SourceChanged` path through
   `SourceChangedIoError`; never serve the newly read bytes as a cache hit.
5. Copy the requested intersection from the published range. A short fill is
   retained as the positive prefix actually returned, never as an assumed
   zero-filled suffix. A zero return at a non-EOF source offset is not
   transformed into a cache entry.

This placement avoids double charging. The outer source-reader call must not
also reserve the logical output length when the read-ahead branch owns the
physical fill reservation. Cache hits consume no `InputBytes`, matching the
existing source-cache convention that a retained successful load does not
re-read the source. Direct no-read-ahead calls continue to use the current
helper unchanged. If keeping the current outer helper is necessary, the
read-ahead branch must be explicitly rejected whenever an
`ExecutionContext` is present; merely documenting the difference is not
enough.

The adapter must delegate `len` and `version`, retain the original identity and
revision, and use checked `u64` arithmetic for every end calculation. It must
return an I/O error on a poisoned mutex, an invalid source count, an offset or
length conversion overflow, or a version-check failure. The OPC owner should
map source-version markers back to `OpcError::SourceChanged`, as it already
does for `SourceReader`; the current benchmark range adapter returns an I/O
error on a changed source and is therefore not evidence for that typed
mapping. A managed implementation should use the existing
`SourceChangedIoError` path.

No read-ahead state should be installed around the verified-cold filesystem
copy, the aligned EOCD-tail proof, or `source_backed/splice.rs`. The cold
proof depends on its exact 65,536-byte tail request and its phase-local cache
snapshot; splice replay explicitly relies on consuming only the parser's
requested bytes. The provider pilot is a warm logical-source experiment and
must remain a separate named arm.

## Required tests before production consideration

The adapter needs focused tests before any formal performance claim:

* Differential random requests against `OwnedSource`, including overlapping,
  split, nested, empty, EOF, and near-`u64::MAX` offsets. Verify returned
  counts, bytes, and untouched sentinel bytes in the caller output.
* A short-prefix source with a requested offset beyond that prefix. This must
  not return a false cache-hit zero caused by alignment. Also exercise a
  short prefix that begins at the requested offset, a zero return, and the
  64-byte provider cap.
* A fill-size instrumented source proving each physical fill is at most the
  configured window, a same-window request causes no fill, and a miss causes
  one fill. Verify the retained range contains only returned bytes.
* `Interrupted` and cancellation cases. The existing context-aware helper
  owns bounded retries and `Work` charges; no partial or interrupted fill may
  be published.
* A mutable/versioned source that changes before a hit and during a fill.
  Both cases must invalidate the cache and return the typed source-change
  result. An unchanged source must preserve delegated `SourceVersion` exactly.
* Managed budget accounting with an explicit context: a 4 KiB fill that
  returns 1,000 bytes must charge 1,000 accepted `InputBytes`; the following
  cache hit must charge zero additional input; a fill refusal must not leave a
  reservation behind. The test must distinguish the physical fill charge
  from the persistent `Memory` reservation.
* Two concurrent readers of one window. They may serialize under the single
  mutex, but must not observe mixed bytes. Duplicate fills must be prevented
  or counted explicitly.
* The pinned DOCX text oracle, source-version checks, package cache
  diagnostics, physical fill ranges, and logical ranges. The candidate's
  physical media overlap must be visible in the report.
* A negative integration test proving the read-ahead arm is absent from
  verified-cold, aligned-tail, and splice replay entry points.

## Pilot gate

The current tool-layer arm should be accepted only if its actual report shows
the same 200 paragraphs and 10,000 text bytes, unchanged source identity, a
finite fill count lower than 19, and explicit logical-versus-physical ranges.
It must disclose fill/requested-byte amplification, compressed-media overlap,
fixed window capacity, allocator observations, and whole-child RSS. A lower
latency by itself is insufficient. If locality is absent or the physical
amplification is disproportionate, drop the candidate arm rather than widen
the policy.

Only after that pilot gate should a production patch move into `litchi-opc`
with managed-context tests and a new formal before/after evidence bundle. The
existing 0491 artifacts remain the unchanged baseline and must not be edited.
