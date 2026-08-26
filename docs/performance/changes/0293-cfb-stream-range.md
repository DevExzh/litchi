# Change 0293: CFB bounded stream ranges

Date: 2026-08-27

Status: Accepted bounded API and resource behavior

`performance_claim: none`

## Decision

`litchi-cfb::OleFile` now exposes `read_stream_range`, which resolves one
stream path and reads a caller-selected logical range into caller-owned
storage. The range is checked with checked arithmetic against the directory
entry's declared stream length before payload I/O. An empty range at exact EOF
is a metadata-only success; an empty range beyond EOF and every overflowing or
out-of-bounds range return typed invalid-data errors.

The direct reader follows only the bounded FAT or MiniFAT traversal required by
the requested range. It keeps scalar chain state, validates table and marker
indexes, and coalesces only physically contiguous spans whose logical output
positions are adjacent. It does not allocate a chain vector, visited set, or
range-sized temporary payload. MiniFAT reads map through the already validated
root-chain index and never populate the existing lazy `ministream` cache.

The reader uses absolute seeks against the captured `file_size`. Existing CFB
behavior for a tolerated truncated final physical sector is preserved: present
bytes are read and the remainder of the caller range is zero-filled, while no
bytes appended after the captured source length can be observed. Actual seek
and read failures remain `OleError::Io`; standard exact-read interruption retry
semantics are retained.

The caller owns the destination and must discard it after any error. A late
source error can follow earlier accepted bytes; the operation does not roll the
destination back. The existing whole-stream `open_stream` path and its
MiniStream cache behavior are unchanged. This API has no source freshness or
cancellation mechanism because `OleFile` owns a seekable reader without the
positional-source contract used by `SharedOleFile`.

## Scope and claim boundary

The bounded-memory statement is relative to stream length: the range path
retains no complete stream payload, chain collection, or MiniStream cache in
addition to the already retained parsed CFB index and caller-provided output.
The existing FAT and MiniFAT tables, directory graph, and caller buffer remain
in scope. The statement is an API/resource invariant, not a measurement.

No claim is made for physical I/O, syscall count, disk or filesystem cache,
allocator traffic, RSS, latency, throughput, speedup, or consumer-level
benefit. The record does not change or characterize `SharedOleFile`, repeated
range caching, whole-stream materialization, or source freshness/cancellation.
It also does not infer physical-read locality from the logical span
coalescing rule.

Focused tests cover fragmented FAT and MiniFAT order, root-chain mapping,
truncated-sector zero fill, empty and exact-EOF ranges, bounds and path/type
preflight, malformed short/invalid/cyclic metadata with bounded termination,
absolute cursor-independent reads, payload locality, interrupted-read retry,
and injected seek/read failures. Malformed states rejected during ordinary
CFB open remain represented by private synthetic-index tests where needed.
