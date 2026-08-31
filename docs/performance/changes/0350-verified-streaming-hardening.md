# Change 0350: verified-streaming hardening

Status: correctness and resource-safety evidence only

`performance_claim: none`

## Scope

This change hardens the verified streaming transport rather than changing a
measured hot path. The production scope is shared overreported-read
validation across `ReaderAt` loops, ZIP verification, and streaming, including
the OPC `BorrowedReaderAt` boundary. Offsets and byte counters are checked
before slicing or advancing.

The bounded sink `read_to*` and `read_entry_to*` paths require strict CRC
equality. Ordinary owned reads retain their documented zero-CRC compatibility.
For borrowed stored access, a nonempty zero-CRC member returns `None`, allowing
the caller to use the owned fallback. If Deflate produces extra output,
`InvalidSize` is returned before any accounting-overflow error can replace it.

## Bounded streaming boundary

Verified transfer uses one fixed-size scratch buffer for one active member. The
statement excludes caller-owned source/archive/index data, the destination
sink and its output retention, caches, and aggregate process memory. It is not
a total RSS or whole-transaction bound. The hardening preserves typed
short-read, zero-write, overreported-read/write, CRC, declared-size, source,
and partial-output failure behavior; transient `Interrupted` reads are
retried.

## Validation evidence

The final successful commands were:

- `cargo fmt --package soapberry-zip --package litchi-opc` succeeded.
- `cargo test -p soapberry-zip --lib -- --test-threads=1`: `287/287`.
- `cargo test -p litchi-opc --lib overreport -- --test-threads=1`: `4/4`, with `261` filtered.

All validation used one build job, incremental/debug compilation disabled,
one test thread, one dedicated disk target, and an 8 GiB `ulimit`. Only these
final successful commands are recorded; transient compile-feedback iterations
are excluded. Independent source-only reviews found no P0/P1 after the
compatibility correction.

## Claim boundary

This record makes no latency, throughput, RSS, allocation, syscall,
decompression, or concurrency claim. It adds no performance selector or raw
result artifact. The evidence establishes bounded validation and failure
semantics only; `performance_claim: none` remains authoritative.

