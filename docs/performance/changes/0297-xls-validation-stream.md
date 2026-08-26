# Change 0297: XLS streaming validation

Date: 2026-08-27

Status: Implemented deterministic bounded validation behavior

`performance_claim: none`

## Decision

XLS validation no longer materializes the complete logical `Workbook` or
`Book` stream before checking its BIFF grammar. After the existing CFB and
declared stream-size preflight, the private validator creates one
`SharedOleStreamCursor` and reads one four-byte BIFF header followed by its
declared payload. The payload is copied into one reusable
`litchi_biff::MAX_RECORD_BYTES` (8,224-byte) scratch buffer. No per-record
buffer, record history, or whole-stream `Vec<u8>` is retained.

The streaming frame reader preserves `litchi_biff::Records` ordering and
classification: record-count limits are checked before the next header,
payload limits are checked before truncation, short headers and declared
short payloads become the existing BIFF-invalid report, and cursor failures
inside declared ranges remain ingress errors. Logical offsets are tracked as
`u64` values. Existing BOF/EOF ownership, BoundSheet8 inventory, protection,
external-reference, and FILEPASS analysis consume each frame immediately.

When a legal FILEPASS record is encountered, its own bounded payload is read
and analyzed, then validation stops before ciphertext or trailing bytes. A
directory-proven encrypted container skips clear BIFF payload reads. Cursor
reads retain the shared source-freshness fences, and report paths retain their
final CFB source-version check.

Encryption presence is fail-closed when a clear scan stops before full logical
Workbook-stream traversal or complete FILEPASS: without independent directory evidence, the
encryption check is stopped by the Workbook/BIFF dependency rather than being
reported as not applicable. A FILEPASS header whose payload is over the BIFF
bound or truncated still records encryption presence and the existing
`xls.encryption.filepass_invalid` issue before the scan stops.

Full clear semantic validation may still read the complete logical Workbook
stream, but it no longer retains that stream. The bounded worksheet inventory,
owner map, and semantic collector state remain governed by the existing XLS
limits and are not replaced by unbounded record history.

## Scope and claim boundary

The bounded-memory statement is an implementation invariant: at most the
8,228-byte header-plus-payload framing scratch is dedicated to one BIFF frame,
in addition to limit-governed semantic metadata and the already retained CFB
index. The caller-owned archive/source and CFB allocation tables remain in
scope.

`performance_claim: none` means this change makes no claim about RSS,
allocator traffic, copy volume, syscall count, physical I/O, cache behavior,
latency, throughput, decompression, or whole-transaction memory. It does not
change the public validation API, add cancellation, alter CFB range behavior,
or characterize source-backed worksheet queries.

Focused tests cover exact and oversized BIFF payload limits, exact/under/over
record-count ceilings, one-, two-, and three-byte tails, truncated payloads,
and a large ciphertext tail observed through a bounded positional source. The
existing malformed and misplaced FILEPASS cases remain parity coverage.
