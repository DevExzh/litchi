# 0840 CFB physical-emission-order candidate

This archive contains an unapplied candidate for the current base
`f43040e1da`. The production candidate changes only
`crates/litchi-cfb/src/writer/core.rs`: it emits the regular-sector ministream
immediately after the already allocated large-stream chains. Allocation order,
directory construction, range-lock clearing, MiniFAT/FAT generation, and DIFAT
handling are unchanged.

The allocator reserves large-stream sectors before the regular-sector
ministream. On a fresh seekable `Cursor<Vec<u8>>`, the previous emission order
therefore sought past the current end before the first large payload and made
the cursor initialize a hole covering the large payload. The candidate makes
the first data write follow the allocation order, so the header is followed by
the large payload and then the ministream. The later directory, MiniFAT, FAT,
and DIFAT writes remain in their existing order.

`candidate.patch` adds the focused integration artifact
`crates/litchi-cfb/tests/writer_emission_order.rs`. It covers both 512-byte
version-3 and 4096-byte version-4 files containing one MiniFAT stream and one
multi-sector FAT stream. A `Write + Seek` wrapper around `Cursor<Vec<u8>>`
counts seeks, writes, and every write that begins beyond the current cursor
length; the candidate requires zero such writes and compares the resulting
bytes with a same-model golden serialization. The test also reverses stream
insertion order and reopens both streams to keep allocation, directory, and
payload semantics under one assertion.

The same integration artifact also round-trips zero-length, 4,095-byte, and
4,096-byte `Tiny` streams in both sector versions and both insertion orders.
That keeps the MiniFAT cutoff and empty-stream transitions in the focused
regression surface without changing the candidate's allocation policy.

The same artifact overwrites an exactly sized sink prefilled with nonzero
bytes and compares every output byte with the golden serialization. An
injected short write followed by one `Interrupted` result must be retried to
completion for both sector versions and produce the same golden bytes. A
separate partial sink failure is required to return `OleError::Io` with the
original I/O kind/message. These checks keep caller-owned sink behavior and
typed error propagation visible while the physical write order changes.

The candidate deliberately does not alter or claim the version-4 range-lock
boundary or DIFAT policy. Existing range-lock and DIFAT tests remain the
authority for those limits. The integration test does not require a hardcoded
whole-file byte array; its golden is the current writer's same logical model,
while root-owned before/after capture must provide the independent byte
equality gate.

## Review constraints

The candidate is confined to the CFB writer owner and adds no public API,
dependency, ambient I/O, runtime, unsafe code, or retention policy. It follows
ADR 0005's measured seekable-output and caller-owned sink contract, ADR 0006's
deterministic and preservation requirements, ADR 0008's before/after evidence
gate, and ADR 0024's current `litchi-cfb` ownership boundary.

No build, test, benchmark, or application of this patch was performed in this
delegated preparation step. Root should first establish the current-base
regression result, then apply the patch in its isolated candidate state and
run the affected CFB/DOC quality gates plus the focused integration test.
