# Change 0349: PhysPkgReader stored-Part borrowed consumer

## Status

Correctness and ownership evidence recorded; no performance result.

`performance_claim: none`

## GOAL alignment

This change aligns with `docs/GOAL.md:398` by allowing an immutable-slice
`PhysPkgReader` to hand an eligible stored OPC Part to an immediate consumer as
a validated source-backed `&[u8]`. The borrow is confined to the reader and
consumer lifetime; it is not retained in an ordinary `Part`, facade model, or
source-backed positional owner.

The Part and archive limits are checked before the borrow is published. The
stored path still performs the complete CRC and ZIP local/central layout scan.

## Exact ownership change

For an eligible immutable-slice Store member, the borrowed consumer avoids the
destination `Vec` allocation and `memcpy` required by the legacy owned read.
It also avoids a Part materialization budget charge and decompression-cache
entry because no payload storage is materialized. This is a logical ownership
and copy-elision result; it is not a claim about allocator traffic or process
memory.

CRC verification, declared-size checks, local/central metadata checks, and
borrowed-span layout validation remain in the operation. No validation step is
removed in exchange for the borrow.

## Safety refinement and fallbacks

Encrypted Store and Deflate members produce typed errors before an owned
fallback is selected. A nonempty member with a declared CRC of zero returns
`None`, preserving the legacy owned-read fallback rather than publishing an
unverified source slice.

Deflate members and ZIP64-EOCD archives use the owned fallback. Generic
`ReadAt`, file-backed, remote, and other positional sources do not acquire a
borrowed lifetime. The existing private structural-member borrowing path is
not expanded into a content-type or signature claim.

## Validation evidence

The crate-scoped formatting check `cargo fmt --package soapberry-zip --package litchi-opc -- --check` passed after formatting.

- `litchi-opc` integration filter: `8/8`.
- Lower borrowed filter: `10/10`.
- Full `soapberry-zip` library result: `281/281`.
- Jobs were serialized with `CARGO_BUILD_JOBS=1`, `test-threads=1`, and an
  8 GiB process ceiling.

The evidence is bounded correctness and ownership evidence. No raw result
artifact is required.

## Claim boundary

No timing, RSS, throughput, physical-I/O, decompression, bytes-copied, or
allocator claim is made. Existing mixed stored-member corpora are too weak to
authorize timing evidence, and all-Stored synthetic OOXML must not be treated
as producer representation. `SourceBackedPackage`, generic `ReadAt`, file,
remote, Deflate, and ZIP64-EOCD borrowing remain outside this change.
