# 0661: owned OPC packages defer part payload decoding until fallible access

Status: retained implementation. `performance_claim: none`. This change
implements the accepted [ADR 0030](../adr/0030-lazy-opc-part-decode.md), the
owner decision in [0652](0652-owner-decisions-for-the-third-wave.md), and row
3 of [0651](0651-queue-refresh-after-the-second-wave.md). It completes the
frozen design in [0610](0610-opc-lazy-part-decode-design.md).

The implementation branch was based at `ab07e2a47` on
`perf/0661-opc-lazy-part-decode`. The coordinator should cherry-pick this
commit before the overlapping 0665 change; 0665 is currently represented by
`37d402379`. If both changes touch OPC publication or source provenance,
resolve the union of the two implementations and rerun the OPC suite.

## What changed

Owned ingress retains one source archive and creates a deferred payload handle
for each typed part. The shared `DeferredPartSource` builds the ZIP index at
most once. Each part's `OnceLock` records either its decoded `Arc<Vec<u8>>` or
the first typed failure. Borrowed `from_bytes*` ingress, explicit eager reader
paths, streamed sources without a retained archive, and authored parts keep
the ready payload representation.

The public fallible accessors force the payload before returning a part:
`get_part`, `get_part_mut`, `main_document_part`, and `part_by_reltype`. The
infallible `iter_parts()` route now yields `PartMetadata`, which exposes the
part name, content type, relationships, and decoded-state observation without
a payload accessor. `try_iter_parts()` yields `Result<&dyn Part>` and is the
fallible byte-bearing iteration route. Internal publication code uses the
undecoded iterator and payload provenance so it can preserve untouched source
members without inflating them.

`Part::blob()` and `Part::blob_arc()` retain their public signatures. Their
implementation is observation-only: they never perform I/O or decompression.
The package invariant is that a `&dyn Part` leaves `litchi-opc` only after
`ensure_payload()` succeeds. Internal code that needs bytes must use a
fallible accessor first. This keeps the existing public trait while ensuring a
decode failure cannot be converted into an empty payload by a hidden retry.

The migration changed payload-reading iteration sites to
`try_iter_parts()` or an explicit fallible part lookup. Metadata-only scans
remain on `iter_parts()`. Optional part probes were audited across the format
crates: only `OpcError::PartNotFound` means that an optional part is absent;
ZIP, I/O, allocation, cancellation, and limit errors propagate. Boolean
source-match helpers remain deliberately conservative where their API is
infallible: any payload error returns `false`, which rejects the exact-source
proof and cannot publish an empty replacement.

## Budget and failure behavior

The central-directory declared-size preflight remains at open. Actual
`PartBytes` and aggregate `TotalPartBytes` charges are applied when a deferred
member is first decoded. A failed `OnceLock` is never retried, so repeated
access returns the same refusal. This preserves typed limit, allocation,
cancellation, I/O, and ZIP/CRC failures while moving only the actual-payload
check for untouched deferred members to first access.

The infallible observation fallback is intentionally not a publication source.
For an unforced or failed deferred member, `decoded_blob()` is `None` and the
writer uses retained source-member provenance. A changed member must first be
obtained through a fallible accessor, so it cannot be changed after a decode
failure. The failed member may be copied as opaque source bytes when it is
untouched; reading that output still produces the same failure. It is never
published as a newly generated empty member.

There are two intentional publication contracts:

* An exact source no-op is authorized to copy the source archive byte-for-byte
  without decoding every member. That includes a known-corrupt member: the
  no-op preserves the caller's exact source and does not claim that the member
  is valid.
* A materialized or changed publication must force the selected member and
  validate its payload before output. A targeted edit can copy an untouched
  failed member's raw source record, but it must regenerate and validate the
  edited member; later access to the copied failed member still refuses.

This separation is what prevents the old `bytes()/arc()` empty fallback from
turning a corrupt first access into valid-looking output while retaining the
exact no-op guarantee required by ADR 0030 and ADR 0006.

## Deterministic evidence

The evidence packet is [results/change-0661](results/change-0661/README.md).
It contains the replay probe, the eager-versus-lazy differential, cross
checkout outputs, and read-set counters. These are correctness and work-set
measurements only; no timing, RSS, allocation, or registered performance
claim is made.

The retained 336-fixture read-set corpus contains 5,077 admitted parts and
45,562,463 declared inflated bytes across 334 successfully opened fixtures;
two fixtures are refused during open. The exact-no-op route inflated zero
parts and zero bytes, with 334 publications and the same two open refusals.
The one-part reblob route inflated 334 parts and 389,581 bytes for its 325
published fixtures; nine signed fixtures were refused by explicit policy.

The XLSX hide route is reported by its actual semantic path rather than by
the older design estimate. Its 33 published XLSX fixtures touched 216 of
550 parts and 6,044,263 of 8,274,037 inflated bytes. The full 336-row route
has 301 operation refusals and two open refusals, so its all-row decoded total
is not presented as a working-set percentage.

The eager-versus-lazy differential has 1,342 `MATCH` rows and zero
`MISMATCH` rows. The before/after cross-checkout artifacts compare equal for
OPC member inspection, exact no-op publication, one-part reblob, and XLSX
hide (`cmp` exit 0 for each operation). The three focused lazy tests cover
cold metadata/no-op access, selected-part-only decode, and stable corruption
failure; all pass.

## Validation

The branch passed the following checks with `CARGO_TARGET_DIR=/tmp/litchi-0661-target`
and at most two Cargo jobs:

* `cargo fmt --all`
* `cargo check --workspace -j2`
* `cargo test -p litchi-opc --test lazy_part_decode -j2` (3 passed)
* `cargo test -p litchi-opc -j2` (437 unit tests, all integration tests, and
  5 doctests passed)

The source audit also ran `git diff --check` clean. The coordinator should
rerun the shared integration gate after cherry-picking 0661 and 0665 together,
because both changes touch OPC publication/source-preservation code.

## Limits and retained gaps

No timing or peak-memory claim is registered here. The read-set counters show
which payloads were inflated; they do not establish latency, RSS, allocator
traffic, or end-to-end speed. The retained source archive and per-part lazy
metadata add bookkeeping for a package whose every payload is eventually
read. Declared-size preflight remains eager, while actual-size checks for
deferred members are first-access checks as authorized by ADR 0030.

The current XLSX semantic hide path reads more than the one-part estimate in
0610; that estimate belongs to the frozen design and is not restated as an
implementation result. Signed, malformed, non-XLSX, and noncompact fixtures
retain their typed refusal outcomes. Exact no-op copying of a corrupt source
is intentional preservation behavior and must not be interpreted as a
validation result.

## Disposition

The implementation is retained for cherry-pick. It satisfies the accepted
lazy-decode seam, the accessor/error-channel invariant, the migration audit,
the exact-source publication contract, and the scoped differential evidence.
The packet withholds any performance claim and leaves broad release-mode
timing/RSS work to a separately authorized measurement.
