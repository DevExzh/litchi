# Cross-source compressed-part transfer

This is a source-only design record for change 0430. The reviewed binary and
production sources are unchanged. The instrumented 0430 FP control profiles
attribute the publication cost to the Deflate path, but this document does not
claim a candidate speedup, memory result, or correctness result. A candidate
requires a later implementation and a new frozen control capture.

## Finding

Unchanged source media and chart parts are decoded while the cross-source PPTX
plan is prepared and are Deflate-compressed again when the destination package
is published.

The path is:

| Stage | Current representation and operation |
| --- | --- |
| PPTX planning | `source_cross_copy.rs::prepare` calls `PartView::data()` for an image or chart. `PreparedImage::bytes` and `PreparedChart::bytes` retain the decoded bytes in an `Arc<Vec<u8>>`. Images are size-checked; charts are also parsed and validated as XML. |
| Topology | `Prepared::into_topology` calls `SourceTopologyPlan::try_add_part_shared`. `TopologyPartAddition::payload` is decoded part content. |
| OPC append | `source_backed.rs::write_topology_to_stream` turns every addition into `RegeneratedEntry::new_shared(...).compression_method(Deflate)`. |
| ZIP publication | `preserve.rs::generated_entry` uses `flate2::DeflateEncoder` to construct the new member before the preservation writer appends it. |

The plan's `reuse_or_clone_payload` can reuse the decoded allocation on a
plan recheck, but it has no compressed source representation. Destination
members that are untouched still use preservation `Copy` and retain their raw
ZIP records. The recompression applies to the newly added source media and
chart members, which explains the recovered `Deflate` plus
`CountingSink` publication ancestry in the 0430 profile.

## Smallest safe change

The first implementation should transfer a verified compressed payload while
keeping the existing decoded validation and source checks. It should not add a
public constructor that accepts arbitrary compressed bytes, a CRC, and two
sizes. Those values are a claim about an archive member and must be issued by
the ZIP reader after it has checked the member.

The bounded API can be shaped as follows; the names are illustrative and the
types should remain private to the owning crates until the ownership boundary
is settled:

```text
VerifiedPrecompressedPart {
    method: Store | Deflate,
    crc32: resolved CRC metadata,
    compressed_size: u64,
    uncompressed_size: u64,
    compressed: Arc<Vec<u8>>,
}
```

The physical proof is owned and issued by `soapberry-zip`; it need not know
litchi's process-local package lineage. `litchi-opc` should wrap it in a
private package-authorized token that carries the source artifact, lineage, and
version guard before handing it to a topology plan. The conceptual
`source_guard` therefore belongs to that wrapper, rather than becoming an
unchecked field on a ZIP value:

```text
AuthorizedPrecompressedPart {
    physical: VerifiedPrecompressedPart,
    source_guard: private source identity/version authority,
}
```

The ZIP issuer obtains the raw compressed range only after
the existing strict local-header/central-directory/layout checks have bounded
the range. It must associate the payload with the resolved CRC (including a
data-descriptor CRC when a descriptor supplies the authoritative value), the
method, and both sizes. The payload is admitted to the transfer API only after
the existing Store/Deflate reader has established the decoded size and CRC
under the archive's current verification policy. A lower-level raw-range
helper may be useful inside the ZIP crate, but a raw range by itself is not a
verified transfer token.

The least invasive first cut can perform the current decoded read and then a
bounded compressed read into the `Arc`, but the second read is not itself
proof that the retained bytes are the bytes that were decoded. Before issuing
the token, the immutable captured bytes must be decoded and strictly verified,
and their decoded output must compare byte-for-byte with the already validated
expected `Arc<Vec<u8>>`. The capture must also have source freshness checks
before and after the capture and validation. A one-pass reader can instead tee
the same strictly decoded bytes into the retained compressed buffer. Either
form removes the destination Deflate pass while preserving the existing
image/chart and candidate validation; a later one-pass implementation can
avoid the extra source traversal if its proof is equivalent.

Capture validation must compute the actual CRC of the retained compressed
bytes' decoded output. A nonzero authoritative central or descriptor CRC must
match that value. Under the current compatibility policy, a zero expected CRC
may omit the expected-value comparison, but the token still records the
computed actual CRC for destination framing. A source mutation, capture read
failure, size mismatch, CRC mismatch, or decode failure is a typed failure;
it must not fall back to an unchecked compressed payload.

The preferred first candidate is to issue this authorized token only during
publication, after the existing `current.matches(plan)` and `verify_candidate`
gates have passed. If the source flow permits, pass the existing prepared
source part and its expected decoded `Arc` to a source-backed helper such as
`try_add_part_from_source(target, part_view, expected_decoded)`. Keep the plan
and `PreparedImage`/`PreparedChart` equality and cache fields unchanged, and
retain the token only in the ephemeral publication topology. A media-only
first change is the narrowest candidate; chart transfer can follow once the
same source-proof path is shown to preserve its XML validation boundary.

`soapberry-zip::RegeneratedEntry` then needs a second payload kind, for example
`Precompressed(VerifiedPrecompressedPart)`, and a constructor with no unchecked
metadata path:

```text
RegeneratedEntry::new_precompressed_shared(
    name,
    verified_part,
)
```

Its prepared-entry path should call the existing
`ZipArchiveWriter::write_precompressed_file_with_accounting` framing operation
or an equivalent direct framing path. It must emit a fresh local header and
central record for the destination offset and target name, then copy the
verified compressed bytes. The existing logical `new_shared` path remains for
XML, relationship, and other generated content. Initially, the one-entry
preparation can copy the compressed bytes into its staging buffer; a later
prepared-entry representation can retain the `Arc` through framing if the
extra copy is measurable.

At the OPC boundary, add a payload variant to the existing topology addition,
conceptually:

```text
TopologyPartPayload::Decoded(Arc<Vec<u8>>)
TopologyPartPayload::Precompressed(opaque authorized source part)
```

`SourceTopologyPlan::try_add_precompressed_part` should accept only the opaque
OPC-authorized value and the target part metadata. It should apply the same
URI, content-type, duplicate, operation-count, relationship, and collision
checks as `try_add_part_shared`. `write_topology_to_stream` maps the physical
part in this variant to the precompressed `RegeneratedEntry` after the source
guard has been checked; generated XML and relationship additions continue to
use Deflate. PPTX retains the decoded `bytes` for its current logical checks
and carries the authorized token alongside them.

This keeps physical ZIP ownership in `soapberry-zip`, OPC topology ownership in
`litchi-opc`, and PPTX dependency-closure and content validation in
`litchi-pptx`. PPTX should not receive ZIP entry IDs or raw local/central
record spans.

## Authorization and publication checks

The token must be source-authorized, not merely content-shaped. The current
PPTX plan already records source lineage and source version, rejects a same
lineage/version copy where required, reruns preparation, calls
`current.matches(plan)`, and verifies the candidate before publication. The
compressed token must remain under those gates.

At minimum, a prepared image or chart should retain:

* target URI, content type, decoded logical bytes, and the existing declared
  and observed size checks;
* the source lineage and source version (or a private source artifact that
  carries both), plus the verified compressed token;
* the method, resolved CRC, compressed size, uncompressed size, and compressed
  bytes, or an equivalent immutable token digest.

`matches` must reject a token from another source package and a token whose
source version changed. Comparing only decoded bytes or only a URI would allow
an equal logical payload to stand in for a different physical source artifact.
The existing `SourceCheckedWriter` source checks before and after sink writes
remain necessary: they fence a source mutation during publication after the
plan has been accepted. If this capability becomes a generic OPC API, the
same source guard must be checked by the OPC writer; otherwise keep the
operation format-owned and private to the source-backed publication path.

The first cut should keep `verify_candidate`'s current decoded image and chart
reads, including chart XML validation. A later optimization may validate the
token against a retained source snapshot and avoid a duplicate decoded read,
but that is a separate proof obligation. All destination freshness, topology,
collision, relationship-closure, and atomic-output checks remain in force.

The source decode currently accepts only the Store and Deflate methods. The
transfer path should preserve that restriction and return the existing typed
unsupported-method or source-read refusal for other methods. Central directory
declarations are attacker-controlled until the entry has been read through the
strict path. A nonempty zero CRC follows the current compatibility semantics,
where a zero expected CRC is not independently checked; the implementation
must compute and record the actual CRC during capture validation rather than
marking the token as unchecked or absent. It must not introduce a new refusal
solely because a compatible source has a zero expected CRC. A descriptor's CRC
must not be replaced with the central declaration when the strict reader has
established the descriptor as authoritative.

The destination writer emits canonical local and central metadata and does not
need to retain the source data descriptor. It also does not preserve source
timestamps, extra fields, flags, local-header bytes, or central offsets. This
is exact compressed-payload reuse, not exact member or framing preservation.

## Limits and memory

The ZIP index already bounds compressed member size, uncompressed member size,
and aggregate uncompressed size. The new capture must apply those limits
before allocating and check every `u64` to `usize` conversion. The topology
limits currently charge decoded payload length; a precompressed addition must
charge both the logical size and compressed staging size. The aggregate budget
must cover the retained decoded bytes plus the compressed `Arc` and any writer
staging copy.

The output bound must include compressed payload bytes, canonical headers,
central records, and possible ZIP64 records. `payload.len() + 4096` is not a
sufficient bound for a precompressed addition. All addition preparation,
limit checks, metadata checks, and source-freshness checks should complete
before the first sink write so a rejected token cannot leave a partial output.

Store remains Store and Deflate remains Deflate in the new wrapper. The ZIP64
writer already selects ZIP64 metadata when compressed or uncompressed sizes
need the sentinel representation; the transfer path must pass the actual
compressed size and exercise both size and offset/entry-count boundaries.

The first cut may increase peak resident memory because decoded bytes are
retained for logical validation while compressed bytes are retained for
publication. The later one-pass capture and direct prepared framing are
optimization candidates, not reasons to weaken the validation boundary. RSS,
source traversal count, and output staging should be measured in the next
batch after the frozen instrumented FP control baseline is captured.

If a valid budget or unsupported-optimization condition prevents the
precompressed path, the operation may use the existing validated logical
recompression path when that fallback is within the configured limits. Budget
accounting must cover the fallback before output begins. A compressed-source
read failure, source mutation, CRC or size mismatch, decode failure, or other
corruption must never trigger that fallback; those conditions return their
typed errors.

## Preservation meaning

Cross-source `PreservationIndex::Copy` cannot be used for this operation. Its
entry IDs, local spans, central offsets, and source framing belong to one
archive. A source member can contribute its compressed data to a new target
member only after the target wrapper is generated.

The proposed path therefore preserves the compressed payload bytes and method
while canonicalizing the destination wrapper. It leaves untouched destination
members on the existing raw-copy path. It does not authorize copying a source
local header, descriptor, central record, or arbitrary opaque member. Current
rejections for encryption, signatures, macros/VBA, unsupported physical
members, active protection, trailing source bytes, and non-leaf media/chart
parts remain unchanged. Dependency-closure, destination collision, dialect,
theme/layout, and relationship checks remain PPTX responsibilities.

Whether a canonical destination wrapper is acceptable under the requested
Preserve semantics is a gate that must be resolved before implementation. If
Preserve requires source timestamps, extras, flags, descriptor framing, or
other member-level details, payload-only transfer is insufficient and the
design must instead specify an explicitly authorized framing transfer. No
precompressed implementation should proceed while that preservation decision
is open.

This is consistent with ADR 0005's requirement for an exact source artifact
and unrevoked authorization, ADR 0006's preserve-versus-normalize boundary,
ADR 0003's dependency-closure transfer rule, and ADR 0011's ownership of
physical OPC packaging by `litchi-opc`.

## Required tests before implementation is accepted

These are design requirements, not tests run for this review.

**ZIP.** Add Store and Deflate precompressed append/reopen cases. Extract the
new member's compressed range and compare it byte-for-byte with the source
payload; verify method, resolved CRC, decoded size, and uncompressed size.
Exercise a source data descriptor, local/central mismatch, truncated or
overlapping range, Store size mismatch, unsupported method, and corrupted
compressed bytes. Verify that invalid tokens fail before writing. Check that
the new wrapper has correct canonical metadata and no source descriptor
dependency. Cover ZIP32 and ZIP64 size, offset, and entry-count boundaries,
including output accounting and a partial sink.

**OPC.** Add a precompressed topology addition test that accepts only a
source-issued token with the expected lineage and version. Verify that target
untouched members, content types, relationships, and ordering retain existing
preservation behavior while the added member has the exact transferred
compressed payload under a new wrapper. Mutate the source before publication
and after accepted sink bytes and require the existing typed stale/incomplete
output errors. Exceed compressed and logical limits and require refusal before
the sink receives bytes. Ensure there is no public arbitrary compressed-payload
constructor.

**PPTX.** Retain the current one-image, multiple-image ordering, shared-image
deduplication, aggregate reservation, read-error, actual-size-mismatch,
media-closure, chart-validation, stale-source, foreign-source, and partial
sink cases. Add Store and Deflate source fixtures and assert that copied image
and chart compressed ranges equal their source payloads while logical part,
relationship, content-type, closure, and unknown-markup checks still pass.
Include descriptor and ZIP64 fixtures where the corpus can provide them.

## Open decisions and next evidence

The following decisions are intentionally unresolved until the implementation
is scoped and measured:

* whether the first release uses the safe two-pass capture or a one-pass tee;
* whether a precompressed entry is exposed as a generic OPC capability or kept
  private to the source-backed PPTX path;
* whether the initial generated-entry staging copy is acceptable or direct
  prepared framing is needed immediately;
* whether payload-only preservation is sufficient for the requested preserve
  semantics. This is a required decision before implementation; exact
  descriptor, timestamp, extra-field, and flag preservation is out of scope
  for the payload-only variant.

The 0430 instrumented FP control profile is recovered attribution for the
unchanged binary. It does not establish correctness or performance of the
proposed change. After the control evidence is frozen and the wrapper
preservation decision is resolved, implement the smallest source-issued token
path, run the focused correctness cases above, and capture matched control,
allocation, RSS, and native evidence with the same lifecycle boundaries.
