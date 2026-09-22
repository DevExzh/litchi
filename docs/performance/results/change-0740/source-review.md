# Change 0740 source review: cross-package compressed media ownership

This is a source-only review of the OPC writer, the source-backed PPTX
cross-copy route, and the owned `opened::CrossSlideCopyPlan` route. I did not
run Cargo, tests, profilers, or workloads. The review records an ownership and
compatibility seam; it makes no measured speedup, allocation, or CPU claim.

## Existing source-backed transfer

The source-backed route already has a controlled compressed-payload transfer
boundary for relationship-leaf media. `read_copy_payload` asks the source
`PartView` for both decoded data and an authorized compressed capture when
capture is requested ([source_cross_copy.rs:5830](../../../../crates/litchi-pptx/src/presentation/source_cross_copy.rs#L5830)). The prepared image and chart entries retain that capture, and
`Prepared::into_topology` reauthorizes it before adding the target member
([source_cross_copy.rs:600](../../../../crates/litchi-pptx/src/presentation/source_cross_copy.rs#L600)). Authorization failure can take the existing decoded-payload fallback
where policy allows it.

The token is an OPC-owned `AuthorizedPrecompressedPart`, not a raw compressed
byte handle. It contains a ZIP-reader-verified physical payload, source
artifact and lineage/version, source Part URI and content type, the expected
decoded allocation, and bounded memory reservations
([source_backed.rs:204](../../../../crates/litchi-opc/src/source_backed.rs#L204)). The source side verifies the exact member layout, decoded size, CRC,
source freshness, relationship-leaf condition, content type, and applicable
archive and execution limits before issuing it
([source_backed.rs:6469](../../../../crates/litchi-opc/src/source_backed.rs#L6469)).

`SourceTopologyPlan::try_add_precompressed_part` checks the destination
content type and reserves the target name. Publication emits a fresh canonical
destination wrapper under the remapped target name and feeds the verified
Store/Deflate payload to `new_precompressed_shared`
([source_backed.rs:1664](../../../../crates/litchi-opc/src/source_backed.rs#L1664), [source_backed.rs:8122](../../../../crates/litchi-opc/src/source_backed.rs#L8122)). It does not copy the source local or central directory record wholesale.

## Owned plan provenance seam

The owned plan crosses the package boundary with semantic parts and decoded
payloads. `SlideCopyPart` records only source and target URI, content type,
decoded byte count, and relationship count
([copy_plan.rs:46](../../../../crates/litchi-pptx/src/opened/copy_plan.rs#L46)). Candidate construction then reads the source part and creates a shared
decoded `BlobPart` ([cross_copy_plan.rs:1202](../../../../crates/litchi-pptx/src/opened/cross_copy_plan.rs#L1202)). No source ZIP entry identity, compression method, compressed span, CRC, source artifact, or authorization token survives in that plan.

The ordinary OPC writer can copy a member only when it belongs to the
destination package's own preservation provenance. New copied-closure members
are appended through `regenerated_shared_entry`, which selects Deflate
([pkgwriter.rs:828](../../../../crates/litchi-opc/src/pkgwriter.rs#L828), [pkgwriter.rs:991](../../../../crates/litchi-opc/src/pkgwriter.rs#L991)); `source_blob_retained` is a same-package payload comparison
([pkgwriter.rs:862](../../../../crates/litchi-opc/src/pkgwriter.rs#L862)).
`OpcPackage::exact_source` and `preservation_source` likewise authorize the
package's own source archive, not a foreign source package
([package.rs:1248](../../../../crates/litchi-opc/src/package.rs#L1248)). Reusing
decoded allocations on reopen does not provide compressed-member provenance.

Change 0656 retains the complete serialized candidate to avoid a second
serialization; its first candidate build still deflates the copied closure
([0656-pptx-cross-copy-candidate-budget.md:24](../../0656-pptx-cross-copy-candidate-budget.md#L24)). Change 0662 parallelizes ordinary generated-member Deflate; precompressed entries use a direct framing path and do not enter that wave
([0662-parallel-changed-member-deflate.md:426](../../0662-parallel-changed-member-deflate.md#L426)). Neither change adds a source-to-source compressed transfer for the owned plan.

## Physical revision and durable-patch compatibility

The physical payload cannot be changed silently after a durable plan is
recorded. `apply_plan` checks source and destination semantic and physical
revisions, freshly prepares or reuses the candidate, compares the fresh target
physical revision and patch with the plan, and checks the published archive
revision ([cross_copy_plan.rs:591](../../../../crates/litchi-pptx/src/opened/cross_copy_plan.rs#L591)).

`apply_patch` is stricter for the forward path: it rebuilds the candidate with
`CandidateArchive::Build`, then requires the fresh target physical revision to
equal the durable patch's `target_physical_revision` before publication
([cross_copy_plan.rs:690](../../../../crates/litchi-pptx/src/opened/cross_copy_plan.rs#L690)). A newly selected source compressed stream, changed method/CRC/framing,
or a Deflate fallback that produces different serialized bytes therefore
cannot be silently substituted under an existing durable patch.

Any future owned fast path needs an explicit compatibility rule:

1. Reauthorize the same verified source member and preserve the physical
   payload selected during planning, so the target physical identity remains
   the recorded one; or
2. if the source token is stale, unavailable, over budget, or no longer
   produces the planned bytes, reject and require a re-plan; or
3. select the decoded-payload fallback during planning, record that candidate's
   resulting target physical revision, and use the same deterministic route at
   apply time.

An apply-time fallback that merely preserves semantic graph equality is not
compatible with the durable physical-revision contract.

## Bounded first candidate

The initial candidate should cover binary image leaves only. XML, slide,
relationship, chart, and other closure members remain on their current owned
publication path until they have a separate provenance and physical-identity
proof. The image fast path should be owned by `litchi-opc`, consistent with
[ADR 0011](../../../../docs/adr/0011-ooxml-physical-package-ownership.md):
the format crate may request an authorization, but it should not bind to the
concrete ZIP token type.

The owned API should carry only bounded per-image provenance and, when policy
allows, a bounded verified compressed capture under the existing size and
memory limits. It should not retain a complete foreign `OpcPackage` or pin an
unbounded source archive merely because a durable PPTX plan may be applied
later. A compact source entry identity plus source physical revision can permit
fresh authorization against the source at application time; if that cannot be
done within the budget or freshness boundary, the explicit re-plan/decoded
fallback rule above applies. The source-backed token's source lineage,
freshness, CRC/size checks, archive policy, and reservation accounting provide
the model, while the owned route still needs an owner-specific API and policy
decision.

This review establishes where compressed payload provenance is currently
preserved and where it is discarded. It does not establish a runtime benefit;
any speed or resource conclusion requires a separately frozen and measured
workload.
