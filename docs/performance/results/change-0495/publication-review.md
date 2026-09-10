# Change 0495: publication and inverse review

Status: bounded read-only review of the ordinary DOCX source-backed publication
path and the frozen 0495 managed-edit harness.  No production source, harness,
gate, build, or capture was changed or run for this review.  The protected
`/home/zhuhe/code/litchi-spec-gaps` worktree was not inspected or modified.

The review uses the requirements in [`docs/GOAL.md`](../../../../docs/GOAL.md),
accepted ADR 0005 (source ownership, finite execution policy, sequential sinks,
and typed partial output), ADR 0006 (preservation and signature boundaries),
and ADR 0008 (authenticated verification).  The central distinction is that a
document XML patch and a complete package-artifact inverse authorize different
operations.

## Result

The managed changed-publication path and the complete-artifact inverse have the
right source and sink proof structure.  The frozen harness's patch oracle is
valid as a document-XML oracle, but it does not exercise or measure the new
complete-artifact inverse API.  The previous unmanaged identity gap is closed
for the ordinary source-backed document commit path: unmanaged snapshots now
retain an identity-only source token, and the scoped equal-XML/different-opaque
regression refuses a foreign commit before output.  The resource `after_drop`
gauge is sampled after package consumption, returned-snapshot drop, and commit
drop; only the `cache_live` label remains earlier than its name implies because
it is captured before publication.

## Findings

### 1. The managed forward path preserves source authority

`Package::publish_document_commit_to_stream()` captures the current main
document, applies the patch against that exact snapshot, rejects operations that
need package dependency publication, and selects the source-authorized branch
only when the changed target carries a `SourceXmlPart`
([`source_backed.rs`](../../../../crates/litchi-docx/src/source_backed.rs#L1293)).
For a managed edit, `main_document_snapshot(..., true)` retains a source XML
owner and `target.source_xml()` is passed to
`SourceTopologyPlan::try_replace_source_xml_part()`
([`source_backed.rs`](../../../../crates/litchi-docx/src/source_backed.rs#L1736),
[`source_backed.rs`](../../../../crates/litchi-opc/src/source_backed.rs#L1515)).
The topology publisher checks the source lineage, source version, Part URI,
content type, original Part bytes, source freshness, limits, signatures, and
physical preservation plan before it starts output
([`source_backed.rs`](../../../../crates/litchi-opc/src/source_backed.rs#L6419)).

This avoids the unsafe `shared_xml()` escape for a changed managed document.
The fallback to an owned/shared XML overlay is the compatibility path for an
unmanaged snapshot.  Signed changes, opaque topology that the preservation
primitive cannot carry, and unsupported dependency operations remain typed
refusals before output, which is consistent with the fail-closed preservation
contract.

### 2. Ordinary and inverse partial-sink handling is source-checked

The ordinary publication path delegates output to OPC preservation writers that
count accepted bytes and run source/context fences around the stream.  The
final source decision preserves a typed `IncompleteOutput { written, ... }`
when a sink or source/context error follows an accepted prefix
([`source_backed.rs`](../../../../crates/litchi-opc/src/source_backed.rs#L10840),
[`source_backed.rs`](../../../../crates/litchi-opc/src/source_backed.rs#L9410)).
The inverse path has the same shape: it authenticates both complete artifacts
before the first byte, checks both sources while copying, counts accepted
output, and converts a later source or sink failure to a typed partial result
([`artifact_restore.rs`](../../../../crates/litchi-opc/src/source_backed/artifact_restore.rs#L32),
[`splice.rs`](../../../../crates/litchi-opc/src/source_backed/splice.rs#L1659)).

`DocumentPublicationWriter` hashes only the bytes accepted by its wrapped sink,
so a returned publication proof cannot describe an unaccepted suffix
([`source_backed.rs`](../../../../crates/litchi-docx/src/source_backed.rs#L185),
[`source_backed.rs`](../../../../crates/litchi-docx/src/source_backed.rs#L267)).
The with-inverse API returns a handle only after ordinary publication succeeds;
partial publication therefore cannot accidentally become an inverse-authorized
publication.  The focused managed tests cover changed ordinary publication
failure, reopened inverse failure, source mutation, cancellation, output
limits, and a foreign artifact with the same main-document XML but different
opaque members
([`source_backed_managed_document_edit.rs`](../../../../crates/litchi-docx/tests/source_backed_managed_document_edit.rs#L344),
[`source_backed_managed_document_edit.rs`](../../../../crates/litchi-docx/tests/source_backed_managed_document_edit.rs#L366),
[`source_backed_managed_document_edit.rs`](../../../../crates/litchi-docx/tests/source_backed_managed_document_edit.rs#L462)).

The 0495 timed rows should still be treated as successful-output rows.  They do
not establish sink atomicity by themselves; partial sink behavior belongs to
the focused production tests and a separate failure receipt if it is reported.

### 3. The complete-artifact inverse is distinct from `Patch::inverse()`

`Patch::inverse().apply()` swaps the transaction's before/after document
snapshots and checks the exact main-document snapshot identity
([`transaction.rs`](../../../../crates/litchi-docx/src/document/transaction.rs#L4218)).
It does not retain ZIP framing, media, opaque members, or a complete output
fingerprint.  It proves only that the semantic XML patch can be replayed in
memory.

The new `DocumentPublication` capability has the stronger authority.  It
retains the original `SourceArtifact`, records the original artifact length and
fingerprint, hashes the accepted forward artifact, and delegates inverse output
to authenticated complete-artifact restoration
([`source_backed.rs`](../../../../crates/litchi-docx/src/source_backed.rs#L1361),
[`artifact_restore.rs`](../../../../crates/litchi-opc/src/source_backed/artifact_restore.rs#L18)).
The reopened inverse test correctly demonstrates the stronger property by
rejecting a package whose main XML is equal but whose opaque member differs,
then restoring the exact original artifact from the authenticated changed
artifact.

The frozen 0495 harness calls only
`publish_document_commit_to_stream()`.  Its preflight
`inverse_restores_source_verified` field is set from
`commit.patch().inverse().apply(commit.snapshot())` and compares XML bytes
([`docx_managed_edit.rs`](../../../../tools/perf-baseline/src/docx_managed_edit.rs#L1024)).
Its `patch_oracles.scope` already says
`untimed_preflight_commit_patch_oracles`, which is the correct scope, but the
field must not be read as “complete artifact inverse” in a report.  The harness
does not time or receipt `publish_document_commit_with_inverse_to_stream()` or
`publish_document_inverse_to_stream()`.  The artifact-level claim must cite the
focused managed tests or a separate capability receipt.

The harness's `foreign_source_refusal_verified` is also narrower than its name
may suggest.  It opens the expected changed output and applies the forward XML
patch to that changed snapshot, so it verifies rejection of a target with the
wrong document bytes.  It still does not test a separately opened package with
the same main XML and different physical members.  The ordinary XML-patch
boundary is covered by the scoped source-backed regression below; the
complete-artifact inverse's foreign-member case remains a separate capability
test, not a frozen preflight oracle.

### 4. Unmanaged primary snapshots now retain scoped source identity

The OPC source-backed package creates a process-local `SourceLineage` for each
open and documents that two adapters with equal version tokens are still
different package instances whose patches must not cross
([`source_backed.rs`](../../../../crates/litchi-opc/src/source_backed.rs#L2702)).
The source-backed DOCX unmanaged branch now calls
`Snapshot::from_shared_xml_with_source_identity()` and stores an
`XmlStorage::OwnedWithIdentity` value.  This retains lineage, source version,
and Part URI without claiming managed cache ownership
([`source_backed.rs`](../../../../crates/litchi-docx/src/source_backed.rs#L1807),
[`transaction.rs`](../../../../crates/litchi-docx/src/document/transaction.rs#L793),
[`transaction.rs`](../../../../crates/litchi-docx/src/document/transaction.rs#L897)).
`Snapshot::same_source()` now compares that identity as well as bytes
([`transaction.rs`](../../../../crates/litchi-docx/src/document/transaction.rs#L1173)).
Every rewritten candidate keeps the identity through `with_rewritten_xml()`;
durable replay maps it to `Lineage::Identified`
([`transaction.rs`](../../../../crates/litchi-docx/src/document/transaction.rs#L926),
[`durable.rs`](../../../../crates/litchi-docx/src/document/transaction/durable.rs#L559)).

The scoped regression
`unmanaged_foreign_package_with_equal_main_xml_refuses_stale_commit_before_output`
opens two unmanaged packages with equal main XML and different media and
opaque members, stages the commit from the first, and verifies that publishing
it through the second returns `TransactionError::StaleSource` with an empty
output (`source_backed_managed_document_edit.rs:927–949`).  The supplied
candidate17 result reports this regression passing.  This closes the previous
ordinary unmanaged main-document cross-lineage finding.

The scope is deliberately narrow: this is evidence for the source-backed DOCX
main-document XML patch/commit path and its identity-preserving rewrites.  It
does not establish that every DOCX facade, durable operation, or separate
complete-artifact inverse scenario has equivalent coverage.  The complete
artifact inverse remains independently authenticated by full artifact length
and fingerprint, rather than by this XML patch oracle.

### 5. The frozen cache diagnostic label is narrower than its sampling point

`publish_document_commit_to_stream()` consumes `Package` by value.  The
current harness also drops the returned `published` snapshot before taking
the live gauge, then drops `commit` before taking `after_drop`
([`source_backed.rs`](../../../../crates/litchi-docx/src/source_backed.rs#L1293),
[`docx_managed_edit.rs`](../../../../tools/perf-baseline/src/docx_managed_edit.rs#L1186),
[`docx_managed_edit.rs`](../../../../tools/perf-baseline/src/docx_managed_edit.rs#L1187),
[`docx_managed_edit.rs`](../../../../tools/perf-baseline/src/docx_managed_edit.rs#L1191)).
Therefore the resource `after_drop` gauge now covers package consumption,
returned-snapshot release, and commit release.  There is no separate `Document`
owner in this measured path; provider, sink, budget, and execution-context
teardown remain outside this gauge.

The narrower issue is `cache_live`: it is captured after edit and commit
staging but before publication
([`docx_managed_edit.rs`](../../../../tools/perf-baseline/src/docx_managed_edit.rs#L1181)).
It is a post-edit/pre-publication cache diagnostic, not a post-publication live
gauge, and the report has no cache after-drop snapshot.  The resource gauge
scope remains valid; report interpretation should label `cache_live` with its
actual phase rather than treating it as a post-publication cache observation.
Cumulative input/output/work counters remain correctly unsuitable as release
gauges.

### 6. Direct settings policy inspection is fail-closed for MCE and namespace aliases

The changed ordinary-document publication path now reads the settings part
through `source_xml()` before inspecting it and applies the package source
fence around that operation
([`document_policy.rs`](../../../../crates/litchi-docx/src/source_backed/document_policy.rs#L40),
[`document_policy.rs`](../../../../crates/litchi-docx/src/source_backed/document_policy.rs#L98)).
Its guard rejects DTDs, processing instructions, non-predefined entity
references, malformed root structure, undeclared prefixes, actual MCE elements,
and every MCE attribute except `mc:Ignorable`
([`document_policy.rs`](../../../../crates/litchi-docx/src/source_backed/document_policy.rs#L257),
[`document_policy.rs`](../../../../crates/litchi-docx/src/source_backed/document_policy.rs#L536)).
An `Ignorable` token must resolve in the active scope and may not resolve to
the Word or MCE namespace.  This makes the direct source scan equivalent for
the protection and tracked-revision decision because a foreign wrapper cannot
promote a nested Word child without an admitted MCE projection.  The direct
Word policy remains owned by `variables::inspect_source_policy`; the separate
scan handles direct `trackRevisions`
([`variables/codec.rs`](../../../../crates/litchi-docx/src/variables/codec.rs#L247),
[`document_policy.rs`](../../../../crates/litchi-docx/src/source_backed/document_policy.rs#L459)).

The namespace declaration check is security-relevant.  `quick-xml` retains
namespace attribute values lexically while its resolver installs those raw
bytes as namespace bindings.  Thus an input such as
`xmlns:x=".../wordprocessingml/200&#54;/main"` can otherwise make
`x:documentProtection` look foreign to the direct scanner even though an Office
consumer resolves the entity-escaped URI to the Word namespace; the analogous
escaped MCE URI can hide `mc:ProcessContent`.  The policy refuses any `&` in a
namespace declaration before expanded-name checks, preserving the fail-closed
boundary without claiming that the resolver performs normalization
([`document_policy.rs`](../../../../crates/litchi-docx/src/source_backed/document_policy.rs#L552)).
The OPC source XML overlay validator applies the earlier refusal first:
`source_xml()` validates the exact decoded part and returns the typed
`SourceBackedOverlayUnavailable { reason: "source XML requires literal namespace bindings" }`
before the settings policy scanner is entered
([`source_backed.rs`](../../../../crates/litchi-opc/src/source_backed.rs#L5826),
[`xml_splice.rs`](../../../../crates/litchi-opc/src/xml_splice.rs#L1182)).
The direct settings check remains defense in depth for any caller that reaches
it with a source representation not already screened by that OPC seam; it is
not the observed refusal category for the escaped-namespace publication cases.
The focused managed cases cover unresolved `Ignorable`, `ProcessContent`,
`AlternateContent`, an entity-escaped MCE URI, and an entity-escaped Word
protection alias, with the last two expecting the precise OPC refusal before
output; the equivalent unmanaged cases expect the same refusal
([`source_backed_managed_document_edit.rs`](../../../../crates/litchi-docx/tests/source_backed_managed_document_edit.rs#L1401),
[`source_backed_managed_document_edit.rs`](../../../../crates/litchi-docx/tests/source_backed_managed_document_edit.rs#L1521)).
The standard `mc:Ignorable` plus `w14` extension case is admitted when it does
not enable a direct policy marker, while the same shape with direct tracking
enabled remains refused
([`source_backed_managed_document_edit.rs`](../../../../crates/litchi-docx/tests/source_backed_managed_document_edit.rs#L1328),
[`source_backed_managed_document_edit.rs`](../../../../crates/litchi-docx/tests/source_backed_managed_document_edit.rs#L1359)).

This finding is scoped to the ordinary changed-document publication policy.
The tail-append topology validator still uses the bounded MCE codec and then
parses its projected settings output
([`tail_append.rs`](../../../../crates/litchi-docx/src/source_backed/tail_append.rs#L1355),
[`tail_append.rs`](../../../../crates/litchi-docx/src/source_backed/tail_append.rs#L1411),
[`tail_append.rs`](../../../../crates/litchi-docx/src/source_backed/tail_append.rs#L1466)).
That codec decodes namespace attribute values before MCE expansion, so the
direct-scan alias finding must not be generalized to that existing projected
path.  Conversely, any future replacement of that projection with direct
inspection must carry the same escaped-namespace refusal and focused cases.
The final4 gate-10 receipt reports all 119 DOCX integration tests passing,
including the 36 managed-document-edit cases
([`final4-gate-10.stdout`](validation/final4-gate-10.stdout)).  No correctness
blocker remains for the frozen ordinary publication scope; the remaining
boundary is evidence scope rather than a permission to claim one settings
policy for every DOCX settings consumer.

## Acceptance reading

For the current implementation and frozen harness, the evidence should be
reported with these boundaries:

* Managed forward publication is source-authorized through `SourceXmlPart` and
  the topology preservation path; typed refusals and partial output remain
  distinct outcomes.
* `Patch::inverse()` proves document-XML reversal only.  It is not a complete
  DOCX artifact inverse and is outside the timed operation.
* `DocumentPublication` plus authenticated artifact restoration is the complete
  inverse API.  Its foreign-member, mutation, cancellation, budget, and
  partial-sink behavior is a focused capability result, not a 0495 latency row.
* The ordinary unmanaged main-document path has a scoped package-lineage
  regression; do not generalize that result to every facade or inverse API.
* The `cache_live` phase and the resource release sampling points must still be
  called out before making persistent-release claims.

No Cargo command or workload capture was run for this review.
