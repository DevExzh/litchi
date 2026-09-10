# Change 0495: current managed DOCX ownership review

Status: bounded, read-only source review. This describes the current
implementation after the namespace resolver, unmanaged source-identity, and
namespace admission fixes. The source review is supplemented by the final
gates and measurements linked below. No production file or protected worktree
was changed by the reviewer.

## Review boundary and evidence

The review covers the 27 Rust paths enumerated by
[development-candidate23-transfer.json](development-candidate23-transfer.json).
The candidate23 manifest has SHA-256
b1d1d877de6f71632a5f0c285200af9ad99e4a4b7ba4bd69da4506562123316a; all 27
listed paths currently match their recorded content hashes. It is a file-scope
record, not build authentication. [Final gates](final-gates.json) authenticate
the builds and passing correctness checks. [Formal verification](verification/formal1.json)
and [profile review](profile-review.md) supply the separate measurement evidence.
The changes since candidate21 were test-only corrections for the exact earlier
OPC namespace refusal, including unmanaged cases.

The assessed contracts are finite resource admission, cancellation checks,
immutable owner-retaining snapshots, exact source identity, lossless namespace
resolution, and typed refusal where publication authority or durable ownership
is absent. This is source inspection only.

## Snapshot ownership and source identity

Snapshot has four storage cases in
[transaction.rs](../../../../crates/litchi-docx/src/document/transaction.rs#L767-L782):

* Owned(Arc<Vec<u8>>) is the standalone unmanaged compatibility case.
* OwnedWithIdentity is still unmanaged, but retains an Arc<SourceIdentity> for
  exact same-source checks when a package cannot retain a managed owner.
* Managed(Arc<ManagedXml>) retains PartData, identity, and a ManagedAdmission.
* Source retains SourceXmlPart, identity, and a ManagedAdmission for
  source-authorized publication.

is_managed() correctly matches only Managed and Source. same_source() requires
equal XML plus matching lineage, version, and equivalent Part URI when both
snapshots carry identity; mixed identity remains unequal. The ordinary
source-backed unmanaged constructor now uses
from_shared_xml_with_source_identity() and therefore rejects equal main XML
from a different package lineage.

Ordinary unmanaged edit candidates and the changed commit fallback use
with_rewritten_xml(), which retains OwnedWithIdentity. Patch::inverse() clones
both snapshots, and Patch::apply() uses the identity-aware same_source() check.
Durable semantic replay reaches the same identity-preserving commit path.
Durable restore receives the current source identity and rebuilds an
OwnedWithIdentity snapshot when one is available. Lineage::Identified carries
XML and identity through durable composition admission. Source inspection finds
no remaining identity drop in edit, commit, inverse, or durable restore.

Changed ordinary publication also runs the source-bound settings policy guard
before emitting output. The guard checks the source version fence, main/settings
dialect and relationship/content-type topology, admits bounded settings scan
workspace, and refuses enabled document protection, tracked revisions, DTD/PI
or non-predefined entities, unresolved prefixes, unsupported markup
compatibility controls, and mixed settings dialects. Exact no-op publication
deliberately bypasses this changed-document guard and preserves the complete
source package. Candidate21 adds focused cases for protection flags, standard
ignorable settings preservation in managed and unmanaged paths, and settings
MCE/DTD/dialect refusal. Those tests and the final gates remain unverified.

## Retained views and namespace resolution

XmlRef retains the full PartData or SourceXmlPart owner and managed admission
alongside every range. Paragraph, run, inline, extension, hyperlink, collapsed,
and symEx views therefore cannot outlive their source owner. The
NamespaceResolverLease owns both the copied resolver and its optional
ManagedNamespaceAdmission; codec calls borrow the resolver only while the lease
is live, and no resolver reference escapes into returned values.

For a non-root range, namespace_resolver() scans the full owner through the
selected start event. This includes declarations attached to that element and
allows NsReader to pop a preceding sibling before the selected event. It then
copies only NamespaceResolver::bindings(), whose effective-binding iterator
omits overridden declarations, into a fresh resolver. The source fixture covers
an inherited sx binding, paragraph and run shadowing, sibling restoration, and
an unrelated foreign prefix. Source inspection finds no binding-lifetime or
shadowing semantic defect in this path.

The final namespace admission reserves
32 * owner_span_len plus fixed state and reader/admission metadata before the
scan. The owner span reaches through the retained fragment, so the envelope
covers the source reader resolver, copied lease resolver, codec resolver clone,
and geometric Vec growth for quick-xml 0.41 namespace metadata. The checked
Work reservation adds
(fragment_len + 1) * (owner_span_len + 1) to scan work before parsing, while
the scan itself charges event_bytes * active_bindings. This bounds the
fragment-by-inherited-prefix cross-term and keeps cancellation/budget failure
before codec state grows. The 32x memory factor is conservative accounting, not
a measured allocation or RSS result.

## Admission and accounting review

The earlier ownership findings remain addressed by inspection:

* managed text no-ops release transient depth before comparing text and return
  before retaining operation or string admissions;
* reverting the last staged replacement clears operation, string, and byte
  accounting and restores the exact base source;
* OperationList::Empty avoids an empty Arc allocation, while managed operation
  admission covers Vec-to-Arc overlap and inverse construction reserves before
  growing its vector;
* ScanAdmission::release_depth() drops transient nesting capacity after owner
  state is retained;
* managed Document retains DocumentIndexAdmission, reserves index state before
  construction, and rejects queries that would expose unbudgeted owner-backed
  semantic views;
* the root-formatted OPC path admits authored Vec<u8> wrapper/parser state and
  output/edit metadata through the execution context before publication.

These are conservative admission envelopes, not measured allocation or RSS
claims. The remaining managed history/redo, durable owner ledger, composition,
and other mutator refusals are typed scope limits. The ordinary managed surface
is the source-authorized direct paragraph-text edit and its source-checked
publication path; signed changed edits and dependency-bearing operations still
fail closed.

No additional concrete identity or namespace blocker was found in this
source review. The final focused tests, Clippy/rustdoc gates, and profile
captures completed; their scope and limitations remain in the linked records.
Full managed durable, history, composition,
and other-mutator support remains unimplemented.
