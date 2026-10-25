# 0362ul87c — shared OOXML VBA publication boundary

Status: source-only implementation on `uo2/ooxml-vba-publication-0362ul87c`;
87cc and 87cd review repairs are included in the follow-up commit.
This slice follows the boundary audit in 0362ul87. It has not run Cargo,
rustdoc, native qualification, or performance measurements; those require a
coordinator validation slot after the four OOXML consumers are composed.

The 87cc repair keeps the selected graph's admitted relationship and
content-types tokens under the effective edit/recipient `SourceLimits` before
publication. Store, remove, and divergent replay compose source-preserving
part overrides, retain the host `.rels` declaration, and remove only obsolete
graph declarations. Aggregate validation now accounts for present relationship
members, XML events, content-type mappings, ZIP members/entries, member names,
and the projected part/relationship budgets. Structural replay tests serialize
and reopen independent recipients before forward and inverse application.

The shared VBA graph now exposes `vba::Edit::new(&OpcPackage, Host,
EditOptions)`, `Edit::{store,remove,reset}`, `Edit::commit()`, and the shared
`Patch` replay family. Construction borrows the source and resolves only its
main-part name; graph capture and the one structural candidate clone are
deferred until the first real mutation. Unchanged payloads remain shared, and
the edit publishes an `litchi_core::edit::Commit`
containing the new `OpcPackage`, exact before/after graph patch, and
content-free diagnostics (or a compact selected-graph proof for an untouched
edit). `Edit::store` accepts owned byte vectors only after
the selected project, supplemental, graph, part-count, and retained
`SourceLimits` policy has admitted them. `Patch::apply(&OpcPackage)` checks
the recipient policy before cloning or publishing, authenticates the recorded
selected graph base, publishes a candidate atomically, and leaves the
recipient unchanged on refusal; a mismatch is `Error::PatchConflict`.
`Patch::inverse()` swaps the exact graph states. The old public free mutators
are gone. Signed-package refusal, atomic graph validation, content-type
transitions, relationship exclusivity, Word's supplemental data, and
byte-preserving payload sharing remain in private candidate helpers behind the
edit family.

Snapshots retain the concrete OPC parts rather than rebuilding `BlobPart`
values from a reduced relationship projection. Replay therefore carries
relationship source captures, admitted identifiers, content-type provenance,
opaque part state, and noncanonical relationship spelling/order for the
affected source/project/supplemental closure while unrelated package parts are
left alone. An identical store is detected before mutation, so it preserves
source bytes, signatures, and unchanged diagnostics even when the existing
graph uses a noncanonical relationship identifier. The source-bound batch
publication restores those captures for the source, project, and supplemental
`.rels` members; it does not silently fall back to `try_add_part` and lose
provenance.

An untouched edit now commits a compact SHA-256 proof of only the selected
graph instead of capturing and retaining both graph snapshots. Proof capture is
`O(selected source/project/supplemental bytes + admitted manifest bytes)` and
shares the existing payload allocations; it is not an `O(package parts)`
snapshot. Replay checks that proof
against the recipient's selected source, project, and supplemental payloads and
relationship members, so it remains content-based while leaving unrelated
package changes outside the patch scope. Changed patches preflight the
recipient's projected aggregate part bytes, relationship-member bytes and XML
events, edge counts, relationship-part count, content-type mappings, archive
member/entry dimensions, per-member ceilings, and content-types manifest size
before cloning or publishing. Aggregate validation necessarily scans package
part metadata (`O(package parts)`); it does not claim an `O(1)` package-wide
budget check. A tighter recipient limit returns the typed VBA limit and leaves
the recipient byte-identical.

The DOCX, PPTX, and XLSX owner edit families continue to publish through their
existing `Edit::commit` triads. Their private staging adapters compose the
shared VBA edit and assign its committed candidate only after the operation
succeeds; the owner source remains unchanged on every refusal. The existing project
discovery and explicit `litchi_vba::Limits` payload parser are read-only and
were not widened into a raw mutation route. XLSB has its own host-local VBA
graph implementation and is unchanged by this shared OOXML slice.

Consumers migrated in this commit:

* `litchi-docx::vba_project::{store_vba_project,remove_vba_project}`;
* `litchi-pptx::presentation::embedded::vba::{store,remove}`;
* `litchi-xlsx::active_x::vba::{store,remove}`; and
* the shared graph tests, including signed refusal, dangling-target atomicity,
  late content-type rollback, equal-copy replay with unrelated edges,
  inverse replay, selected-graph patch conflict, reset/no-op diagnostics,
  pre-staging limit refusal, noncanonical relationship round-trip/inverse,
  compact no-op proof replay, and atomic rejection under tighter recipient
  aggregate part/relationship limits.

No public semantic helper in this slice takes `&mut OpcPackage`. The physical
OPC graph methods and explicit-limit VBA reader remain codec/substrate APIs;
no alias or deprecated forwarding entry point was added.

## 87cd review repair

The selected graph capture path now admits retained part lengths from
`PartMetadata` before any deferred payload is decoded. Graph inspection uses
relationships and content types only; source, project, and Word supplemental
payloads are decoded only after the effective operation/recipient policy has
accepted their declared lengths, with a post-decode length check for corrupt
or inconsistent source metadata. This keeps untouched proof replay and
candidate staging content-based without an up-front full-owner clone.

Recipient validation re-admits every retained OPC part, relationship member,
opaque ZIP member, central-directory name, metadata budget, compressed-size
ceiling, and per-entry ceiling. `OpcPackage::retained_archive_metrics` keeps
the retained opaque-member count and aggregate bytes in the physical source
boundary, so divergent replay charges the actual recipient-composed
`[Content_Types].xml` and preserves unknown ZIP payload capability. The
prospective archive member/entry and aggregate checks replace the old
projection-only estimate; the scan is intentionally `O(package parts)`.

The source content-types and relationships splicers now select elements by
expanded OPC namespace, matching the ingress parsers. A vendor namespace
element with the local name `Override` or `Relationship` is opaque: it is not
removed, counted, or used for a selector, and foreign attributes with the
same local name do not become OPC attributes. The review regression set covers
foreign namespace preservation, pre-hydration refusal, opaque aggregate
accounting, and recipient-composed manifest limits.
