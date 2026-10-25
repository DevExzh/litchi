# 0362ul87c — shared OOXML VBA publication boundary

Status: source-only implementation on `uo2/ooxml-vba-publication-0362ul87c`;
87cc, 87cd, 87ce, 87cf, 87cg, and 87ch review repairs are included in the
follow-up commit.
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

## 87ce review repair

Retained ZIP metadata admission now counts each entry's variable central
directory fields exactly, without charging the 46-byte fixed central header,
and adds the terminal EOCD comment plus ZIP64 extensible data. This avoids
false metadata refusals while still admitting every variable archive field.
Opaque-member accounting follows the immutable `SourceMemberKind` captured at
ingress and the writer's preservation plan, rather than the package's current
part map; removed known source parts therefore do not become opaque while
unknown source members remain charged when the writer carries them forward.

`Patch::apply` and the no-op proof path now run the complete recipient
metadata/physical admission before selected graph capture. Central-directory
names, metadata, entries, per-member sizes, retained aggregate bytes, OPC part
metadata, relationship totals, and content-types mappings are checked without
decoding a selected payload. A regression recipient with an unrelated
oversized member proves the refusal is atomic and leaves the selected deferred
decode counter at zero. The divergent content-types replay test also derives
the composed archive aggregate and covers both the exact limit and one-byte
under boundary.

## 87cf review repair

The source archive admission pass now charges only the immutable source
members the targeted writer will copy for the recipient's current graph.
Removed known parts and their `.rels` entries are omitted before name,
metadata, compressed-size, entry-size, and aggregate checks; unknown source
members, including directory markers, remain retained according to their
original `SourceMemberKind`. Regenerated known members are admitted through
the current metadata and prospective graph checks instead of their historical
central-directory sizes. Missing preservation provenance is a typed refusal,
not an historical-source fallback.

Final archive member and entry counts are composed from the actual current
known parts, retained empty relationships members, one mandatory content-types
member, the relationship members (including the package-level `.rels` member),
and immutable opaque entries.
The former source-count-minus-current-part-count estimate is gone. Store and
remove staging use the same metadata-only recipient admission as replay before
selected graph capture. A serialized/reopened recipient regression removes a
known unrelated part with a relationship, adds a new unrelated part with its
own relationship, and checks exact and one-byte-under member, entry, and
aggregate limits atomically.

## 87cg review repair

The retained-source admission and targeted writer now share the same
preservation predicates. Equal decoded payloads use the source central record
and compressed representation when the writer does; a deferred source payload
is not hydrated merely to admit an unrelated recipient member. Selected
publication may force the exact source comparison needed for its output plan.
Relationship admission likewise follows the writer through source capture,
semantic source binding, and canonical-byte comparison, so a mutate-then-revert
`.rels` collection retains its noncanonical source bytes and central metadata.

Retained metrics include the current publication image's variable metadata,
including preserved records, regenerated known names, and the archive tail.
Prospective checks replace the selected graph's part and relationship records
with the target names and add newly appended generated entries, while keeping
opaque source records and their metadata. Exact and one-byte-under metadata
limits cover a serialized VBA replay. A compact recipient admission summary is
carried from staging/preflight into publication validation, avoiding a second
retained-central-directory or whole-recipient relationship scan.

## 87ch review repair

The physical relationship-member projection now uses one convention end to
end: `RelationshipTotals::members` already includes the package-level
`_rels/.rels` member, so recipient and structural archive counts add only the
single `[Content_Types].xml` member. Authored and retained projections use the
same formula, and the serialized replay boundary test asserts the resulting
member count against the relationship total rather than a historical source
count.

The admission and targeted writer preservation predicates now use one exact
bounded comparison. When a caller replaces a deferred source payload with a
same-length candidate, the retained ZIP member is streamed through the fixed
verified reader and compared byte-for-byte; a mismatch is regenerated, while
equal bytes still copy the source central record and compressed representation.
The comparison does not call deferred `force()` and therefore does not retain a
payload-sized allocation or charge the deferred decode counter. The regression
uses a same-length unrelated replacement, a tight compressed-size and metadata
policy, and verifies that the replacement is published without a second
deferred hydration. Existing equal-byte central-metadata and relationship
round-trip coverage remains in place.

Structural archive validation reuses the completed recipient admission summary
for package-wide names, sizes, payload, and opaque-entry checks. It validates
only the prospective target graph entries after composition, so store, remove,
and replay do not repeat the package-wide scan or hydrate unrelated deferred
payloads before publication.

## 87ci review repair

Retained archive metadata now accounts for the canonical `_rels/.rels` member
name when an accepted source archive omitted the package-level relationships
member. The ordinary writer emits that generated member on its fallback path,
so the recipient projection now charges its variable central-directory metadata
exactly once while still avoiding a duplicate charge for sources that carried
the member. A no-root source regression checks the exact and one-byte-under
metadata limits and verifies that publication emits the generated root member.
The equal-payload and mutate-then-revert relationship cases also cover the
same tight metadata boundary.

## 87cj compiler repair

The first isolated common-crate compiler pass exposed only integration errors
in the shared OPC implementation. The retained graph admission branch now
uses an ordinary `match` for its non-candidate case, and the writer passes
byte-slice views to the exact source-payload predicate. Verified payloads use
the same slice comparison as ready and transferred payloads; the deferred
comparison keeps its fixed reader scratch buffer and ends each `fill_buf`
borrow before `consume`.

The namespace-aware content-types and relationships splicers extract a boolean
for the exact OPC namespace before asking the reader for an event offset. This
removes the `NsReader` borrow overlap without allocating a namespace string or
weakening expanded-name matching. The missing shared `ContentType` import and
the now-unused buffered-reader import are corrected as part of the same
source-only repair.
