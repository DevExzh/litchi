# ADR 0013: PowerPoint notes ownership and atomic deletion

- Status: Accepted
- Date: 2026-08-03

## Context

The PowerPoint writer could create and read speaker notes, but neither an
opened package nor a mutable slide exposed a deletion operation. The existing
notes graph loader copied every notes, master, and theme payload, while its
store operation deliberately rejected ownership changes. Deleting only a
slide relationship or only a notes part could leave a dangling graph, collect
a shared notes master or theme, mutate an inactive or malformed graph, or
invalidate a signature before the requested edit was known to succeed.

The ordinary facade must select slides by producer-visible meaning rather than
relationship IDs or part names. Deletion must remain idempotent and must not
turn a malformed package into a partially edited one.

## Decision

`litchi-pptx::notes` is the sole owner of the PresentationML notes model,
bounded XML codec, package graph service, plain-text producer, and
notes-master assets. There are no forwarding aliases or long type names. Its
short contextual vocabulary is `Conformance`,
`Theme`, `Master`, `Slide`, and `Graph`; physical topology fields are private
and available only through diagnostic accessors.

`load` returns a lifetime-free editable graph, copying each validated notes,
master, and theme payload exactly once. The focused `slide` read copies only
the selected notes payload, and deletion continues to use the metadata-only
index without payload copies. XML replacement returns the previous allocation.
The owner accepts the presentation, slideshow, and template main-part content
types in both macro-free and macro-enabled families, not only `.pptx`.

The consuming `put` operation validates and stages the complete replacement,
moves caller-owned buffers into canonical OPC parts, and invalidates signatures
only after commit. It cannot retarget or orphan the existing coherent resource
set; an exact graph is a signature-preserving no-op. Plain-text output defaults
to Transitional conformance, while the explicit conformance-aware producer is
used for Strict graph edits. Fresh-package notes-master generation remains a
Transitional producer path; editing a loaded Strict graph preserves its
validated master and theme resources.

`pptx::Package` owns the concise opened-package operations:

- `notes() -> Result<Option<Graph>>` and consuming
  `put_notes(graph) -> Result<()>` expose the canonical owner without aliases;
- `remove_notes(slide) -> Result<bool>` accepts the existing exact-name-first
  `SlideKey` conversion and a checked zero-based position;
- `clear_notes() -> Result<usize>` removes notes from every slide; and
- `MutableSlide::clear_notes() -> bool` handles not-yet-packaged authoring.

Missing notes are not exceptional. Missing or ambiguous names, out-of-range
positions, a dirty legacy writer, malformed XML, invalid relationships, and
unexpected incoming edges are typed failures before mutation. Package graph
reads and mutations reject dirty legacy-writer state because it could otherwise
return stale data or overwrite the accepted edit during later materialization.

The package operation first builds a metadata-only index of the complete
Strict or Transitional notes graph. The index records actual stored `PackURI`
keys after OPC resolution, but those keys remain below the ordinary facade. It
validates each notes slide's backlink and notes-master edge, rejects orphan or
multiply-owned notes parts, and scans every package relationship for unexpected
incoming ownership.

Mutation uses a staged plan. Each selected slide owner is cloned before the
first package change; built-in parts share their immutable payload allocation,
while the staged relationship collection removes only the notes edge. After
all allocation, lookup, target, and relationship checks succeed, commit
replaces the staged slide owners, removes the exact stored notes-part keys, and
invalidates signatures. These commit operations are infallible under exclusive
package ownership. The shared notes master and its theme are retained.

## Consequences

- Callers can remove one slide's notes by name or position without observing
  package identities, and repeated removal is safe.
- Selected notes parts and owning relationships disappear together; unrelated
  slide XML, ordering, shapes, notes, and shared notes infrastructure remain.
- Planning memory scales with relationship metadata and staged relationship
  collections for built-in parts, not with copied notes or slide payloads.
- Whole-graph editing is move-first on storage, but the lifetime-free loaded
  graph deliberately pays one bounded payload copy. A separate borrowed graph
  view is not introduced until it can remain ergonomic and cannot outlive its
  package snapshot.

## Verification

Focused tests must cover Strict and Transitional graphs, conformance-aware text
replacement, name and position selection on a saved and reopened two-slide
deck, single and all-note deletion, idempotence, byte-identical slide XML,
retained master/theme resources, case-folded stored keys, malformed graphs,
unexpected incoming edges, ambiguous or missing names, out-of-range positions,
signature-preserving no-ops, and dirty-writer rejection through both `Package`
and `Presentation` reads and every package mutation entry.
