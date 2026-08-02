# ADR 0013: Atomic, package-aware PowerPoint notes deletion

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

`pptx::Package` owns the concise opened-package operations:

- `remove_notes(slide) -> Result<bool>` accepts the existing exact-name-first
  `SlideKey` conversion and a checked zero-based position;
- `clear_notes() -> Result<usize>` removes notes from every slide; and
- `MutableSlide::clear_notes() -> bool` handles not-yet-packaged authoring.

Missing notes are not exceptional. Missing or ambiguous names, out-of-range
positions, a dirty legacy writer, malformed XML, invalid relationships, and
unexpected incoming edges are typed failures before mutation.

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
  This structural property is not a latency, throughput, or peak-memory
  benchmark claim.
- Native PowerPoint open/resave evidence is still required before expanding
  this focused graph guarantee into a broad compatibility claim.

## Verification

Focused tests cover Strict and Transitional graphs, name and position
selection on a saved and reopened two-slide deck, single and all-note deletion,
idempotence, byte-identical slide XML, retained master/theme resources,
case-folded stored keys, malformed graphs, unexpected incoming edges,
ambiguous or missing names, out-of-range positions, and dirty-writer rejection.
The 7 focused graph tests, mutable-slide regression, and 4 reopened-package
integration tests pass, together with warning-denied Clippy, rustdoc,
formatting, diff validation, and workspace lint. Native PowerPoint verification
and the previously green full-workspace test suite are not repeated for this
slice.
