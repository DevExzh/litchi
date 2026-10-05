# 0915 DOCX section-layout design review

## Scope and disposition

The reviewed path is the source-backed existing-section layout query and
publication path:

- `crates/litchi-docx/src/source_backed.rs:2093` captures the current main XML,
  checks the source version, computes the complete `SourceArtifact` fingerprint,
  and constructs `layout::Snapshot`.
- `crates/litchi-docx/src/section/layout.rs:66` validates the source XML and
  parses the section `Inventory`.
- `source_backed.rs:1283` applies the patch and writes the overlay; the changed
  target is parsed again by `Patch::apply` at `layout.rs:547`.

The admissible small candidate is limited to
`ensure_source_section_inventory_xml` (`source_backed.rs:2698`). It changes
`reader.resolver().clone()` to a borrow and removes the two now-unnecessary
`&resolver` arguments at the `validate_source_section_element` calls. This is
the frozen three-line candidate: it removes the guard's per-event
`NamespaceResolver` `String`/`Vec` copies and their allocations while keeping
the scan and all helper behavior unchanged. The supplied profile identifies
the guard's per-event resolver clones as allocation work. This review makes no
latency or allocation-count claim; the frozen eight-pair gate is required for
acceptance (`ratio >= 1.05` and lower confidence bound `> 1`).

The analogous resolver clones in `section/inventory.rs:691` and
`section/layout.rs:857` are separate, untested follow-up candidates. They are
not part of this result or its performance gate.

The borrow is legal. quick-xml 0.42 exposes `NsReader::resolver() ->
&NamespaceResolver` and `NamespaceResolver::resolve_event(&self, event)`, and
the same direct-borrow pattern is already used by
`crates/litchi-pptx/src/presentation/source.rs:3728`. `resolve_event` is
read-only. `NsReader<&[u8]>` events borrow the input XML, not the reader's
resolver storage; no mutable reader operation occurs while the event is
matched. Keeping the existing `decoder` read, `read_event` call and error
mapping before resolution preserves error ordering. The candidate must not
widen to the inventory or fragment scans above, or to
`ensure_source_document_xml` (`source_backed.rs:2461`): those are separate
paths, and the managed document path has explicit per-event budget accounting
for its clone work.

## Identity, freshness, and retention boundaries

The candidate does not change source identity, freshness, or publication
authorization. `SourceArtifact::fingerprint()` (`crates/litchi-opc/src/source_backed.rs:3763`)
still fences the generic `ReadAt`, reads the complete artifact, charges its
work, and returns the exact digest. `source_version()` still fences the
adapter. A generic `Arc<dyn ReadAt>` is not an immutable owned snapshot: the
trait erases whether the implementation can mutate, and its version token is
the adapter's freshness contract. `OwnedSource` is immutable, but the current
source-backed API does not preserve an owned-source capability or allocation
identity after erasure. Consequently this review admits no digest skip or
memoization and no source-version-only shortcut.

`Patch` deliberately retains only exact `before`/`after` byte arcs and
authorization (`layout.rs:413`). A parsed inventory is not carried into a
patch. `Commit` already owns a cheap-cloneable, fully validated target
`Snapshot` (`layout.rs:379`), so a separate commit-only publication helper
could be investigated later: it would first run the current snapshot and all
existing patch authorization checks, then reuse that already-retained target.
That is distinct from adding a parsed inventory to `Patch` and does not justify
weakening source or fingerprint checks. It is outside this three-line resolver
candidate and requires a matched profile before implementation.

## Required verification

The resolver change must pass the existing DOCX section tests, especially:

- `source_snapshot_reads_only_mandatory_main_and_retains_exact_version`;
- `source_snapshot_applies_limits_and_rejects_a_changed_read_at`;
- `source_snapshot_refuses_mce_dtd_and_entity_syntax`;
- `source_backed_existing_section_layout_is_exact_and_cell_safe`;
- `source_backed_section_layout_noop_mce_strict_and_limits_are_typed`;
- `section_layout_signed_noop_changed_and_partial_sink_are_typed`;
- `section_layout_detached_patch_cannot_publish_to_source_package`.

The focused regression matrix must continue to report the same typed
`SectionInventoryLimit` resources (`input bytes`, `XML events`, `XML depth`,
and section bytes), `UnsafeEdit` refusals for MCE/DTD/entity syntax, XML and
invalid-format errors, exact no-op bytes, and exact stale/foreign/fingerprint
authorization errors. No new freshness test is required for this three-line
borrow: it preserves the existing `read_event`, source-version, artifact
fingerprint, and publication control flow. The existing freshness and
publication controls remain required. A future commit-only target-snapshot
reuse would need a separate mutation/freshness matrix before implementation;
that is a coverage gap for that future candidate, not a requirement here.

Budget behavior must remain unchanged: input/event/depth/section limits still
refuse at the same boundary, managed section-layout calls still return the
existing owned-snapshot `UnsafeEdit`, and no `Resource::Memory` or
`Resource::Work` charge is removed from the managed source paths. The frozen
eight-pair latency gate is the only performance decision; no digest, cache, or
retained-inventory result is inferred from this source review.

## ADR fit and blockers

This candidate fits ADR 0003's immutable source-bound snapshots and exact
patches, ADR 0005's measured work-elimination rule, and ADR 0006's typed
validation boundary. It changes neither public API ownership nor OPC physical
package ownership (ADRs 0010, 0011, and 0024). It also satisfies ADR 0032 by
retaining no derived value at all: the resolver is borrowed only for one event.

There is no ADR blocker for the resolver-borrow candidate. Digest reuse for a
generic mutable `ReadAt` would be blocked by the source identity/freshness
contract. Retaining a parsed inventory in `Patch` would be a variable-size
retained state and would require the ADR 0005 finite-ceiling, holding/release,
and accounting design; the current patch does not provide that contract. The
out-of-scope inventory and fragment resolver borrows have no result here and
must be measured independently.
