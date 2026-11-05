# 0362re — DOCX section-layout raw tree

Status: source migration and focused validation in progress; performance is
not qualified by this record.

The standalone section-layout surface now exposes a bounded raw codec model in
`litchi_docx::section::layout::raw`:

- `raw::Tree::from_xml` and `from_xml_at` require `&section::Limits`.
- `Snapshot::raw_tree` and `raw_tree_at` intersect the requested policy with
  the policy retained when the snapshot was admitted.
- `raw::Tree` provides read-only getters, mutable section-layout setters,
  `to_xml`, and an exact reversible `raw::Patch` with `inverse` and `apply`.
- Raw output retains the complete source XML, including unknown markup, and
  validates the rewritten document before returning bytes or a patch.
- The detached record `Edit`, `Commit`, and `Patch` remain crate-private.
  Source-backed callers use `Package::edit_section_layout` and
  `Package::apply_section_layout`. A raw patch can still be converted into
  the source-backed family patch and applied to a package, preserving the
  former offline-authoring route; that bridge remains XML-content scoped.
  Package-family patches created by an owner retain complete artifact
  identity; their inverse restores the exact package artifact. Raw patches
  re-admit output under the recipient snapshot's limits.
- Section layout never had a serialized `to_bytes`/`from_bytes` patch API.
  The former stream `Publication` route was replaced by the canonical family
  `Commit<Package, Patch, Diagnostics>` and deferred `Package::write_to`/
  `save` route in record 0355a. The raw patch remains durable in memory and
  reversible across reopened matching XML snapshots.

All source-backed section-layout tests were migrated to the package family.
The detached raw test covers policy intersection, forward/inverse replay on a
reopened equal-content snapshot, a second edit, and fitting versus over-limit
recipient admission. The owner bridge test applies that raw patch through the
family and applies its inverse. The package test covers equal-artifact
reopening and rejects an unrelated-member change. Before the final
error-taxonomy cleanup, the focused `litchi-docx` library and test checks
passed with the capped environment and an external target directory. The
taxonomy cleanup and the full affected suite still require a fresh check; a
workspace gate and first/repeated output measurements remain pending. No
benchmark binaries or large evidence are stored in this record.

The changed owner edit now computes the complete post-splice XML length before
`replace_range` reserves the next document buffer. A result above the retained
`max_input_bytes` policy returns the typed `SectionInventoryLimit` with
`LimitKind::InputBytes`; it does not allocate the full rejected document first.
The regression changes a section by adding margins under an exact-size input
ceiling and verifies the source allocation remains unchanged. This repair is
source-formatted and statically checked; fresh Cargo validation remains
pending the coordinator's build hold.

All section-layout base and provenance mismatches use the crate's canonical
`Error::PatchConflict`; a partial provenance tuple is an internal invariant
failure. The supported raw-to-owner bridge remains usable: its detached patch
checks the main-document XML scope and recipient limits, while an owner-created
family patch additionally proves the complete package artifact.

The external performance harness was migrated to the source-backed family and
raw patch oracle. Its former detached-commit retained-base byte fields are
reported as unavailable rather than inferred from a private record
implementation; allocator lifecycle and exact publication/inverse checks
remain active.
