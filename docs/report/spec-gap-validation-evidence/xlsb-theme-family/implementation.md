# Complete Theme-family host integration

The shared fragment model from the previous batch is now reachable through a
complete DrawingML Theme and the XLSB workbook-owned Theme surface.

## Caller surface and ownership

- `theme::family::part::read_family` projects optional metadata from borrowed
  Theme bytes. The part module also provides source-range/profile inspection,
  an immutable shared snapshot, and bounded add/replace/remove XML helpers.
- XLSB `theme::Snapshot` and source-backed `theme::View` expose `family()`.
  Their existing detached transaction adds `set_family` and `remove_family`.
  Family and base Theme edits can participate in the same commit.
- The optional whole-Theme owner transaction exposes family staging for
  creation and existing-owner edits. Removing the complete Theme clears its
  family metadata as part of the same owner operation.
- `WorkbookWriter::set_theme_family` supports authoring with the writer's
  selected Theme. XLSB additions use the fixture-backed native discriminator;
  shared helpers also expose the normative namespace-URI profile.

Existing-family replacement copies the requested scalar values onto the
current source-backed family. Incoming opaque content does not replace the
current owner's unknown attributes or children. Removing family metadata
deletes the selected `themeFamily` element, its `ext` if otherwise empty, and
the root `extLst` if that too becomes empty. XML whitespace alone does not
retain a container; comments, unrelated content, and foreign attributes do.
The inverse patch restores the exact pre-removal source. This closure policy
supersedes the initial retention behavior; its fresh evidence is in
`../theme-family-removal-closure/`.

Family XML ownership remains in DrawingML. XLSB resolves the workbook Theme
relationship and performs source-checked, signature-aware package publication.
Publication uses the existing OPC owned-XML splice seam so native declarations
and line endings remain source-backed through save/reopen. It does not turn
normal save into XML normalization or repair.

## Namespace and source model

Only the direct namespace-resolved `theme/extLst/ext/themeFamily` path under an
admitted extension identifier is typed. Identifiers follow their XML Schema
`token` normalization, while their original spelling survives in the source.
Multiple admitted owners and malformed recognized content are refused.
Selective family add, replacement, and removal refuse MCE-wrapped owned
extension lists or admitted extensions, even when empty, and MCE children
inside admitted extensions. Read-only projection keeps these branches opaque.
See `../theme-family-mce-refusal/` for the expanded boundary and fresh checks. Direct owned extension containers
reject non-whitespace text, CDATA, and references. A validated standalone UTF-8
BOM is stripped before fragment embedding; XML declarations remain refused
when inserting a child.

Extracted `Family` values retain one namespace-complete fragment source.
Inherited namespace declarations are supplied during extraction, including
bindings that opaque QName-valued attributes may need. Consequently this
standalone fragment can include closure declarations absent from the original
subtree. Exact original bytes remain available from the complete Theme source
and the shared family range. No second raw fragment copy is retained merely
for that convenience.

The shared part snapshot shares its complete source and immutable owner
metadata. XLSB's borrowed projection path does not copy the complete Theme
into that snapshot; source-backed views retain their existing managed
`PartData`. Semantic no-op transactions retain the original complete XML.

## Bounds and limits of support

Complete Theme input/output has the existing 8 MiB hard ceiling; family
fragments, including namespace closure declarations, have a 1 MiB ceiling.
Caller-specific output limits are checked before insertion-container and complete
output allocation. Family serialization remains separately bounded by its
1 MiB ceiling. Active namespace scope is capped at 257 bindings, including the
implicit XML binding; inherited namespace removal uses one bounded pass.
XML node, depth, attribute, namespace, and scalar bounds remain finite. The
reader enforces its XML 1.0/UTF-8 profile and rejects DTDs, processing
instructions, invalid declarations, forbidden controls, and unknown entity
references.

This is scalar family metadata CRUD. It does not apply a theme's visual
formatting, interpret all effect/style matrices, edit MCE branch ownership,
provide durable patches or merging, or add DOCX/PPTX package integration.
Native inputs and offline schemas support the stated XML/package behavior;
changed outputs have not been opened in a native Office application.

Final gate receipts, independent review, the standalone caller, and the
performance directory record validation of the committed implementation.
