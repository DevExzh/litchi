# 0362ul81a — PPTX metadata structural review return

This review closes the structural 81a findings for the eight metadata families
owned by the 81 slice: custom shows, guides, sections, handout masters,
protection, templates, color maps and presentation structure.

The compatibility wrappers `Snapshot::read`, projection aliases, draft
`snapshot` accessors and the custom-show `push` alias were deleted.  Existing
consumers use the canonical owner or draft accessors; the minimal and full
PPTX glossary profiles register the scoped bans.

All detached writers now admit dynamic output through a fallible bounded
writer.  Escaped values are preflighted for their exact encoded length before
the writer grows, replacement and insertion paths calculate the final output
length before reservation, and source attributes/text are checked before
decoded values are owned.  The custom-show source layout reuses the checked
parser attributes instead of reparsing an unbounded raw start tag.

Handout and protection remain canonical detached models: their documented
typed writers do not promise opaque unknown-markup round trips.  The
source-preserving guide, section, custom-show, structure and color-map paths
retain their opaque bytes and strict/transitional namespace behavior.  The
common 61 physical source scan and no-op base check are still separate
integration returns.

This review was checked with rustfmt, `git diff --check`, Python syntax/JSON
validation and focused source sweeps.  Cargo, rustdoc, clippy and ABBA were
not run in this source-only lane.  The final commit must be combined with the
record-84 typed `Limit { kind, observed, maximum }`/`InvalidLimit` migration.
