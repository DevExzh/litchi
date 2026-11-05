# 0362ul81b — PPTX metadata structural review return

- Date: 2026-10-25.
- Status: source-only return for the 81b review; typed-limit integration and
  Cargo validation remain with the coordinator.
- Base: `921a9df55b225e047c0384596c358d2f8162b937`.

This return removes the remaining public `color_map::Draft::snapshot` wrapper
and migrates its tests to `Draft::value`; the minimal and full PPTX glossary
profiles now ban the wrapper by scope.  Borrowed guide and color-map sources
are admitted under their selected input policy before their retained source
copy is made.

Guide and source-backed custom-show edits now calculate the exact aggregate
replacement delta before allocating owned replacement buffers.  The preflight
covers all changed owner spans, inserted entries and newly rendered slide
relationships, including escaped XML lengths and empty-owner expansion.

The structure parser admits raw element names, attributes, text, CDATA,
comments, general references and declarations while quick-xml still borrows
the event; ownership therefore cannot bypass the selected string and
attribute limits.  Structure graph edits use the source-preserving custom-show
wire layout instead of rebuilding `custShowLst`, retaining unknown nested
attributes/elements, comments and whitespace.  The structure fixture and
regression assertions cover an opaque nested custom-show list attribute and
element.

The common 61 physical scan/no-op base check and record-84 typed
`Limit { kind, observed, maximum }`/`InvalidLimit` migration remain separate
coordinated work.  No Cargo, rustdoc, clippy, native build or ABBA run was
performed in this source-only lane.
