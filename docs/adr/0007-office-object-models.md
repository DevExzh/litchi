# ADR 0007: Office object models

- Status: Accepted
- Date: 2026-07-31

## Cross-format rules

DOC/DOCX, PPT/PPTX, and XLS/XLSB/XLSX have parallel semantic vocabularies but
distinct concrete payload types. Cross-format algorithms use static-dispatch
traits and content-event streams; autodetection returns a flat data-bearing
`File` enum. Conversion is explicit and fidelity-reported.

Content extraction defaults to visible accepted content and can request review
or all-content projections. Events carry provenance and visibility. Semantic
conversion is separate from high-fidelity rendering; rendering requires ordered
font sources and reports substitutions and metrics.

Templates compile into schema-aware typed values where possible and a dynamic
fallback otherwise. Plain strings are inert. Formula, field, link, rich text,
image, and other active interpretations require explicit wrappers.

Zero-argument creation returns a deterministic, compile-known valid baseline:
DOCX has a valid body and final section, XLSX has one visible worksheet, and
PPTX has a coherent theme/master/layout graph with no slide. Customized creation
uses a fallible builder. Ownership-consuming insertion errors return the rejected
value so callers never lose it.

## Word

Word exposes data-bearing `Block` and `Inline` enums with unknown fallbacks.
Main body, headers, footers, notes, comments, and text boxes share borrowed
`Story` traversal while retaining owner-specific editors. Sections are logical
ranges; header/footer roles resolve inheritance and require an explicit
edit-source or local-override choice.

Bookmarks, comments, permissions, revisions, and other overlapping ranges are
affinity-aware anchored annotations, not forced tree nesting. Foot/endnotes are
atomic marker-plus-story objects. Fields retain instruction, displayed cache,
freshness, and delimiters as one logical object. Hyperlinks unify relationship
and field encodings without becoming executable.

Lists use semantic start/continue/restart/indent operations rather than numeric
IDs. Content controls are typed and keep custom-XML bindings atomic. Tables
distinguish actual empty cells, covered spans, and omitted slots; lossless span
updates are automatic and ambiguous content disposition is explicit. Drawings
retain identity across typed inline/floating placement. Equations combine a
semantic `litchi-math` tree with lossless Office representation.

Generated TOCs, indexes, authorities, and bibliographies retain rules, sources,
displayed cache, and freshness. Cross-references target semantic handles;
captions are target-linked objects. Mail merge consumes caller-provided data and
never opens configured sources implicitly.

## Spreadsheets

Programmatic ranges are checked, zero-based, and half-open. Parsed A1/R1C1 forms
retain inclusive syntax and absolute/relative semantics. Sparse `cells(range)`
and dense budgeted `grid(range)` are distinct. Declared, stored, content, and
formatted extents are distinct views.

Cell lookup distinguishes missing, empty, exact value, formula, legacy array,
dynamic-array anchor/spill, data table, covered merge, and unknown states.
Cached and calculated values carry freshness/provenance separately. Exact stored
values never change merely because a number format resembles a date or currency.

Defined names are scoped typed expressions. Tables own a stable schema, range,
columns, totals, formulas, filters, and style, with optional validated typed row
projections. Formula ASTs retain dialect, lexical text, relative/absolute
anchors, and stable semantic targets. Shared-formula records are storage details.

Validation and conditional formatting use ordered data-bearing rules over stable
selections. Stop validation blocks changed values; advisory rules diagnose.
Filter visibility and manual hiding remain separate. Applying a sort is a
previewed reversible permutation, distinct from storing criteria.

Pivot definitions, caches, and rendered output are facets of one protected
object. Slicers/timelines are semantic controls, not shapes. Connections,
queries, and outputs are inert separate objects refreshed only with explicit
resolvers. Number formats use a lossless AST plus ergonomic builders.

## Presentations and DrawingML

Slides are added through validated semantic layout selectors. Duplication assigns
new local lineage identities and remaps references; review history requires an
explicit keep/drop policy. Deletion requires disposition for incoming behavior
references. Ordering is identity-relative, with checked numeric positions as a
convenience. Sections are contiguous partitions; custom shows are independent
selections.

Slide masters, layouts, notes masters, and handout masters are distinct types.
Resolved scenes include inherited layout/master content with provenance;
mutating inherited content explicitly edits its source or creates a local
override. Cross-presentation copy chooses source look or destination-theme
mapping and reports substitutions.

`Shape` is a non-exhaustive data-bearing semantic enum. Placeholders retain
placeholder identity and typed content. Shape properties, coordinate spaces,
z-order, accessibility reading order, fills, strokes, effects, and media are
typed. Raw numeric shape tags are not the facade. Accessibility reading order is
separate from z-order.

Transitions and timing are data-bearing trees. Common animations compile from a
short deterministic facade into the full timing tree. Media is inert and
explicitly embedded or linked. Speaker-note stories are separated from the full
notes scene.

Charts contain ordered data-bearing plots, family-specific series, typed axes,
typed data sources, separate caches, and explicit calculation/writeback. Classic
charts and ChartEx share semantic views without losing their native vocabulary.
2-D and 3-D are distinct short types. SmartArt remains a semantic diagram plus a
separate visual fallback; edits require an explicit layout provider or supported
defer-to-Office policy.
