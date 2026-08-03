# ADR 0020: Typed PPTX table-style ownership

- Status: Accepted
- Date: 2026-08-03

## Context

The OOXML migration host owned PresentationML table-style discovery and a
metadata-only parser while the concrete `litchi-pptx` crate had no complete
catalog or package mutation boundary. The old facade exposed long names,
treated display names as identities, and could not create, replace, or remove
the relationship safely. Its legacy presentation writer also recreated a
hard-coded Transitional `tableStyles.xml` edge, which could resurrect a
removed catalog, lose a producer's noncanonical target, or overwrite a Strict
relationship.

The checked-in ECMA-376 Strict and Transitional `dml-main.xsd` files define
`CT_TableStyle` as an ordered sequence of fourteen optional regions followed by
`extLst`. `styleId` and `styleName` are required, but `styleName` is an ordinary
string: empty and duplicate display names are valid. Microsoft producer
examples under `[MS-OE376]` contain empty names. Stable GUID identity must
therefore remain distinct from developer-facing name search.

## Decision

`litchi-pptx::table::style` is the sole owner of the bounded table-style XML
grammar, typed catalog, and OPC topology. Its short vocabulary is
`Conformance`, allocation-free `Id`, two-byte `Parts`, `Def`, and `List`.
`Id` accepts the required braced GUID form and formats a canonical uppercase
value. `Parts` represents `tblBg`, whole-table, row/column band, edge, and
corner regions in the schema's exact order; extension lists remain bounded
opaque content.

`List::get(Id)` is the primary stable selector, `at` is the checked raw-order
fallback, and `named` returns every match rather than silently choosing one
duplicate display name. `add`, `replace`, `rename`, `remove`, and
`set_default` validate before publication. A loaded list owns its XML once;
each definition retains checked ranges into that source. An unchanged
load-to-put consumes and moves the original allocation back to OPC. Renaming a
definition reconstructs only its typed wrapper and preserves its opaque cell
formatting body. `reset_parts` is deliberately named as a destructive operation
that replaces detailed formatting with selected empty region declarations.

Package `load`, consuming `put`, and `remove` accept all six presentation,
slideshow, and template main content types, with or without macros. They
validate one internal package main-document relationship, matching
PresentationML and table-style conformance families, one optional internal
catalog edge, the required content type, an inert relationship-free catalog
part, and package-wide inbound ownership. Orphan, shared, external,
wrong-content-type, mixed-dialect, and duplicate graph states fail before
mutation. `put` returns `false` for a byte or semantic no-op and retains
signatures; its semantic comparison ignores formatting whitespace only at the
known list/style container layers. Inherited `xml:space` and every text node in
deeper opaque formatting or extension payloads remain exact, so a meaningful
payload edit cannot be discarded as a no-op. `remove` is idempotent and returns
the moved prior `List`.

The migration host exposes only `styles`, `put_styles`, and `remove_styles`.
When its Transitional legacy writer materializes presentation changes, it
preserves a validated optional table-style relationship's exact ID, type, and
target, including absence and noncanonical targets. The generated slide-master
relationship ID is propagated into presentation XML instead of assuming
`rId1`. Strict table-style CRUD is supported, but Strict legacy presentation
materialization is refused before mutation until that larger writer is
conformance-aware. The old host parser, aliases, template accessor, and
duplicate table-style assets are deleted; new decks use the concrete owner's
deterministic default bytes.

## Consequences

- GUID identity, duplicate names, region presence, conformance, and optional
  graph ownership are explicit in the types and concise API.
- Catalog payloads move by value at mutation boundaries. Large unchanged XML
  has one list-owned allocation and is not copied per definition; no public
  runtime lock wrapper is introduced.
- Detailed cell/table formatting bodies remain opaque. `reset_parts` can author
  empty region shells, but a fully typed formatting editor is later work.
- A definition detached from producer XML can depend on namespace declarations
  held by its original list root. Moving such an exact opaque definition into
  an unrelated list can be rejected by the mandatory round-trip check; making
  detached definitions namespace-self-contained remains follow-up work.
- The legacy writer still rebuilds wider presentation topology. Preserving the
  catalog edge closes this ownership seam, but it is not a general lossless
  presentation transaction or a performance result.

## Verification

Owner tests cover GUID parsing, empty and duplicate names, the exact region
sequence including background and extensions, source-byte preservation,
semantic CRUD, all six main content profiles in both conformance families,
graph creation/replacement/removal/no-op, shared inbound refusal, root-dialect
agreement, malformed XML, schema-order rejection, and conservative opaque-text
no-op detection. Focused host tests cover new-package ownership, create/read/
update/remove, semantic no-op reporting,
absence, noncanonical target and relationship preservation across slide
materialization, dynamic relationship-ID propagation, and safe Strict-writer
refusal. The owner passes 71 unit tests, three doctests, and one compile-fail
test; the focused host integration passes all eight tests. Exact producer-asset
parity passes. Warning-denied Clippy is green for owner and focused host,
warning-denied rustdoc is green for the owner, and targeted formatting and diff
checks pass.

The `owner_native_smoke` example creates a slide and a two-row table, adds a
typed definition for the table's referenced GUID, stores background, whole-
table, and first-row region declarations, saves through the legacy writer
composition seam, reopens, and checks the exact typed definition. Desktop
PowerPoint for macOS opened the generated `table-style-owner.pptx` without a
repair prompt, rendered the table and its text, and exposed the native Table
Design and Table Layout tabs when it was selected. This is open-and-inspect
evidence for that Transitional artifact and empty-region definition only. No
PowerPoint edit/resave, reverse-read after an Office save, Strict native check,
application-version matrix, allocation, latency, or detailed-style rendering
claim follows.
