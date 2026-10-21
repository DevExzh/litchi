# 0360 — W4 XLSX and shared OOXML API contract

- Status: in progress on `uo2/w4-xlsx`; not accepted or merged into trunk.
- Date: 2026-10-21.
- Contract: [0351](0351-api-contract.md), source-backed shape [0355a §1](0355a-w1-docx-gaps.md).
- XLSB is tracked separately by record 0360b on `uo2/w4b-xlsb`.
- Performance: ABBA acceptance remains required; no performance claim is made here.

## Implemented migration

The branch contains the immutable XLSX package and output/signature routes,
record and part edit families, source-backed families with deferred output,
required-limit codec names, removal of `BoxedResult`, and typed limits for
pivot, XML maps, cell values, row visibility, shared OOXML and XLDM.
All consumers move with the corresponding API changes. The shared work also
provides DOCX task-pane, custom-XML, ribbon and VBA `LimitKind` exports and
makes DOCX `apply_durable_ink` return the ink family's `Commit`.

The branch merged 0361a trunk in `00b8074574`. The recorded merge resolution
for the ODP performance harness was stale: the resolution keeps trunk's
committed-owner route, as record 0360b requires.

## Final naming and gate registration

Borrowed XLSX editing handles now live in `workbook::edit::{Sheet, Tab, Row,
Column, Defaults, FormControl}` (also exported by the focused `edit` module).
They mutate their parent edit and do not publish independent snapshots.
The old `*Edit` handle names and root exports are deleted. Source-backed
patch inputs are `cell_values::{Change, SheetChange}` and
`row_visibility::Change`.

Shared Web Extensions custom-function input is `web::custom_function::Change`.
It is a semantic request carried by `WebIntent`, not an independent staging
handle. Its former `Edit` name is deleted; the durable planner and its tests
consume the new type.

Both XLSX glossary profiles register all 53 edit families (52 without VBA
inspection), the immutable package and read-only projections, and the contract
bans. Both shared OOXML profiles register `properties` and `custom_xml` with
`family_methods`. The new names and bans are recorded in `API_GLOSSARY.md`.

## Revision limits

Workbook revision failures now carry `Error::Limit` with
`LimitKind::Revisions(revisions::LimitKind)`, the observed count and the maximum.
The dimensions cover part bytes, XML depth/nodes, text bytes, users, headers,
logs, per-log and total records, and sheet identifiers. Resource failures no
longer use `Invalid` or a formatted resource string. Existing malformed-input
checks stay separate. The unified error classifier already maps `Limit` to
`limit_exceeded`.

An empty XML element now consumes the same depth budget as an explicit
start/end pair. A regression covers both encodings at the boundary and one
level beyond it. This closes a depth-check gap; it does not relax any bound.

## Naming and revision-limit slice verification

[The compact receipt](results/0360/naming-and-revisions.json) records the
commands and external log digests. All-target, all-feature checks and clippy
pass for XLSX, shared OOXML, ops, the facade and Python bindings. Library tests
pass (1,801 XLSX and 468 shared OOXML), as do 97 source-backed integration
tests, warning-free rustdoc and the performance harness compilation. All four
XLSX/shared OOXML glossary profiles pass, plus 35 glossary/matrix tool tests.
The fresh inventory has zero flagged rows; its coverage does not prove that
aliases or untyped resource failures are absent.

## Required work before acceptance

- Finish typed limits for every remaining family whose empty `LimitKind`
  still conceals an `Invalid` resource failure, including Data Model,
  external links, OLE objects, connections, data-type icons and rich-value
  refresh intervals. Check shared error kinds are preserved at each host.
- Remove remaining convenience aliases in changed XLSX/shared OOXML APIs,
  migrate their consumers, and check option accessors and retained limits.
- Complete the fresh public inventory and all glossary profiles; inspect
  every remaining row and preserve all capabilities.
- Complete crate/consumer tests, feature checks, clippy, rustdoc, harness,
  Python, corpus and boundary gates.
- Measure release ABBA for leaf and unified routes, including source-backed
  output and repeated saves; register evidence in the benchmark matrix.

These are engineering work under the owner decision, not accepted exceptions.
The redaction audit (memo item 22) alone remains deferred; changed diagnostics
must still follow its content-free rules.
