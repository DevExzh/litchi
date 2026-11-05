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

## Duplicate convenience entries

A second naming slice removes 58 documented duplicate methods and aliases.
Collection snapshots keep their existing canonical loaders and accessors;
hyperlink edits keep `put`, `remove` and `set`; drawings keep explicit lazy
payload reads; pivot snapshots expose lists through their typed values.
All affected tests and callers move with the removals. Neither the underlying
operation nor its output algorithm changes. The glossary bans every removed
entry at its caller-visible scopes (119 paths per full XLSX/shared profile
pair), including reexports.

[The alias receipt](results/0360/convenience-aliases.json) records a passing
all-target/all-feature workspace check, all four glossary profiles, 247
focused tests, clippy, warning-free rustdoc and harness compilation.

The additional exact-forwarder audit remains part of the final naming review;
constructors using `Default` and public entry points calling private helpers
are not duplicate public operations.

## Data-type-icon resource failures and depth enforcement

The family now exposes typed `LimitKind` dimensions for extension/part bytes,
XML events/depth and incoming relationships, carried by the crate's `Error::Limit`
with observed and maximum counts. Allocation failures preserve their source
through `Error::Allocation`. Empty extension input remains malformed input.

The extension reader, extension rewriter and part rewriter now all enforce
the declared depth ceiling for start and empty elements. Tests cover both XML
spellings exactly at the ceiling and one level beyond, preserving byte-exact
no-op output for accepted input. All 11 family tests pass, together with the
sheet-view tests, XLSX/ops all-target clippy, rustdoc and both XLSX glossary
profiles; commands and log digests are in the
[verification receipt](results/0360/data-type-icon-limits.json).

## P8–P12 typed-limit and consumer closure

The branch now closes the remaining shared OOXML typing work required by
P8–P12. Data Model, external links, OLE objects, connections, rich values and
refresh intervals expose explicit `LimitKind` dimensions with factual
`observed`/`maximum` values. Form-control owner limits use the same typed
resource field, including alternate-content choices and branches; malformed
input and checked-arithmetic overflow remain separate from resource refusal.
The shared MCE processor now distinguishes invalid caller bounds from an
exhausted bound, reports a saturated `usize::MAX` observation when arithmetic
overflow prevents an exact count, and records the ninth observed interrupted
read rather than the retry policy ceiling.

Every changed producer and consumer is migrated in this slice. XLSX chain,
connections, mail-merge, transition and spreadsheet-drawing boundaries retain
the typed MCE error instead of converting it to a diagnostic string. DOCX,
PPTX, DrawingML and the unified `litchi-ops` classifier consume the same
wrapper; invalid MCE bounds classify as `invalid_argument`, while exhausted
MCE resources classify as `limit_exceeded`. The owner mappings preserve
dimensions that have no XLSX equivalent as the shared typed MCE error rather
than relabeling them as another counter.

The corresponding raw/record limits, options-carried policy, accessor and
consumer migrations are included in the branch. The public API scan reports no
remaining public `Transaction`, `Mutable*`, `*_with_limits` or `BoxedResult`
in the XLSX/shared OOXML paths. Existing glossary and family/gate rows remain
registered; native release ABBA is deliberately still open and is not claimed
by this record.

The completed validation slice is:

- 1,812 XLSX library tests, including the repaired chain MCE boundary;
- 52 form-control unit tests, 18 form-control property tests, 39 owner-read
  tests, 32 scalar-lifecycle tests, 26 connections tests, 55 Data Model tests
  and 6 rich-value tests;
- 466 shared OOXML library tests, including the focused invalid-bound,
  retry-observation and four overflow-saturation checks;
- 32 unified error-classifier tests, 26 DOCX mail-merge tests, 1 PPTX
  transition test and 235 DrawingML tests;
- warning-denied all-target Clippy for shared OOXML, XLSX, DOCX, PPTX and
  unified ops (including all ops features), and workspace formatting and diff
  checks;
- both XLSX glossary profiles pass: 211 banned names, 52 minimal-feature
  families and 53 full-feature families, with five immutable owners and all
  registered family methods present.

An earlier pre-fix XLSX run had one expected-failure update at the chain
wrapper boundary; the final 1,812-test run is the completed zero-failure
result. The source-backed native performance and final ABBA acceptance remain
pending for the coordinator's separate measurement window.

## Required work before acceptance

- Complete the coordinator's final public inventory/glossary/gate sweep and
  inspect any rows outside the covered XLSX/shared OOXML paths.
- Complete rustdoc, Python, corpus and boundary gates for the integrated
  workspace state; the crate and consumer checks above are the W4 typing-slice
  evidence.
- Measure release ABBA for leaf and unified routes, including every new
  source-backed family, structural replay/inverse, retained patch history and
  repeated saves; register evidence in the benchmark matrix.

These are engineering work under the owner decision, not accepted exceptions.
The redaction audit (memo item 22) alone remains deferred; changed diagnostics
must still follow its content-free rules.
