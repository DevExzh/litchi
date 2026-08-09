# Format implementation review

Date: 2026-08-10

Revision under review: the uncommitted remediation worktree based on
`7432c18cd317eba10e0417c56595741a5c4dda58`.

This revision is authoritative for the current worktree. Earlier revisions remain useful
history for the defects already closed, but their scores, lint claims, and line/test counts
must not be read as current evidence.

## Scope and method

The score table covers every non-iWork format that appeared in the preceding review:
DOCX, XLSX, PPTX, DOC, XLS, XLSB, PPT, RTF, ODT, ODS, ODP, OpenDocument Formula,
ODC, ODG, ODM, ODI, OTH, ODB, and Markdown.

Pages, Keynote, Numbers, and the shared IWA stack are deliberately excluded from this
remediation score. Their code was not changed or regraded here, and no iWork score should
be inferred from this document.

Claims were checked against the live source and tests, ADRs 0001, 0003-0008 and 0023-0027,
and the checked-in normative material, notably `[MS-DOC]` section 2.7,
`[MS-XLS]`/`[MS-XLSB]` record enumerations, `[MS-PPTX]` sections 2.2 and 2.3,
`[MS-OE376]` page-break notes, RTF 1.9.1 appendix B, and OpenDocument 1.4 parts 2-4.
Documentation-only capability claims were not credited without corresponding code and
executable evidence.

Scores use the existing rubric:

- **Functional**: breadth and depth of read/write/edit support, preservation, real-producer
  fixtures, round-trip evidence, validation, and current gate status.
- **API conformance**: ADR-0003 snapshot/edit/commit/patch behavior, ADR-0004 semantic
  types, ADR-0005 finite I/O/resource behavior, ADR-0006 safety/losslessness, and relevant
  format ADRs.

A score of 95 means release-grade evidence across the claimed format surface, not merely a
large model or a green focused test. No current format has that evidence.

## Current scores

| Format | Functional | API conformance | Current evidence and principal limiter |
|---|---:|---:|---|
| DOCX | 90 | 81 | Deepest WordprocessingML coverage; new package-published paragraph transactions, but attached mutation, no durable format patch, and 4,452 production clippy errors |
| RTF | 88 | 75 | Very broad control-word model; smart quotes, HTML controls, shading, cell spacing, and paragraph edits landed; transaction remains single-operation and clippy has 2 production errors |
| XLSX | 84 | 88 | Best format-owned transaction/join model; page-break owner and budgeted history landed; no durable wire patch and 4,445 production clippy errors |
| PPTX | 84 | 75 | Broad part/extension coverage; typed `creationId`/`modId` transactions close the prior 2.2.9 gap; opened-presentation graph is still not generally editable and clippy has 1,267 production errors |
| ODT | 84 | 74 | Deep ODF text/package support and first durable deterministic patch integration; durable operation is whole-document replacement, legacy mutable APIs remain, and clippy has 4,301 production errors |
| DOC | 78 | 73 | Wide binary reader/writer and now versioned DOP through 2013; existing-file edits remain narrow and clippy has 4,418 production errors |
| XLSB | 78 | 78 | Broad BIFF12 read/write and substantially wider existing-cell scalar/cache/style-index edits; record coverage and edit families remain incomplete, with 2,074 production clippy errors |
| PPT | 77 | 79 | Strong binary read/create surface; length-changing simple shape text and durable patch integration are real; general opened-file structure/format edits remain absent |
| XLS | 77 | 73 | Wide BIFF8 coverage plus a real-fixture `Number` edit seam; most opened-workbook CRUD and many record families remain absent, with 4,557 production clippy errors |
| ODS | 76 | 72 | Worksheet, definitions, and RDF now have proper reversible transactions in addition to existing owners; rich spreadsheet feature CRUD is incomplete and clippy has 1,366 production errors |
| ODP | 71 | 68 | Real slide transactions plus owned embedded-chart and inert object inventory; table/form/style depth, encryption/signature lifecycle, and unified package editing remain incomplete |
| ODF Formula | 73 | 77 | Bounded caller-selected limits, real LibreOffice fixtures, and root transactions landed; the semantic model remains a small generic MathML tree without schema validation |
| ODC | 67 | 75 | Flat and packaged axis transactions plus whole typed chart replacement; no real ODC/FODC producer fixture or fine-grained chart-data CRUD |
| Markdown | 61 | 66 | Safe export with stronger escaping and DOCX semantics; still no import, and cross-source projection is incomplete |
| ODG | 49 | 64 | Real `.odg` ingress and richer layer/shape projection with three narrow edit verbs; most drawing structure/style/resource CRUD is absent |
| ODI | 47 | 59 | Normative frame/source structure and resource CRUD are materially improved; no real `.odi` artifact and semantic/edit breadth remain small |
| ODM | 44 | 56 | Title and linked-section edits are reversible; no real packaged `.odm` fixture and most master-document semantics are read-only |
| OTH | 42 | 53 | Pretty LibreOffice producer XML is exercised and multi-paragraph splices are reversible; typed text-web semantics/resources remain very narrow |
| ODB | 41 | 54 | Real `.odb` catalog evidence now covers connections, keys, indices, forms/reports inventory, and one query edit; it is not a database/schema CRUD surface |

## Workspace-wide findings

### Build and test state

The combined non-iWork all-target/all-feature compile is green. Focused tests for every
material remediation claim are also green; exact commands are recorded below. This is
strong evidence that the wave is internally coherent, but it is not a substitute for full
per-crate suites, native application reopen/resave, or warning-denied lint.

The current strict-lint result is materially worse than the preceding review reported.
The following command was run separately for each scored crate:

```text
cargo clippy -p <crate> --all-targets --all-features --no-deps -- -D warnings
```

Green: `litchi-ppt`, `litchi-odp`, `litchi-odf-formula`, `litchi-odc`,
`litchi-odg`, `litchi-odm`, `litchi-odi`, `litchi-oth`, `litchi-odb`, and
`litchi-markdown`.

Red production libraries, measured again with `--lib`, are:

| Crate | Production errors | All-target result |
|---|---:|---|
| `litchi-docx` | 4,452 | lib-test ends with 7,997 errors |
| `litchi-xlsx` | 4,445 | lib-test ends with 8,777 errors |
| `litchi-pptx` | 1,267 | lib-test ends with 1,346 errors |
| `litchi-doc` | 4,418 | lib-test ends with 6,965 errors |
| `litchi-xls` | 4,557 | lib-test ends with 6,827 errors |
| `litchi-xlsb` | 2,074 | lib-test ends with 4,085 errors |
| `litchi-rtf` | 2 | the same 2 errors repeat for lib-test |
| `litchi-odt` | 4,301 | lib-test ends with 5,676 errors |
| `litchi-ods` | 1,366 | lib-test ends with 1,660 errors |

The RTF failures are specifically `match_same_arms` in
`codec/writer/codec/output/content.rs` and `shadow_same` in `edit.rs`.
The larger crates fail crate-wide denied lints such as missing error documentation,
shadowing, numeric casts, and `unwrap`/`expect`; they are not isolated to this remediation.
ADR-0008's continuously lint-clean target therefore remains unmet for nine scored formats.

### ADR-0003 common patch status

The shared `litchi-core::patch` layer is no longer merely an envelope. It now has bounded
canonical JSON, reversible/forward-only modes and `seal`, blob bundles, fingerprints,
budgeted `History<T>`, exact-key `SubEdit` composition, deterministic conflict sets, and a
non-mutating three-way plan. Twenty focused patch tests pass.

Adoption is still narrow:

- ODT converts exact whole-package before/after artifacts into the shared durable wire form.
- PPT converts its focused shape-text operation to a semantic durable operation with exact
  artifact and expected-text preconditions.
- XLSX reuses only the shared budgeted `History<Workbook>` type; its format patch remains
  in-memory.
- The other scored formats do not integrate the common durable/composition layer.

No format currently demonstrates end-to-end durable semantic patches plus independent
sub-edit join, conflict resolution, three-way merge, cross-document dependency transfer,
and history. Generic infrastructure alone does not satisfy ADR-0003 for a format.

### Global compact/minimal XML requirement

The static producer-asset work is sound but the global publication claim is not yet true.

What is verified:

- `xml-minifier`'s asset registry covers 37 readable/generated OOXML asset pairs plus the
  three explicitly registered compact DOCX assets: all 77 checked-in XML files under
  production `src/` trees are covered. The parity/registry tests pass.
- The bounded auditor rejects indentation containing CR/LF/tab, padded attributes,
  whitespace before closes, DTDs, malformed XML, invalid encodings, and limit overruns. It
  preserves character data, CDATA, comments, processing instructions, and inherited
  `xml:space="preserve"` content.
- ODT's changed-package audit includes `.xml`, `.rdf`, and manifest-declared XML media types.
  Several newer focused transactions also emit byte-minimal changed fragments.

Why the global requirement still fails:

1. Both shared package writers accept arbitrary part bytes. `litchi-opc::PackageWriter`
   writes every `Part::blob()` without an XML audit; ODF `PackageWriter::add_file` likewise
   has no central compactness gate. Format-local tests therefore cannot prove every dynamic
   XML publication path.
2. Both general auditors intentionally accept plain space-only text nodes such as
   `<p>   </p>` and `<p><b>a</b> <i>b</i></p>`. That is correct for semantic whitespace,
   but it also means a schema-neutral audit cannot distinguish semantic spaces from an
   extra inter-element space in element-only content. Several feature matrices explicitly
   concede that absolute minimality is not guaranteed. A 95-level claim needs
   schema/content-model-aware classification or construction-by-provenance, not a blanket
   deletion of space nodes.
3. Several ODF rewrite guards inspect only paths whose extension is `.xml` and copy all
   other auxiliary members. This includes ODG, ODI, ODM, ODC, and ODB package snapshot
   paths, and ODP chart/edit audits. Consequently `manifest.rdf` or another
   `application/rdf+xml` member can be emitted unchanged without audit. ODS's focused RDF
   transaction validates the graph it authors and `META-INF/manifest.xml`, but that does
   not close every other package publication path.
4. The OOXML asset test scans literal production `include_str!` paths ending in `.xml`.
   It does not prove runtime XML assembled by format writers or arbitrary caller-provided
   XML parts. Generated `.rels` and `[Content_Types].xml` are compact today, but the common
   OPC writer has no package-wide assertion covering every XML content type.

The required fix is a bounded, content-type-aware package publication audit shared by OPC
and ODF, including `.xml`, `.rels`, `.rdf`, `[Content_Types].xml`, manifest-declared XML
media types, and signature XML when generated. It must preserve semantic whitespace and
refuse—or prove source-preserved—unclassifiable formatting whitespace. Real package tests
must enumerate every emitted XML-bearing member, including `manifest.rdf`.

## Per-format assessment and minimum path to 95

### DOCX — 90 / 81

The new `document::{Snapshot, Edit, Commit, Patch}` is real: it performs bounded
length-changing simple-paragraph replacement and insertion, readback, exact-source apply,
inverse, and atomic package publication. Ordered paragraph/run projections also retain
unknown inline XML and avoid leaking relationship IDs. These close important earlier gaps.

Before 95 is defensible: make the document transaction the ordinary editing root and retire
the attached `MutableDocument`; cover tables, nested blocks, multi-run text, fields,
hyperlinks, revisions, controls, and cross-part dependencies; adopt durable semantic patches,
join/merge/history; eliminate the 4,452 production lint errors and remaining deliberate
panic sites; enforce package-wide minimal XML; and add native Word open/resave evidence for
the new edit paths, including unknown-content preservation.

### XLSX — 84 / 88

The page-break owner correctly provides checked row/column collections and reversible
package transactions, and its real POI manual-break test passes. `History<Workbook>` is now
budgeted through the common layer. The established workbook edit/join/conflict model remains
the best ADR-0003 implementation in the workspace.

Before 95: serialize its semantic patches through the common durable vocabulary; connect
history to commits and prove undo/redo budgets; implement format-owned three-way merge and
cross-document dependency transfer; close remaining typed package/write gaps (including
read-only page setup and incomplete modern formula/chart families); eliminate 4,445
production lint errors; add package-wide XML minimality; and provide current Excel
open/resave evidence for the expanded owners.

### PPTX — 84 / 75

`change_tracking::{Snapshot, Edit, Commit, Patch}` now models `[MS-PPTX]` 2.2.9/2.3
`creationId` and `modId`, validates shape-ID uniqueness, preserves unknown extension XML,
supports Strict namespaces, and is exercised on a real PowerPoint fixture. This closes the
previous extension-family hole.

Before 95: unlock a unified transaction over opened presentations for slide/shape/text,
tables, charts, notes, and relationship dependencies; replace whole-graph clone rollback
and ordinary `&mut` mutation; adopt durable join/merge/history; finish partial extension
semantics rather than inventory only; eliminate 1,267 production lint errors; enforce all
XML-part minimality; and add native PowerPoint edit/resave evidence.

### DOC — 78 / 73

Versioned document properties now cover the normative DOP chain through `Dop2007`,
`Dop2010`, `Dop2013`, and `DopMth`, whose sizes/roles match `[MS-DOC]` 2.7.8-2.7.10 and
2.7.17. The existing body-text transaction remains restricted to a single Unicode piece
and equal UTF-16 length.

Before 95: provide general opened-document edits for length-changing text, formatting,
tables, all stories, fields/revisions, images, and dependency closure; unify the many
feature editors under one immutable transaction root; add durable join/merge/history;
eliminate 4,418 production lint errors and panic debt; and expand real Word fixture/native
resave coverage across DOP generations and edit paths.

### XLS — 77 / 73

The new `cell_values` owner edits an existing BIFF8 `Number` record's eight-byte Xnum field,
reopens the complete package, preserves every other CFB stream, and reverses exactly on a
real POI workbook. It is a useful but intentionally narrow seam.

Before 95: cover RK/MulRK, strings/SST, Boolean/error/blank, formula caches/tokens, styles,
cell creation/removal, rows/columns/sheets, and calculated dependency closure on opened
workbooks; integrate the existing feature editors and durable patch ecosystem; materially
raise the roughly three-fifths typed record coverage and writer formula coverage; eliminate
4,557 production lint errors; and obtain native Excel round-trip evidence.

### XLSB — 78 / 78

The cell transaction now inventories ordinary scalar and cached-formula families and can
edit Xnum, exact RK, Boolean, typed errors, shared-string indexes, equal-length strings, and
24-bit style indexes while retaining storage families. Three real-fixture tests pass.

Before 95: support cell creation/removal, length-changing/rich strings, formula-token and
style-resource edits, structural sheet CRUD, and the still-absent BIFF12 families (data
model/rich values/MDX/smart tags/ActiveX); integrate durable join/merge/history; eliminate
2,074 production lint errors; audit all XML host parts; and expand real/native XLSB evidence.

### PPT — 77 / 79

Length-changing single-atom shape text now updates enclosing record framing and simple style
coverage, and the focused patch can round-trip through the common deterministic JSON
envelope. The crate is warning-denied lint-clean.

Before 95: make opened-presentation edits general across slides, shapes, formatting, tables,
charts, media, masters, and relationships; remove the dual trait-object/tagged shape API;
integrate durable composition/merge/history beyond shape text; finish chart authoring;
broaden real fixtures and add native PowerPoint resave evidence; and remove residual public
panic assumptions.

### RTF — 88 / 75

The remediation closes the previously identified appendix-B gaps for smart quotes,
HTML controls, paragraph/character shading patterns, and cell spacing aliases, and adds a
checked paragraph-text transaction. The focused transaction tests pass.

Before 95: allow multiple composable span/property/structure edits rather than one semantic
operation; retire the large attached `&mut` model; provide durable patches, join/merge, and
history; finish known destination/group internals and writer-option behavior; fix the two
production lint errors; and add broader Word/LibreOffice native round-trip evidence.

### ODT — 84 / 74

ODT is the first ODF owner with shared durable deterministic patches. Parsing, inversion,
sealing, source-digest checks, and noncanonical-JSON rejection are tested. Builder metadata
no longer consults the ambient clock.

Before 95: encode semantic operations rather than whole-document blob replacement; adopt
sub-edit composition, conflict resolution, three-way merge, transfer, and history; delete
legacy attached mutable roots/raw index verbs; eliminate 4,301 production lint errors and
the large panic-macro backlog; make XML/RDF minimality central across every package path;
and add current LibreOffice native edit/resave plus encrypted/signed lifecycle evidence.

### ODS — 76 / 72

Worksheet structure/cell CRUD, named definitions, and RDF graph/triple edits now have
source-bound snapshots, checked selectors, batched commits, readback, exact no-ops, inverse
patches, and mutable-facade adapters. All nine focused tests pass.

Before 95: consolidate all spreadsheet mutations behind one package transaction and remove
the mutable facade; implement package CRUD for rich text, styles, conditional formatting,
sparklines, hyperlinks/images/shapes, formulas and dependency closure; add durable
composition/merge/history; expose encryption/signature lifecycle; eliminate 1,366
production lint errors; fix XML/RDF publication globally; and add broad real/native Calc
round trips.

### ODP — 71 / 68

Owned chart snapshots now add/remove/replace chart parts with semantic readback and exact
patch rehydration; embedded object inventory covers objects, OLE payloads, applets, plugins,
and floating frames inertly. The crate remains lint-clean.

Before 95: unify chart, slide, master, annotation, and RDF owners into one package edit;
provide typed tables, forms/controls, rich text/style graph, fine chart-data editing, and
resource dependency closure; replace tagged optional-field shapes; integrate durable
join/merge/history; support/test encryption writes and signatures; close XML/RDF audit gaps;
and add native Impress resave evidence.

### ODF Formula — 73 / 77

Transactions, configurable hierarchical limits, deterministic namespace-prefix assignment,
and two real LibreOffice `.odf` fixtures are verified. The crate is lint-clean.

Before 95: validate MathML content models/arity/value domains against the checked-in schema;
add granular semantic edits and durable composition/history rather than whole-root replace;
model StarMath rather than retaining annotation text only; prove every output XML member is
minimal; add malformed/property/fuzz coverage and native LibreOffice resave evidence.

### ODC — 67 / 75

Packaged axis edits and explicit complete chart-definition replacement now publish atomically
and reverse exactly while retaining auxiliary payloads. The focused tests are green and the
crate is lint-clean.

Before 95: add real ODC and FODC producer fixtures; implement granular plot/series/axis/data
and style/resource CRUD; validate formula/range grammar; expose caller-selected limits;
adopt durable join/merge/history; add encryption/signature lifecycle; fix all XML/RDF package
audits; and obtain native chart round trips.

### ODG — 49 / 64

Package semantics now distinguish page-local/global layers and expose geometry,
accessibility, styles, and z-order; shape text/name/layer edits are reversible. A real
LibreOffice `.odg` opens byte-exactly. The crate is lint-clean.

Before 95: implement structural page/layer/shape/group CRUD, geometry/style/resource/media/
form ownership, templates, encryption/signatures, one unified transaction, durable
composition/history, complete XML/RDF auditing, substantially broader malformed/producer
fixtures, and native Draw edit/resave evidence.

### ODI — 47 / 59

The crate now enforces the ODF image family structure, authors a valid baseline image,
models accessibility/geometry, and supports package-local resource CRUD with reversible
patches. It is lint-clean.

Before 95: add real `.odi` producer artifacts; model image maps, styling, metadata, resource
graphs and broader frame semantics; unify flat/package edits; adopt durable merge/history;
support encryption/signatures; close `.rdf`/XML audit gaps; and obtain native producer
round-trip evidence.

### ODM — 44 / 56

Linked-section source attributes and placement are now validated, and exact-name/position
edits complement title edits with source checks and inverse patches. The crate is lint-clean.

Before 95: add real packaged `.odm` fixtures (the current master fixture is XML packaged in
tests); model and edit section trees, subdocument/resource graphs, styles and full metadata;
unify title/link edits; adopt durable merge/history; support encryption/signatures; audit
all XML/RDF members; and prove native master-document resaves.

### OTH — 42 / 53

The text-web projection now covers headings, paragraphs, styles, links, and ODF whitespace;
multi-paragraph edits are atomic, reversible, and exercised against pretty LibreOffice
template source XML. The crate is lint-clean.

Before 95: use a real packaged `.oth` producer artifact; model lists, bookmarks, fields,
formatting, resources/forms/embedded objects and full metadata/styles; provide structural
transactions plus durable merge/history; support encryption/signatures; audit all emitted
XML/RDF; and obtain native Writer/Web resave evidence.

### ODB — 41 / 54

Catalog reads now include connections, validated column types/constraints, keys, referential
actions, indices, and inert forms/reports. One stored-query command/escape-processing edit
has exact golden XML, no-op, stale/refusal, inverse, and real noncompact-package refusal
tests. The crate is lint-clean.

Before 95: implement schema/table/column/key/index/query/connection and component CRUD in a
unified transaction; add relations and complete producer extensions; provide durable
merge/history; support encryption/signatures; handle real LibreOffice packages without
requiring compact-source refusal where lossless editing is possible; audit XML/RDF globally;
and add native Base round-trip evidence without executing database content.

### Markdown — 61 / 66

CommonMark escaping is stronger, DOCX mapping now uses ordered inline/run semantics, and
ambiguous drawing/link placements are typed refusals. Leaf tests and the DOCX semantic
golden tests pass.

Before 95: either implement a bounded lossless Markdown reader/editor or explicitly revise
the product rubric to define this crate as export-only; complete equivalent projections for
DOC, RTF, ODT, Pages, and the other advertised sources; convert rather than merely refuse
links, images, footnotes, fields, tables, quotes, and code where representable; add a large
real-fixture/golden CommonMark corpus and parser-compatibility tests; and expose a coherent
snapshot/transaction story if Markdown itself becomes editable.

## Commands and exact results

Current decisive gates:

```text
cargo check -p litchi-docx -p litchi-xlsx -p litchi-pptx -p litchi-doc \
  -p litchi-xls -p litchi-xlsb -p litchi-ppt -p litchi-rtf -p litchi-odt \
  -p litchi-ods -p litchi-odp -p litchi-odf-formula -p litchi-odc \
  -p litchi-odg -p litchi-odm -p litchi-odi -p litchi-oth -p litchi-odb \
  -p litchi-markdown --all-targets --all-features
# exit 0

cargo test -p xml-minifier --test audit --test ooxml_assets
# audit: 12 passed; assets: 4 passed, 1 ignored

cargo test -p litchi-core patch
# 20 passed, 112 filtered out

cargo test -p litchi-docx document::transaction::tests
# 4 passed
cargo test -p litchi-docx paragraph::tests
# 32 passed
cargo test -p litchi-xlsx --test xlsx_page_breaks_owner
# 4 passed
cargo test -p litchi-xlsx --test xlsx_history
# 1 passed
cargo test -p litchi-pptx change_tracking::tests
# 8 passed
cargo test -p litchi-doc opened_document_exposes_versioned_document_properties
# 1 passed
cargo test -p litchi-xls --test xls_cell_values
# 1 passed
cargo test -p litchi-xlsb --test cell_value_edit
# 3 passed
cargo test -p litchi-ppt text_edit::tests
# 7 passed
cargo test -p litchi-rtf --test transactions
# 7 passed

cargo test -p litchi-odt --test packaged_transactions
# 10 passed
cargo test -p litchi-ods --test worksheet_transactions \
  --test definition_transactions --test rdf_transactions
# 9 passed
cargo test -p litchi-odp charts::tests
# 8 passed
cargo test -p litchi-odf-formula --test transactions_and_limits
# 8 passed
cargo test -p litchi-odc --test package_edit
# 3 passed
cargo test -p litchi-odg --test package_semantics
# 7 passed
cargo test -p litchi-odm --test linked_sections
# 6 passed
cargo test -p litchi-odi --test semantic_api
# 10 passed
cargo test -p litchi-oth --test semantic_api
# 16 passed
cargo test -p litchi-odb --test query_transactions
# 4 passed

cargo test -p litchi-markdown
# 19 unit passed; 18 doctests passed, 3 ignored
cargo test -p litchi --test markdown_docx_semantics --features markdown,docx
# 3 passed
```

The final scoped `git diff --check -- docs/FORMAT_IMPLEMENTATION_REVIEW.md` is green.

## Bottom line

The remediation wave closes real gaps: DOCX has a genuine main-document transaction;
XLS now edits a real BIFF8 cell; XLSB edits many more scalar/cache families; PPT supports
bounded length-changing text and durable patches; PPTX closes change-tracking IDs; RTF
closes several appendix-B controls; ODT adopts durable patches; and every small ODF owner
gained meaningful semantics or transactions.

It does not justify 95 for any format. The largest blockers are: nine red crate-wide lint
gates; narrow or fragmented opened-file transaction surfaces; shared durable patch adoption
in only two formats (plus XLSX history); missing format-level merge/transfer workflows;
thin or absent real/native fixture evidence for the small ODF families; and the unclosed
global XML publication rule, especially schema-neutral space nodes and unaudited
`manifest.rdf`/XML auxiliary members.
