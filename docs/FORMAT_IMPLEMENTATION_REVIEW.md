# Format implementation review

Date: 2026-08-10

Revision under review: the current uncommitted worktree based on
`7432c18cd317eba10e0417c56595741a5c4dda58`.

This is the authoritative review for the current worktree. Earlier revisions are retained in
version history but their scores, lint status, XML conclusions, and test counts are superseded.

## Scope and scoring

This review covers exactly these 19 formats: DOCX, XLSX, PPTX, DOC, XLS, XLSB, PPT, RTF, ODT,
ODS, ODP, OpenDocument Formula, ODC, ODG, ODI, ODM, OTH, ODB, and Markdown.

All iWork formats and the shared IWA implementation are excluded. No iWork code was inspected,
scored, or changed, and no iWork result is used as evidence here.

The review checked the live diff and public APIs, the format feature matrices,
`docs/CRUD_Scenario_Checklist.md`, ADRs 0001, 0003-0006, 0008, and 0023-0027, relevant checked-in
specifications, and executable tests. Scores mean:

- **Functional**: public read/create/opened-file edit breadth, preservation, validation, real
  producer evidence, and production readiness.
- **API conformance**: immutable snapshots, checked semantic selectors and values, atomic commits,
  reversible and durable patches, deterministic composition/conflicts, bounded history and I/O,
  dependency closure, and lossless-or-refuse behavior required by ADR-0003 through ADR-0006.

`95` requires release-grade evidence across the claimed surface, including the applicable CRUD
checklist, warning-denied gates without material correctness quarantines, real producer and native
edit/resave evidence, and end-to-end XML/package proof. A deep reader, a large test count, or one
excellent transaction owner is not sufficient. No current format reaches 95.

## Authoritative scores

| Format | Functional | API conformance | Principal current limiter |
|---|---:|---:|---|
| DOCX | 92 | 84 | Broadest WordprocessingML surface and clean direct production Clippy, but main-document editing is narrow and neither durable nor composable |
| XLSX | 87 | 89 | Best general workbook transaction/join surface; durable wire, format merge/transfer, and integrated history remain absent, with broad lint quarantines |
| PPTX | 86 | 79 | Deep package and extension coverage; opened-presentation mutation remains fragmented and non-durable |
| DOC | 79 | 74 | Deep binary read/write inventory; ordinary existing-body edit is still equal-length/single-piece and lint cleanliness relies on broad allowances |
| XLS | 78 | 74 | Broad BIFF8 model and focused existing-record editors; no general opened-workbook CRUD and broad lint allowances remain |
| XLSB | 80 | 80 | Wide BIFF12 coverage and expanded existing-cell edits; structural CRUD and several modern record families remain incomplete, with the broadest lint quarantine |
| PPT | 83 | 88 | Real durable/composable/history-backed slide ordering plus durable text edits; the rest of opened-presentation editing is still fragmented |
| RTF | 89 | 78 | Very broad RTF semantics with a panic-denied production policy; editing remains a one-operation seam over a retained mutable raw model |
| ODT | 86 | 78 | Deep ODF Text support and durable exact-artifact patches; semantic durable operations, composition, and attached-mutable retirement remain open |
| ODS | 79 | 76 | Strong focused worksheet, definition, chart, RDF, protection, and tracked-change owners; no unified workbook transaction or rich spreadsheet CRUD |
| ODP | 78 | 78 | Unified slide/shape/media/RDF publication is real; charts, masters, annotations, styles, tables, forms, and security remain separate or partial |
| ODF Formula | 83 | 87 | Bounded MathML validation, granular edits, durable patch chains, and real formula fixtures; still a bounded presentation-MathML subset without StarMath semantics |
| ODC | 79 | 85 | Granular detached chart CRUD/history plus packaged style/resource transactions; no genuine ODC/FODC producer package or durable/merge layer |
| ODG | 65 | 77 | Unified page/layer/shape/geometry/style/resource package edits now work; drawing semantics, durable conflicts, and security remain thin |
| ODI | 60 | 73 | Shared flat/package frame edits, image maps, resource graph, metadata, and history are real; producer evidence and semantic breadth remain small |
| ODM | 61 | 86 | Excellent durable merge/history behavior for title and link targets; most master-document structure is still read-only and the producer fixture is transformed before ingress |
| OTH | 62 | 76 | Real LibreOffice template ingress, richer semantics, disjoint join, and history; structural/resource mutation and durable patches remain narrow |
| ODB | 64 | 74 | Unified inert schema/query/connection/component CRUD on real pretty LibreOffice XML; no durable merge and no database runtime/security lifecycle |
| Markdown | 74 | 80 | New bounded exact-source CommonMark/GFM reader and reversible block edits close the import gap; the semantic model and transaction remain top-level/single-operation only |

## Cross-cutting findings

### Build, tests, and strict lint

The combined all-target/all-feature check for all 19 formats is green. The integrated full-suite
run also reports green format libraries, including DOCX 824, XLSX 712, PPTX 446, DOC 942 with two
ignored, XLS 963, XLSB 508, PPT 1,040 with one ignored, ODT 515, ODS 126, RTF 287 plus integrations,
shared OOXML 207, and DrawingML 116. This review independently reran the new high-risk format tests
and the shared publication tests recorded below; all passed.

The following production command is green for all 19 format crates and the reviewed shared crates
`litchi-core`, `litchi-opc`, `litchi-ooxml-common`, `litchi-odf-common`, and `xml-minifier`:

```text
cargo clippy -p <selected/common crates> --lib --all-features --no-deps -- -D warnings
```

That is a genuine warning-denied production-library gate, but it is not equivalent to direct
remediation. `-D warnings` does not override explicit `#![allow(...)]` attributes:

- DOCX has no broad production Clippy quarantine. PPT and RTF have narrow organization allowances;
  RTF additionally denies panic, indexing, string slicing, `unwrap`, and `expect` lints in
  non-test builds. Most small ODF crates and Markdown use only narrow, reasoned exceptions.
- XLSX, DOC, XLS, XLSB, ODT, and ODS use broad crate-level allowances. Several suppress narrowing
  casts, `unwrap`/`expect`, ignored error results, shadowing, missing error/panic documentation, or
  large sets of style lints. XLSB and ODT have especially large quarantines.
- PPTX has a broad schema/API/style quarantine, though not a production `unwrap`/`expect`
  quarantine. `litchi-ooxml-common` also suppresses several parser/error/cast lints and affects all
  OOXML hosts.

Therefore the old thousands-of-errors lint blocker is closed as a gate, but residual lint debt is
not zero. For 95, every correctness-relevant allowance must be replaced by local proof/direct fixes
or narrowed to generated/spec-shaped code with executable invariants. The all-target Clippy and
rustdoc portions of the CRUD release checklist also need one reproducible final run; this review's
independent strict command was production `--lib` only.

### XML byte-minimal publication

The prior global XML defect is materially remediated.

`xml-minifier::audit::verify_authored` now rejects indentation, padded markup, space before closes,
DTD/custom entities, and ambiguous all-space text nodes outside `xml:space="preserve"`. Ordinary
`verify` still accepts semantic or source-preserved space-only nodes, which is the correct read-side
behavior. Both modes have finite byte/event/attribute/token/text/depth limits.

The shared publication boundaries now classify XML by `.xml`, `.rels`, `.rdf`,
`[Content_Types].xml`, `application/xml`, `text/xml`, and `+xml` media types:

- OPC audits generated content types and relationships and every authored or changed XML-bearing
  part. Only a byte-identical XML part captured from the opened source is exempt.
- ODF audits every authored/changed XML-bearing file plus generated `META-INF/manifest.xml`.
  `copy_auxiliary_files_from` bypasses the authored audit only for exact source members, including a
  pretty `manifest.rdf`.
- Raw negative tests reject arbitrary authored `.rdf`, manifest-declared XML, and signature-XML
  payloads. Enumeration tests prove coverage of content types, relationships, manifest XML,
  `manifest.rdf`, media-type-only XML, and signature XML. Static OOXML asset parity remains green.

This is the required provenance distinction: byte-minimal for authored/changed bytes, exact for
untouched source bytes. Semantic spaces must be actual character data or explicitly protected by
`xml:space="preserve"`; a schema-neutral publisher no longer guesses.

One exception remains visible. ODB's existing-package transaction directly rebuilds ZIP output and
byte-splices compact generated fragments into exact producer `content.xml` so that pretty source
formatting survives. Its private operations, full reparse, compact extension validation, and real
pretty-fixture test make the current seam defensible, but it bypasses the shared ODF writer's typed
origin boundary. Before a global 95 claim, this path should use a shared provenance-bearing splice
publication API and have raw negative tests for every authored fragment class.

Several small-format feature matrices are stale about XML: ODC/ODI/ODM still say space-only
authored nodes can pass even though the final shared writer now refuses them. Those rows must be
updated before certification.

### ADR-0003 durable patches, composition, and history

The common patch layer is substantial: bounded deterministic JSON, reversible/forward-only modes,
blob bundles, fingerprints, deterministic conflict sets, sub-edit composition, three-way planning,
and budgeted history are implemented and tested. Format adoption is no longer limited to two owners:

- PPT slide ordering has semantic durable operations, independent disjoint joins, structured
  overlap conflicts, exact inverse, and budgeted history; focused PPT text also has durable patches.
- ODM has deterministic durable exact-artifact patches, sealing, same-source semantic merge with
  conflicts, inverse application, and bounded history.
- ODF Formula has granular semantic operations and bounded durable patch chains/history.
- ODT retains a durable whole-document replacement vocabulary.
- ODC/ODG/ODI/OTH/ODB and XLSX have useful composition or history subsets, but not the full durable
  format contract.

No format yet demonstrates the entire ADR-0003 product contract across its ordinary editing root:
durable semantic operations, independent sub-edit join, non-mutating three-way conflict planning,
dependency-closure cross-document transfer, and budgeted undo/redo. Generic infrastructure does not
raise a format to 95 until the format owns and tests those workflows.

### Fixture and documentation quality

Real fixture evidence improved: PPT slide ordering uses a real PPT; ODG opens a real LibreOffice
ODG; OTH opens and edits a LibreOffice 24.2 Writer/Web template; ODB minimally edits a real pretty
LibreOffice database; ODF Formula has real `.odf` inputs; and ODC reads two real LibreOffice chart
subdocuments.

Important limitations remain:

- ODC fixtures are charts extracted from FODS/FODT, not genuine ODC/FODC producer artifacts.
- The ODM test reads a genuine LibreOffice `.odm` as test data, compacts selected XML parts, and
  repackages them before `Master::from_bytes`; it does not prove original-package ingress or
  byte-preserving edits against the producer artifact. Its matrix currently overstates this.
- No checked-in real ODI producer artifact exists.
- Markdown has a useful synthetic corpus but not the CommonMark specification examples, a broad GFM
  compatibility corpus, or real round-trip fixtures.
- Most formats still lack current native Office/LibreOffice Litchi-edit/resave/reopen evidence for
  the newly added transaction paths. Old native evidence certifies only the exact historical slice
  it exercised.

## Exact minimum remaining path to 95

| Format | Minimum remaining work before 95 is defensible |
|---|---|
| DOCX | Make one immutable package/document transaction the ordinary root; cover complex runs, tables, fields, hyperlinks, revisions, controls, styles/resources, and dependency closure; add durable semantic join/merge/history/transfer; retire attached mutation and panic assumptions; run all-target lint/rustdoc; add adversarial and current native Word edit/resave evidence. |
| XLSX | Put the existing workbook operations on the common durable wire; connect commits to bounded history; add three-way merge and dependency-closure copy/move; close remaining modern formula/chart/rich-value/write gaps; narrow correctness lint allowances; prove all XML parts and current Excel edit/resave workflows. |
| PPTX | Provide one opened-presentation transaction spanning slides, shapes/text, notes, tables, charts, masters/layouts, media, comments, and relationships; replace broad mutable/clone rollback paths; add durable join/merge/history/transfer; narrow allowances; add raw XML negatives and current PowerPoint edit/resave proof. |
| DOC | Replace equal-length/single-piece text editing with lossless general story, formatting, table, field/revision, drawing, and resource CRUD; unify feature editors; add durable composition/history/transfer; localize cast/unwrap proofs; expand versioned real DOC and native Word resave coverage. |
| XLS | Add opened-workbook cell create/remove and all storage families, formula/style/string/resource edits, row/column/sheet structure and dependency closure; unify focused owners; add durable merge/history; narrow legacy lint allowances; provide broader BIFF record and native Excel evidence. |
| XLSB | Add structural workbook/sheet/cell CRUD, length-changing rich/shared strings, formula tokens/caches/styles, and missing modern record families; add durable merge/history/transfer; drastically narrow the crate lint quarantine; enumerate XML host parts and add current native XLSB evidence. |
| PPT | Extend the durable slide-order/text model to general slides, shapes, formatting, tables, charts, media, masters, comments, and relationships; remove dual shape APIs; add format-wide three-way merge/transfer; finish chart authoring; add broader real fixtures and native PowerPoint edit/resave evidence. |
| RTF | Support multiple composable span/property/structure edits, durable patches, merge, and history; retire attached mutable raw editing from the ordinary path; complete destination/writer option coverage; add large hostile/producer corpora and Word/LibreOffice edit/resave evidence. |
| ODT | Replace whole-artifact durable operations with semantic ones; add join/three-way merge/history/transfer at the package root; retire `MutableDocument` and raw-index verbs; narrow the large lint quarantine; test signed/encrypted edit policy and current LibreOffice edit/resave workflows. |
| ODS | Unify worksheet, definitions, charts, annotations, RDF, protection, DataPilot, and tracked changes in one package transaction; add rich text/style/formula/conditional-format/sparkline/drawing/resource dependency CRUD; add durable merge/history/transfer and security lifecycle; narrow lint allowances; add broad Calc resaves. |
| ODP | Bring charts, masters/layouts, annotations, styles, tables/forms, and security into the unified slide/shape/media/RDF transaction; add fine chart/table/text CRUD and durable merge/history/transfer; test encryption writes/signatures; add real complex decks and native Impress resaves. |
| ODF Formula | Complete the checked MathML content/value model and schema corpus; define and parse StarMath semantics or keep it explicitly opaque; add independent join/three-way merge and transfer; exercise actual package member publication and malformed/fuzz boundaries; add native Math edit/resave evidence. |
| ODC | Add genuine ODC and FODC producer artifacts; connect granular definition edits to fine-grained opened-package chart operations; add durable semantic join/merge/history/transfer; complete style/data/range validation and security lifecycle; update stale XML claims and obtain native chart resaves. |
| ODG | Complete page/group/shape geometry, paths, styles, resources/media, forms, templates, and dependency-aware structural CRUD; add durable conflicts/merge/transfer and security lifecycle; broaden real malformed/producer corpora and obtain native Draw edit/resave evidence. |
| ODI | Add genuine producer ODI/FODI artifacts; broaden frame/style/metadata/image-map/resource semantics and unify flat/package durable operations; use byte-budgeted common history plus merge/transfer; add encryption/signatures, hostile boundaries, and native producer resaves; update stale XML claims. |
| ODM | Prove original genuine `.odm` package ingress/editing rather than transformed fixture XML; extend the strong durable transaction from title/links to section trees, resources, styles, metadata, and dependency closure; add three-way planning/transfer, security lifecycle, and native master-document resaves. |
| OTH | Add structural CRUD for headings/lists/bookmarks/fields/formatting/styles/resources/forms/objects and metadata; serialize durable patches and add merge/transfer; support explicit security lifecycle; broaden real templates and native Writer/Web resaves. |
| ODB | Route source-splice publication through shared provenance enforcement; add durable join/three-way merge/history/transfer; complete remaining producer extensions and dependency dispositions; expose signature/encryption policy; broaden real databases and native Base resaves without executing content. |
| Markdown | Expand the reader from top-level block ranges to a lossless block/inline AST and reference graph; support multi-operation structural edits, joins, durable patches, merge/history, and checked dependency updates; complete all advertised source projections; add CommonMark/GFM conformance corpora and parser round trips. |

## Commands and results

Independent decisive gates run for this review:

```text
cargo check -p <all 19 format crates> --all-targets --all-features
# exit 0

cargo clippy -p <all 19 formats and reviewed common crates> \
  --lib --all-features --no-deps -- -D warnings
# exit 0

cargo test -p xml-minifier --test audit --test ooxml_assets
# audit: 14 passed; assets: 4 passed, 1 ignored
cargo test -p litchi-opc pkgwriter
# 9 passed
cargo test -p litchi-odf-common core::writer
# 22 passed
cargo test -p litchi-odf-common --test odf_corpus
# 2 passed

cargo test -p litchi-markdown --test reader
# 10 passed
cargo test -p litchi-odc --test capability_crud --test libreoffice_corpus
# 6 passed
cargo test -p litchi-odg --test capability_transactions
# 2 passed
cargo test -p litchi-odi --test capability_semantics
# 4 passed
cargo test -p litchi-odm --test capability_transactions
# 7 passed
cargo test -p litchi-oth --test semantic_api
# 19 passed
cargo test -p litchi-odb --test unified_transactions
# 4 passed
cargo test -p litchi-odp --test odp_unified_transaction
# 3 passed
cargo test -p litchi-ppt slide_order::tests
# 3 passed
```

## Bottom line

This remediation wave is materially stronger than the previous review: all production library
Clippy gates are green, authored XML publication is centrally strict with explicit exact-source
exemptions, PPT has a strong durable slide-order owner, Markdown now reads and edits exact source,
and the small ODF formats gained substantial real transactions and fixture evidence.

It still does not support a 95 score. The next wave should prioritize ordinary-root transaction
unification, format-wide durable composition/merge/history/transfer, narrowing correctness-relevant
lint quarantines, honest producer fixtures, and current native edit/resave evidence rather than
adding more isolated feature owners.
