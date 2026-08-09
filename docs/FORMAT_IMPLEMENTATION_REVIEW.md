# Format implementation review

Date: 2026-08-10

Revision under review: committed tree `4b0da8d439b107e7f34231d32dee384456d5dfd0`.

This is the authoritative independent review of that commit. The tree was clean before this
document-only update. Earlier score tables, test counts, and XML conclusions are superseded.

## Scope and rubric

This review covers exactly 19 formats: DOCX, XLSX, PPTX, DOC, XLS, XLSB, PPT, RTF, ODT, ODS,
ODP, OpenDocument Formula, ODC, ODG, ODI, ODM, OTH, ODB, and Markdown.

All iWork formats and shared IWA code are excluded. No iWork implementation, test, or score was
used as evidence.

The review compared public APIs and implementations with `docs/CRUD_Scenario_Checklist.md`, ADRs
0001, 0003-0006, 0008, and 0023-0027, format feature matrices, checked-in specifications,
fixtures, and executable tests. Scores use the document's existing 100-point judgment rubric:

- **Functional/Completeness** measures public read/create/opened-file edit breadth, lossless or
  explicit-refusal behavior, validation, genuine producer coverage, full reopen/readback, and
  practical production readiness.
- **API/Quality** measures immutable snapshots, checked selectors and values, failure-atomic
  commits, reversible and durable patches, stale checks, deterministic joins and three-way
  planning, bounded history, dependency-aware transfer, bounded I/O, and release-gate hygiene.

`95` requires release-grade evidence across the ordinary format root. An excellent narrow owner,
large reader inventory, or generic common-layer capability is insufficient by itself. Every format
is below 95 in at least one dimension, so the user's required threshold is not met.

## Authoritative scores

| Format | Functional/Completeness | API/Quality | Decisive reason it remains below 95 |
|---|---:|---:|---|
| DOCX | 93 | 91 | Broad WordprocessingML coverage, but the durable root edits only focused main-body paragraph, hyperlink, and simple-cell text plus insertion; no root three-way plan or dependency transfer |
| XLSX | 91 | 94 | Ordinary workbook durable replay, history, three-way planning, and page-break transfer are strong; transfer is not format-wide, several facade gaps remain, and broad correctness lint allowances remain |
| PPTX | 89 | 88 | The new opened root composes slide order and existing shape/notes text durably, but does not own general slide/shape/resource CRUD, three-way planning, or transfer |
| DOC | 84 | 84 | Real length-changing multi-paragraph body edits now rebuild and fully reopen, but formatting/structure coverage is narrow, no three-way/transfer exists, and legacy correctness lints are quarantined |
| XLS | 82 | 86 | Existing fixed-width cell edits are durable, composable, reversible, and fully reopened; absent-cell structural CRUD, rich/string formula cases, three-way planning, and transfer remain absent |
| XLSB | 87 | 92 | Structural scalar/formula-result cell CRUD and dependency validation are substantial; rich/shared-string and style-resource creation plus broader format-root transfer remain missing, with the largest lint quarantine |
| PPT | 86 | 93 | Slide order plus focused shape text now has durable replay, join, three-way planning, history, and bounded transfer; drawings and most presentation semantics remain outside that root |
| RTF | 92 | 93 | Strong immutable multi-operation editing, durable replay, three-way planning, history, and producer corpus preservation; transaction coverage still omits much of the retained RTF semantic model |
| ODT | 89 | 92 | The ordinary package transaction now has semantic durable text operations, join, three-way planning, and history; many legacy verbs are memory-only, transfer is incomplete, and lint quarantine remains broad |
| ODS | 85 | 93 | A unified durable package transaction spans many owners and resources with join/three-way/history; rich text, style graph, conditional formatting, drawings/forms, formula evaluation, and security lifecycle remain incomplete |
| ODP | 84 | 92 | Unified durable edits cover slides, shapes, media, charts, layouts, masters, annotations, and RDF with history and merge planning; rich text/tables/forms and dependency-aware cross-document transfer are missing |
| ODF Formula | 89 | 93 | Checked MathML, semantic operations, joins, transfer plans, three-way planning, history, and durable envelopes are strong; MathML breadth and StarMath semantics remain bounded rather than complete |
| ODC | 84 | 93 | Canonical package charts have granular durable edits, three-way joins, transfer, and history; arbitrary opened chart XML is read-only/lossless-refuse and no genuine standalone producer ODC/FODC exists |
| ODG | 78 | 92 | Package shape/layer/geometry/path/resource edits now have durable merge/history; style/form/drawing semantics and cross-document transfer remain narrow |
| ODI | 75 | 92 | Flat/package semantic patches, three-way plans, transfer, resource CRUD, and byte-budgeted history are strong; semantic breadth is small and the only checked-in FODI is explicitly synthetic |
| ODM | 80 | 94 | Original LibreOffice ODM ingress and a strong durable section/style/resource/metadata root are now verified; broader master-document semantics and complete dependency transfer/security lifecycle remain incomplete |
| OTH | 74 | 89 | Genuine LibreOffice template ingress plus durable text/list/whole-part edits and merge/history are real; structural, resource, form/object mutation and transfer are mostly absent |
| ODB | 76 | 93 | The unified inert database transaction now uses the shared provenance splice boundary and supports durable join/three-way/history/transfer; database schema breadth and security/native-resave evidence remain limited |
| Markdown | 84 | 93 | Exact-source block/inline views, a reference graph, multi-edit durable patches, merge planning, and history are real; the AST/edit surface is still top-level-block oriented and the corpus is not the normative CommonMark/GFM suite |

## Cross-cutting findings

### Build and lint evidence

The all-target/all-feature check for all 19 format crates passes. Warning-denied production-library
Clippy also passes for all 19 formats plus `litchi-core`, `litchi-opc`, `litchi-ooxml-common`,
`litchi-odf-common`, and `xml-minifier`:

```text
cargo check -p <all 19 format crates> --all-targets --all-features
# exit 0

cargo clippy -p <all 19 formats and reviewed common crates> \
  --lib --all-features --no-deps -- -D warnings
# exit 0
```

That is a genuine strict production gate, but it does not mean the lint debt is removed. Explicit
`#![allow(...)]` attributes take precedence over `-D warnings`:

- DOCX, ODP, OpenDocument Formula, ODC, OTH, and Markdown have no material crate-wide correctness
  quarantine. PPT, RTF, ODG, ODI, ODM, and ODB use narrow or mostly organizational exceptions.
- XLSX, DOC, XLS, XLSB, ODT, and ODS retain broad crate-wide allowances. These include combinations
  of narrowing/sign casts, `unwrap`/`expect`, ignored `Result`s, wildcard handling, and parser-state
  assumptions. XLSB is the broadest; DOC and ODT also suppress correctness-relevant families.
- PPTX's allowances are broad but mostly schema/API/style-oriented; shared `litchi-ooxml-common`
  retains broad parser and style quarantines inherited by every OOXML host.

For 95, each correctness-relevant allowance needs direct remediation or a narrowly located proof
next to generated/spec-shaped code. A reproducible all-target Clippy and rustdoc pass from the CRUD
release checklist is also still required; this review ran strict Clippy for production libraries.

### Authored and referenced XML publication

The previous global XML blocker is closed in the reviewed code.

`xml-minifier::audit::verify_authored` rejects indentation, padded markup, whitespace before tag
closes, DTD/custom entities, and ambiguous all-space character nodes outside
`xml:space="preserve"`, under finite byte/event/attribute/token/text/depth limits. Read-side
`verify` remains permissive enough to preserve genuine source whitespace.

OPC publication classifies XML by conventional XML/RDF/relationship paths and XML media types,
audits generated content-types and relationships, and audits every authored or changed XML-bearing
part. Only a byte-identical XML part captured from the opened package has exact-source exemption.

ODF publication applies the same authored-versus-exact distinction. The new
`XmlSourcePart`/`XmlSourceRange`/`AuthoredXmlFragment`/`XmlSplicePublication` API binds ranges to one
exact source archive and part, audits each authored markup/start-tag/text fragment, refuses stale or
overlapping ranges, fully reparses the assembled XML, and publishes through the shared ODF writer.
ODB now uses this boundary; no reviewed format crate directly constructs production ZIP output.

Raw malformed/prettified packages are constructed with `zip::ZipWriter` only in tests. The raw
negative fixtures cover `.rdf`, manifest-declared XML, `+xml`, signature XML, arbitrary noncompact
fragments, foreign provenance, stale ranges, overlaps, and bounded rebuilds. Focused publication
tests pass: ODF splice 5, XML audit 14, static OOXML assets 4 with 1 explicit regeneration test
ignored, and OPC package-writer 9.

Two feature-matrix statements are stale rather than code defects: ODI and ODM still state that
space-only authored inter-element text can pass. The final shared authored writer rejects it.

### Durable patches, composition, history, and transfer

Commit `4b0da8d4` materially expands ADR-0003 adoption:

- XLSX, PPT, RTF, ODT, ODS, ODP, ODC, ODG, ODI, ODM, ODB, and Markdown now demonstrate substantial
  durable or exact-artifact replay plus deterministic join/three-way/history subsets at an ordinary
  format root.
- DOCX, DOC, XLS, XLSB, and PPTX now have durable focused roots with stale checks and exact inverse;
  their composition/planning/transfer breadth varies and does not cover the whole document.
- OpenDocument Formula has granular tree operations, transfer planning, joins, three-way planning,
  bounded history, and durable sidecar evidence.

The remaining issue is not lack of common machinery. Most formats still do not apply durable
semantic operations, independent joins, non-mutating three-way planning, dependency-closure
cross-document transfer, and commit-coupled history to the complete ordinary editing surface.
Whole-artifact durable envelopes are useful but do not substitute for semantic replay and conflict
resolution after independent evolution.

### Fixture and native-application evidence

The focused tests use genuine artifacts where stated: DOC uses two real multi-generation DOCs;
XLSX uses an Apache POI corpus workbook; PPTX opens a real PowerPoint fixture; XLSB structurally
edits multiple real XLSB fixtures; PPT edits real PPTs; RTF exercises large LibreOffice/Word corpus
files; ODS edits a real Calc workbook; OpenDocument Formula opens two LibreOffice `.odf` files; ODG
opens a LibreOffice `.odg`; ODM now opens the original LibreOffice `.odm` without repacking; OTH
edits a LibreOffice Writer/Web `.oth`; and ODB uses real LibreOffice databases including a signed
fixture.

Important provenance limits remain:

- ODC has real chart subdocuments extracted from FODS/FODT but no standalone producer ODC/FODC.
- ODI has no genuine producer ODI/FODI; its checked-in normative fixture is explicitly synthetic.
- Markdown's conformance cases are original repository fixtures, not the normative CommonMark
  examples or a broad independent GFM compatibility corpus.
- A producer-created input is not the same as current native application interoperability. Most
  new transaction roots still lack documented Litchi-edit -> current Word/Excel/PowerPoint or
  LibreOffice resave -> Litchi-reopen evidence. Exact no-op evidence does not certify changed files.

### Focused executable evidence

All independently selected high-risk suites passed:

| Target | Result |
|---|---:|
| DOCX main-document transaction | 8 passed |
| XLSX durable workbook | 5 passed |
| PPTX opened presentation | 5 passed |
| DOC body transaction | 2 passed |
| XLS cell-value transaction | 9 passed |
| XLSB cell-value edit | 5 passed |
| PPT slide-order root | 6 passed |
| RTF transaction | 16 passed |
| ODT packaged transaction | 14 passed |
| ODS document transaction | 6 passed |
| ODP unified transaction | 5 passed |
| OpenDocument Formula transaction/limits | 8 passed |
| ODC next-wave transaction | 5 passed |
| ODG package semantics | 14 passed |
| ODI semantic planning | 7 passed |
| ODM advanced transaction | 4 passed |
| OTH semantic API | 22 passed |
| ODB advanced transaction | 7 passed |
| Markdown reader and conformance | 20 passed |

These focused runs verify the scored high-risk seams; they do not replace the final full-workspace,
all-feature, rustdoc, fuzz, and native-application release matrix.

## Smallest actionable remediation to reach the 95 threshold

| Format | Smallest defensible next remediation wave |
|---|---|
| DOCX | Extend the existing durable document root across complex runs, nested/rich tables, fields/revisions/controls and dependent relationships; add non-mutating three-way planning and dependency-aware transfer; prove changed output through current Word resave/reopen. |
| XLSX | Generalize the current transfer API beyond page breaks to cells/formulas/styles/shared strings/drawings with dependency closure; close remaining facade-write gaps; narrow correctness lint allowances and run current Excel changed-file interoperability. |
| PPTX | Bring slide/shape creation/removal, formatting, tables/charts/media, notes, masters/layouts, comments, and relationship closure into `opened`; add semantic three-way planning, history redo, and cross-deck transfer; narrow shared/host lint debt. |
| DOC | Extend the body transaction from paragraph text/direct bold to general runs, tables, fields, revisions, stories, drawings and resources; add three-way planning/transfer and localize every cast/unwrap proof. |
| XLS | Support absent-cell insertion/removal, length-changing strings, formula/string caches, row-block/index/dimension closure, and style/string resource creation in the durable root; add three-way planning and transfer; narrow legacy lint allowances. |
| XLSB | Add rich-string/shared-string/style-table resource CRUD and make transfer dependency-complete across the ordinary workbook root; shrink the crate-wide cast/error/style quarantine and validate changed output in current Excel. |
| PPT | Extend the strong slide-order/text root to drawings, formatting, tables, charts, media, masters, comments, and relationships; make transfer close drawing/resource dependencies and add current PowerPoint resave evidence. |
| RTF | Bring fields, tables, lists, styles, objects, headers/footers and other retained destinations into the immutable multi-operation root; add transfer semantics and changed-file Word/LibreOffice resave coverage. |
| ODT | Put remaining RDF/form/chart/resource and rich structural verbs on the semantic durable wire, add dependency-aware cross-document transfer, retire the attached mutable path from ordinary use, and remove correctness lint quarantines. |
| ODS | Add rich-cell text, style graph/conditional-format, drawing/form and formula-dependency CRUD to the unified root; make cross-document transfer update references; add explicit signature/encryption lifecycle and Calc resave evidence. |
| ODP | Add rich text, tables, forms, fine chart data and style/resource dependency CRUD to the unified root; implement dependency-aware cross-deck transfer and current Impress resave evidence. |
| ODF Formula | Complete the checked MathML model/constructor corpus, either parse StarMath semantically or keep it explicitly opaque in all operations, and add broader native Math changed-file interoperability and malformed/fuzz coverage. |
| ODC | Add genuine standalone LibreOffice ODC/FODC fixtures, support lossless semantic edits of noncanonical producer charts, and prove style/data/resource transfer plus changed-file native resave. |
| ODG | Expand shape/group/style/form/resource semantics and dependency-aware cross-drawing transfer, add signature/encryption lifecycle, and validate complex changed drawings in LibreOffice Draw. |
| ODI | Add a genuine producer ODI/FODI artifact, broaden frame/style/map/resource semantics, correct the stale XML matrix row, and validate changed output through a native producer. |
| ODM | Expand the durable root from the current section/style/resource/metadata subset to complete master-document structure and reference closure; add full cross-master transfer, security lifecycle, and native resave evidence; correct the stale XML matrix row. |
| OTH | Add structural rich-text/list/bookmark/field/style/resource/form/object CRUD, semantic durable encoding for every operation, and cross-template transfer; add broader producer and native resave coverage. |
| ODB | Extend typed schema/query/component semantics and dependency dispositions, expose signature/encryption policy lifecycle, and prove changed databases through current LibreOffice Base resave without executing database content. |
| Markdown | Adopt the normative CommonMark examples and broad independent GFM corpus; expose lossless nested block/inline structural edits and reference-aware transfer rather than only top-level block replacement. |

## Bottom line

The commit is materially stronger than the prior review. The authored XML publication contract is
now correctly enforced with controlled exact-source provenance, including the formerly exceptional
ODB splice path. All selected builds, strict production Clippy, transaction suites, raw negatives,
and full-reopen checks pass.

The threshold still is not met. The highest scores are DOCX at 93 Functional/Completeness and XLSX
and ODM at 94 API/Quality. The shortest path to 95 is to finish ordinary-root semantic breadth and
dependency-aware transfer, remove correctness-relevant lint quarantines, and obtain genuine
producer plus current native changed-file interoperability evidence—not to add more isolated
feature owners.
