# Format implementation review

Date: 2026-08-10

Revision under review: committed tree `f5fd760d5eaed150b3222fa6b71419b888289c03`.

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
| DOCX | 94 | 96 | The ordinary document root now has rich-owner durable edits, joins, three-way planning, history, and relationship-checked transfer; complex fields, nested controls/tables, structural run content, and complete relationship-subgraph transfer remain bounded |
| XLSX | 93 | 94 | Cell-region transfer closes formulas, shared strings, and exact styles, with page-break transfer and full merge/history support; rich-run transfer, formula evaluation, facade-write gaps, and broad correctness lint allowances remain |
| PPTX | 93 | 95 | One opened transaction now spans slide order/removal, text, tables, charts, media, layouts/masters, comments, and relationship closure; fully general shape/part creation, modern-comment/extension breadth, and current PowerPoint changed-file evidence remain incomplete |
| DOC | 89 | 91 | Story, field-result, table-cell, revision, and formatting edits now share a durable root with merge/history/limited transfer; OfficeArt/resource closure and broad legacy correctness lint debt remain |
| XLS | 88 | 91 | The durable root now inserts/removes scalar cells, assigns existing XFs, renames sheets, safely shifts rows/columns, rebuilds row indexes, and supports merge/history/transfer; new SST/XF authoring and formula/drawing dependency closure remain incomplete |
| XLSB | 91 | 93 | The workbook root now transfers SST/rich-string-font/style dependency closure and supports rename, three-way planning, history, and exact durable inversion; general resource authoring and the workspace's broadest correctness lint quarantine remain |
| PPT | 89 | 94 | The durable root adds shape/table geometry, slide removal, merge/history, and simple dependency transfer; drawing/comment/external relationship transfer is refused and most presentation semantics remain outside the root |
| RTF | 93 | 95 | Retained body, paragraph, table-cell, header/footer, and formatting destinations now share durable multi-operation editing, merge/history, and dependency-free transfer; fields, lists, styles, objects, and richer structural transfer remain outside the ordinary root |
| ODT | 92 | 94 | Styles, fields, revisions, RDF, protection, forms/resources, and script blobs now participate in semantic durable replay, merge/history, transfer checks, and producer full reopen; structural breadth and broad correctness lint quarantine remain |
| ODS | 91 | 94 | One provenance-splice root now spans sheets/cells, structure, styles, conditional formats, sparklines, drawings/forms, definitions, annotations, RDF, protection, pivots, changes, charts, and resources; deep style/form/chart semantics, security lifecycle, and broad lint debt remain |
| ODP | 88 | 94 | The unified root has durable presentation owners plus typed chart transfer and genuine Impress full reopen; rich text, tables/forms, fine chart data, and general cross-deck dependency closure remain incomplete |
| ODF Formula | 92 | 96 | Checked MathML arity/value structure, granular durable edits, transfer, merge/history, bounds, and changed LibreOffice-package reopen are strong; MathML construction breadth, semantic StarMath, and current native Math resave evidence remain incomplete |
| ODC | 89 | 96 | Canonical packages have typed chart/style/data/resource transfer plus durable merge/history, and noncanonical XML supports checked exact-span edits; there is still no genuine standalone producer ODC/FODC or native changed-file evidence |
| ODG | 87 | 95 | Durable drawing transactions now cover groups, layers, geometry/path, styles, forms, resources, and dependency-checked group transfer on genuine Draw files; arbitrary style/form breadth, collision rewriting, security writes, and native resave remain incomplete |
| ODI | 82 | 94 | Frame/image/style/map/resource semantics now share durable transfer/merge/history with compact publication and hostile-input refusal; no genuine producer ODI/FODI exists, the normative fixture is synthetic, and one lint exception remains |
| ODM | 86 | 95 | Genuine raw LibreOffice ingress plus durable tree/style/resource/metadata CRUD, dependency-closed linked-section transfer, merge/history, and exact inverse are verified; master-document breadth, security writes, and native resave remain incomplete |
| OTH | 80 | 93 | Genuine template ingress now has durable heading/paragraph/list/style/metadata edits, merge/history, optional-part deletion, and style-parent transfer; rich inline/nested lists, resources/forms/objects, and complete transfer remain absent |
| ODB | 84 | 95 | The inert root now has durable dependency-closed schema/component/resource transfer, merge/history, genuine changed Base reopen, encrypted open, and signature policy; linked payload copying, broader typed database semantics, re-sign/re-encrypt, and native resave remain incomplete |
| Markdown | 91 | 96 | Exact ranged nested block/inline edits, reference-graph preflight, dependency-aware transfer, durable merge/history, selected upstream CommonMark/GFM examples, and real-document round trips are verified; corpus coverage is selected rather than the complete normative suites |

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

### Durable patches, composition, history, and transfer

Commit `f5fd760d` materially extends ADR-0003 adoption. Every reviewed format now has source-checked
durable replay, exact or semantic inverse, deterministic joins, non-mutating three-way planning,
bounded commit-coupled history, and at least an explicit bounded transfer disposition at its claimed
ordinary root. The decisive additions include:

- DOCX rich owners and PPTX's opened presentation now compose ordinary edits with durable merge,
  history, and relationship-aware transfer rather than leaving those capabilities in isolated
  owners.
- XLS/XLSB now perform structural workbook changes and dependency-aware cell transfer. XLSX cell
  transfer closes relative-formula, shared-string, and exact-style dependencies; unsupported
  range-owned formulas and unavailable receiver resources fail atomically.
- ODT's residual semantic families are now on the durable wire; ODS spans its advanced owners in
  one provenance-spliced root; ODP transfers typed chart data; ODG/ODM/ODB close substantial
  package dependencies; Markdown transfers nested content with reference-definition preflight.

The remaining problem is breadth, not missing common machinery. A transfer implementation that
correctly refuses drawings, rich text, nested lists, style collisions, or source-local dependencies
is safe and API-valuable, but it does not make those format features complete. Likewise, exact
artifact envelopes do not replace semantic replay after independent evolution. The scores therefore
credit failure atomicity and explicit refusal while keeping Functional/Completeness below 95 where
ordinary documents still contain common unsupported owners or dependencies.

### Fixture and native-application evidence

The focused tests use genuine artifacts where stated: DOCX uses real Office-produced packages and
an Open XML SDK fixture; XLSX uses Apache POI and LibreOffice workbooks; PPTX uses real PowerPoint
decks including a complex chart deck; DOC uses multi-generation Word binaries; XLSB transfers cells
and dependency closure across producer and third-party fixtures; PPT edits real binaries; RTF uses
LibreOffice/Word corpus files; ODT, ODS, ODP, OpenDocument Formula, ODG, ODM, OTH, and ODB exercise
genuine LibreOffice-family packages. ODB includes signed and encrypted lifecycle fixtures.

Important provenance limits remain:

- ODC has real chart subdocuments extracted from FODS/FODT but no standalone producer ODC/FODC.
- ODI has no genuine producer ODI/FODI; its checked-in normative fixture is explicitly synthetic.
- Markdown now checks selected vendored upstream CommonMark 0.31.2/GFM examples and real-document
  fixtures, but not the complete normative CommonMark suite or a broad independent GFM corpus.
- A producer-created input is not the same as current native application interoperability. Most
  new transaction roots still lack documented Litchi-edit -> current Word/Excel/PowerPoint or
  LibreOffice resave -> Litchi-reopen evidence. Exact no-op evidence does not certify changed files.

### Focused executable evidence

All independently selected high-risk suites passed:

| Target | Result |
|---|---:|
| DOCX document transaction | 10 passed |
| XLSX durable workbook | 9 passed |
| PPTX opened presentation | 9 passed |
| DOC body transaction | 3 passed |
| XLS cell-value transaction | 12 passed |
| XLSB workbook structure | 3 passed |
| PPT slide-order root | 8 passed |
| RTF transaction | 20 passed |
| ODT packaged transaction | 18 passed |
| ODS advanced document transaction | 4 passed |
| ODP unified transaction | 6 passed |
| OpenDocument Formula capabilities and limits | 18 passed |
| ODC next-wave transaction | 6 passed |
| ODG package/capability transactions | 23 passed |
| ODI semantic planning | 8 passed |
| ODM advanced transaction and raw ingress | 9 passed |
| OTH semantic API | 24 passed |
| ODB advanced transaction | 13 passed |
| Markdown reader and selected normative corpus | 22 passed |
| XML authored audit | 14 passed |
| Static OOXML asset audit | 4 passed, 1 explicit regeneration test ignored |
| ODF exact-source splice provenance | 5 passed |
| OPC package writer publication | 9 passed |

These focused runs verify the scored high-risk seams; they do not replace the final full-workspace,
all-feature, rustdoc, fuzz, and native-application release matrix. A broad `--tests --all-features`
attempt was also started, but linking all integration targets exhausted the review environment's
disk before execution. It reported no assertion failure and is not counted as a passing gate; the
focused targets above were then built and run individually.

## Smallest actionable remediation to reach the 95 threshold

| Format | Smallest defensible next remediation wave |
|---|---|
| DOCX | Extend the strong root across complex field sequences, block/nested controls, nested/rich tables, structural run content, and complete relationship subgraphs; prove a changed document through current Word resave/reopen. |
| XLSX | Add rich-run and drawing dependency transfer, complete defined-name/page-setup/property facade writes, narrow correctness lint allowances, and prove changed output in current Excel. |
| PPTX | Complete general shape/part creation and modern comment/extension handling, narrow remaining host/shared lint debt, and prove a complex changed deck through current PowerPoint resave/reopen. |
| DOC | Close OfficeArt/drawing/resource dependencies across story/table/field edits and transfer, then localize or remove the broad cast/unwrap/error lint quarantine and validate current Word output. |
| XLS | Author new SST/XF resources, support string formula caches and safe packed deletion, close formula/range/drawing dependencies, narrow the lint quarantine, and validate current Excel output. |
| XLSB | Generalize shared/rich-string and style resource authoring beyond transfer, close formula/drawing dependencies, remove the broad correctness lint quarantine, and validate current Excel output. |
| PPT | Extend the root to drawing, formatting, chart/media, master, comment, and external-relationship owners; close those transfer dependencies and add current PowerPoint resave evidence. |
| RTF | Bring fields, nested tables/lists, styles, objects, and their dependencies into the immutable root and transfer plan; add changed-file Word and LibreOffice resave coverage. |
| ODT | Complete rich structural and chart/resource transfer semantics, remove broad correctness lint allowances, and add current Writer changed-file resave/reopen evidence. |
| ODS | Complete deep style graphs, rich cell runs, form/drawing/chart dependencies, and signature/encryption write policy; remove broad lint allowances and add current Calc resave evidence. |
| ODP | Add rich text, tables/forms, fine chart data, and general style/resource dependency transfer; add current Impress changed-file resave evidence. |
| ODF Formula | Complete the checked MathML constructor/model corpus, keep StarMath uniformly opaque or parse it semantically, and add current LibreOffice Math changed-file resave plus broader fuzz evidence. |
| ODC | Obtain a genuine standalone producer ODC/FODC, broaden noncanonical exact-span edits beyond axis name/style, and prove transferred changed output through a native chart producer. |
| ODG | Broaden arbitrary style/form semantics, support safe style/resource collision rewriting in transfer, add security-write policy, and validate complex changed drawings in current LibreOffice Draw. |
| ODI | Obtain a genuine producer ODI/FODI, broaden frame/map/style semantics, remove the remaining correctness lint exception, and validate changed output through that producer. |
| ODM | Complete master-document structure/reference semantics, security-write policy, and collision-safe cross-master transfer; validate changed output in current LibreOffice Writer. |
| OTH | Add rich inline/nested-list, bookmark/field, resource, form, and object CRUD plus dependency-complete template transfer and current Writer/Web resave evidence. |
| ODB | Copy or explicitly remap linked component payloads, deepen typed schema/query semantics and active-content inventory, add re-sign/re-encrypt support or documented refusal, and prove a changed Base resave. |
| Markdown | Run the complete normative CommonMark suite and a broad independent GFM corpus, close any discovered parser/edit gaps, and publish those results as a reproducible release gate. |

## Bottom line

The commit is materially stronger than the prior review. The authored XML publication contract
remains correctly enforced with controlled exact-source provenance; no reviewed format production
source directly constructs ZIP output, while raw malformed/prettified ZIP construction stays in
tests. The selected builds, strict production Clippy, ordinary-root transactions, transfer/history
tests, raw negatives, and full-reopen checks pass.

The user's threshold still is not met. Several roots now reach 95-96 API/Quality, but no format also
reaches 95 Functional/Completeness; DOCX is highest at 94 Functional/Completeness. The remaining
release gap is common-format semantic and dependency breadth plus genuine current native
changed-file interoperability. Formats with broad correctness lint quarantines also cannot receive
95 API/Quality until those assumptions are locally proved or removed.
