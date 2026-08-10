# Format implementation review

Date: 2026-08-10

Revision under review: committed tree `2ca898bfc1df2279b984fd5a87b615eed6c5fbcc`.

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
| DOCX | 94 | 97 | The ordinary root now covers complex-field results, nested controls/tables, structural run content, and complete bounded relationship-subgraph transfer; nested hyperlinks/crossing selectors and current Word changed-file resave evidence remain absent |
| XLSX | 94 | 94 | The root adds rich shared-string ownership, selected picture-graph transfer, defined-name/page-setup/print/worksheet-property writes, and durable closure; selected rather than general drawing transfer, no recalculation, stale feature-matrix rows, and broad correctness lint allowances remain |
| PPTX | 94 | 96 | Common shape creation/removal, picture and grouped-connector identity/dependency transfer, and modern comments/extensions now share the opened root; unknown shape kinds/external connectors and current PowerPoint changed-file resave remain explicit gaps |
| DOC | 92 | 92 | Managed embedded objects and bounded singleton inline/floating pictures now close field/preview/Data/ObjectPool or drawing dependencies on genuine DOCs; picture changes are not on the durable wire, richer drawing graphs remain refused, and broad lint debt remains |
| XLS | 92 | 92 | Plain SST interning, validated XF authoring, string formula caches, and edge `MulRk` deletion now reopen and invert; rich SST runs, formula compilation/evaluation, interior packed deletion, and cross-range/drawing shifts remain refused under broad lint quarantine |
| XLSB | 94 | 93 | The root now authors SST/rich strings, complete typed styles, formula cache families, images, and image transfer; broader formula/drawing dependency semantics and the workspace's broadest correctness lint quarantine still prevent release-grade quality |
| PPT | 92 | 94 | External-media edits and bounded slide transfer now reuse matching drawing, hyperlink, sound, author, and extension owners; mismatched/native-ID dependencies, active OLE/actions, and much broader presentation mutation remain refused |
| RTF | 94 | 96 | Ordinary transfer now closes passive fields, nested tables, styles, picture-bullet lists, drawings, and inert object/result-picture dependencies; opaque destinations, active links, broader structural editing, and current native changed-file evidence remain gaps |
| ODT | 94 | 89 | Rich notes/ruby, forms, charts, and package resources now use semantic durable replay and transfer; the strict production Clippy gate fails with 334 surfaced `unwrap`/`expect` diagnostics, and advanced layout/security/native-resave coverage remains incomplete |
| ODS | 94 | 96 | Deep automatic-style graphs, rich cells, typed forms/events, grouped text drawings, charts/resources, and explicit stale-signature stripping now share the provenance root; full style/form/geometry breadth, encryption writes, and current Calc resave remain incomplete |
| ODP | 92 | 95 | Rich text boxes, tables, inert controls, fine chart data, and collision-remapped dependency transfer are durable; arbitrary story/list/table/form models, producer extensions, security operations, and current Impress resave remain incomplete |
| ODF Formula | 94 | 97 | The public checked model/constructors now cover every accepted Content MathML symbol and consistent opaque StarMath boundaries with mutation/property corpora; native Math changed-file resave and broader independent/fuzz evidence remain absent |
| ODC | 91 | 97 | Noncanonical exact-span editing now covers chart/plot/series/ODF-1.4 geometry and participates in verified transfer; no genuine standalone producer ODC/FODC or native changed-file evidence exists |
| ODG | 92 | 96 | Nested-group edits, arbitrary automatic styles/forms, collision-remapped style/form/resource transfer, active-content inventory, and signature-removal policy are durable; broader drawing semantics, encryption/signing writes, and current Draw resave remain incomplete |
| ODI | 84 | 95 | Package-origin frame transfer now carries exact resource/style dependencies and the remaining correctness lint exception was removed; no genuine producer ODI/FODI exists, the fixture remains synthetic, and native resave evidence is unavailable |
| ODM | 90 | 96 | Common master structure, active-content policy, complete style/resource closure, collision remapping, and absent-`styles.xml` creation now share the durable root; broader schema validation, signing/encryption writes, and current Writer resave remain incomplete |
| OTH | 87 | 95 | Rich inline blocks, nested-list edits, form-catalog CRUD, durable replay, and full reopen are added; resource/object payload mutation and resource-bearing or rich/nested transfer remain unsupported |
| ODB | 89 | 96 | Local component subtrees now transfer with collision remapping, typed query columns, active-content dispositions, and machine-readable protection capabilities; broader database semantics, re-sign/re-encrypt, and current Base resave remain incomplete |
| Markdown | 97 | 94 | All 652 CommonMark 0.31.2 and 670 pinned GFM examples pass with exact ranges and reversible edits, but the checked-in release script fails its all-target Clippy step on `tests/normative.rs` |

## Cross-cutting findings

### Build and lint evidence

The all-target/all-feature check for all 19 format crates passes. Warning-denied production-library
Clippy passes for 18 formats plus `litchi-core`, `litchi-opc`, `litchi-ooxml-common`,
`litchi-odf-common`, and `xml-minifier`, but fails for ODT:

```text
cargo check -p <all 19 format crates> --all-targets --all-features
# exit 0

cargo clippy -p <18 formats except ODT and reviewed common crates> \
  --lib --all-features --no-deps -- -D warnings
# exit 0

cargo clippy -p litchi-odt --lib --all-features --no-deps -- -D warnings
# exit 101: 334 clippy::unwrap_used / clippy::expect_used diagnostics
```

ODT changed those two lints from crate-wide `allow` to `warn`, which is useful visibility but means
the required warning-denied production gate is concretely red. For the passing crates, explicit
`#![allow(...)]` attributes still take precedence over `-D warnings`:

- DOCX, ODP, OpenDocument Formula, ODC, OTH, and Markdown have no material crate-wide correctness
  quarantine. PPT, RTF, ODS, ODG, ODI, ODM, and ODB use narrow or mostly organizational exceptions;
  ODS removed its crate-wide cast/`expect` exceptions and ODI removed `map_err_ignore`.
- XLSX, DOC, XLS, XLSB, and ODT retain broad crate-wide allowances for other families. These include
  combinations of narrowing/sign casts, ignored `Result`s, wildcard handling, and parser-state
  assumptions. XLSB remains the broadest; DOC and ODT retain correctness-relevant exceptions in
  addition to ODT's newly surfaced panic diagnostics.
- PPTX's allowances are broad but mostly schema/API/style-oriented; shared `litchi-ooxml-common`
  retains broad parser and style quarantines inherited by every OOXML host.

For 95, each correctness-relevant allowance needs direct remediation or a narrowly located proof
next to generated/spec-shaped code. Markdown adds a reproducible release script and its hashes,
format check, complete tests, and rustdoc pass, but the script itself exits 101 because all-target
Clippy reports `clippy::struct_field_names` on the `Example.example` test fixture field. Production
library Clippy is green; the advertised release gate is not.

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
This commit adds an audited end-tag class and a bounded single-contiguous-delta helper: it derives a
checked source range, classifies only the replacement fragment, preserves every unrelated producer
byte, and fully verifies the assembled document. If a candidate cannot be proven as that exact
splice it falls back to strict whole-part authored compactness rather than receiving a source
exemption. ODB and the expanded ODF roots use this boundary; no reviewed format crate directly
constructs production ZIP output.

Raw malformed/prettified packages are constructed with `zip::ZipWriter` only in tests. The raw
negative fixtures cover `.rdf`, manifest-declared XML, `+xml`, signature XML, arbitrary noncompact
fragments, foreign provenance, stale ranges, overlaps, and bounded rebuilds. Focused publication
tests pass: ODF splice 5, XML audit 14, static OOXML assets 4 with 1 explicit regeneration test
ignored, and OPC package-writer 9.

### Durable patches, composition, history, and transfer

Commit `2ca898bfc` closes substantially more common dependencies without weakening ADR-0003. Every
reviewed format retains source-checked durable replay, exact or semantic inverse, deterministic
joins, non-mutating three-way planning, bounded commit-coupled history, and an explicit bounded
transfer disposition at its claimed ordinary root. The decisive additions include:

- DOCX copies complete bounded internal relationship subgraphs and addresses complex-field,
  structural-run, nested-control, and nested-table owners. PPTX creates common shapes, edits modern
  comments/extensions, and remaps group, connector, relationship, and non-visual identities.
- XLSX transfers rich shared strings and selected picture graphs and adds ordinary worksheet/catalog
  writes. XLS authors SST/XF/string-cache resources. XLSB authors rich strings, styles, formulas,
  and images. DOC closes managed embedded and bounded native-picture graphs; PPT and RTF close
  materially broader external-media, drawing, style, list, table, field, and inert-object owners.
- ODT puts rich notes/ruby, forms, charts, and resources on the semantic wire. ODS closes deep
  style/form/drawing/chart dependencies. ODP adds rich text/table/form/fine-chart owners. ODG,
  ODI, ODM, OTH, and ODB expand collision-aware package closure and active/security dispositions.
- OpenDocument Formula exposes complete checked constructors for its accepted Content MathML
  model. Markdown vendors and executes the complete pinned CommonMark and GFM example corpora.

The remaining functional problem is breadth, not missing common machinery. A transfer implementation that
correctly refuses drawings, rich text, nested lists, style collisions, or source-local dependencies
is safe and API-valuable, but it does not make those format features complete. Likewise, exact
artifact envelopes do not replace semantic replay after independent evolution. The scores therefore
credit failure atomicity and explicit refusal while keeping Functional/Completeness below 95 where
ordinary documents still contain common unsupported owners or dependencies. Separately, ODT's
strict production lint failure and Markdown's failing advertised release script are direct quality
failures even though their focused behavioral tests pass.

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
- Markdown's pinned corpora now include all 652 CommonMark 0.31.2 examples and all 670 examples from
  the recorded `cmark-gfm` specification, with checked hashes, licenses, expected HTML, exact ranges,
  deterministic reparsing, and reversible edits. That closes the prior corpus-breadth finding.
- A producer-created input is not the same as current native application interoperability. Most
  new transaction roots still lack documented Litchi-edit -> current Word/Excel/PowerPoint or
  LibreOffice resave -> Litchi-reopen evidence. Exact no-op evidence does not certify changed files.

### Focused executable evidence

All independently selected high-risk suites passed:

| Target | Result |
|---|---:|
| DOCX document/package root | 12 passed |
| XLSX durable workbook and cell dependencies | 21 passed |
| PPTX opened presentation | 13 passed |
| DOC body/resource transaction | 5 passed |
| XLS cell/resource transaction and genuine fixtures | 19 passed |
| XLSB workbook structure/resources | 4 passed |
| PPT slide-order/media root | 10 passed |
| RTF transaction and dependency transfer | 23 passed |
| ODT packaged transaction | 22 passed |
| ODS advanced document transaction | 8 passed |
| ODP unified and rich-content transactions | 9 passed |
| OpenDocument Formula capability/property corpora | 19 passed |
| ODC next-wave transaction | 8 passed |
| ODG package/capability transactions | 28 passed |
| ODI semantic planning | 10 passed |
| ODM advanced transaction | 9 passed |
| OTH semantic API | 27 passed |
| ODB advanced transaction | 15 passed |
| Markdown all targets | 49 passed, including 1,322 normative examples |
| XML authored audit | 14 passed |
| Static OOXML asset audit | 4 passed, 1 explicit regeneration test ignored |
| ODF exact-source splice provenance | 5 passed |
| OPC package writer publication | 9 passed |

These focused runs verify the scored high-risk seams; they do not replace the final full-workspace,
all-feature, fuzz, and native-application release matrix. Markdown's corpus hashes, format check,
all-target tests, and rustdoc pass independently; its all-target Clippy step fails as described
above, so the checked-in release script is not counted as passing.

## Smallest actionable remediation to reach the 95 threshold

| Format | Smallest defensible next remediation wave |
|---|---|
| DOCX | Add checked nested-hyperlink and non-crossing composite selectors where lossless ownership is provable, then run a complex Litchi edit through current Word resave and Litchi reopen. |
| XLSX | Generalize selected-picture transfer to the supported drawing/chart graph, update stale matrix rows for ordinary writes, localize broad cast/unwrap/error assumptions, and run current Excel changed-file interoperability. |
| PPTX | Support or precisely classify remaining common shape kinds and externally attached connectors, narrow host/shared lint exceptions, and prove the changed opened root through current PowerPoint resave/reopen. |
| DOC | Put bounded native-picture changes on the durable wire, extend closure beyond singleton/nonshared graphs, localize broad lint assumptions, and validate the changed binary in current Word. |
| XLS | Support interior packed deletion and dependency-safe formula/range/drawing shifts, close remaining rich-SST/formula authoring gaps, narrow lint allowances, and validate current Excel output. |
| XLSB | Close broader formula/name/table and drawing dependencies, remove the crate-wide correctness quarantine, and validate rich/resource-bearing changed workbooks in current Excel. |
| PPT | Remap rather than require identical safe drawing/hyperlink/sound/comment owners, broaden ordinary-root presentation mutation, and add current PowerPoint changed-file evidence. |
| RTF | Extend typed editing beyond the transferred retained owners into the remaining common destinations while keeping active/opaque refusal, and add current Word plus LibreOffice changed-file resave coverage. |
| ODT | Eliminate the 334 warning-denied `unwrap`/`expect` diagnostics first, localize remaining correctness allowances, then complete advanced layout/security owners and current Writer changed-file resave evidence. |
| ODS | Complete automatic-style replacement/removal and package-wide resolution, broader form/drawing geometry, encryption policy, and current Calc changed-file resave evidence. |
| ODP | Complete arbitrary story/list/table/form read-edit models and producer-specific dependency families, add signature/encryption policy, and run current Impress changed-file resave/reopen. |
| ODF Formula | Add independent schema/fuzz validation beyond deterministic mutation, and prove a changed package through current LibreOffice Math resave/reopen. |
| ODC | Obtain a genuine standalone producer ODC/FODC and current native changed-file round trip; retain exact-span provenance for any newly encountered producer layout. |
| ODG | Complete remaining advanced drawing/style semantics and password/signing lifecycle, then validate collision-remapped complex changed drawings in current LibreOffice Draw. |
| ODI | Obtain a genuine producer ODI/FODI and native changed-file round trip, then add semantic active-content inventory and broader granular style/security lifecycle support. |
| ODM | Broaden master-document schema validation and editing beyond the common structure, add signature/encryption write lifecycle, and validate current Writer resave/reopen. |
| OTH | Add resource/object reference and payload mutation plus dependency-complete rich-inline/nested-list/resource transfer and current Writer/Web resave evidence. |
| ODB | Deepen typed query/schema/component semantics, add supported re-sign/re-encrypt lifecycle or final explicit scope, and prove resource-bearing changed databases through current Base resave. |
| Markdown | Rename or narrowly justify the `Example.example` test field so the checked-in all-target Clippy step passes, then keep the complete corpus gate green in CI. |

## Bottom line

The commit is materially stronger than the prior review. The authored XML publication contract
remains correctly enforced with controlled exact-source provenance; the new contiguous-splice path
retains only proven source bytes and audits every authored delta. No reviewed format production
source directly constructs ZIP output, while raw malformed/prettified ZIP construction stays in
tests. All selected ordinary-root, dependency-transfer, history, raw-negative, and full-reopen tests
pass, as does the 19-format all-target build.

The user's threshold still is not met. Markdown reaches 97 Functional/Completeness but only 94
API/Quality because its advertised release script fails all-target Clippy. Several other roots reach
95-97 API/Quality but remain below 95 Functional/Completeness due to documented common-feature or
producer/native-interoperability gaps. ODT is additionally blocked at 89 API/Quality by 334 strict
production Clippy errors. ODC and ODI still lack genuine standalone producer evidence. Those are
release evidence and quality failures, not score-table formalities, so no reviewed format receives
both required scores.
