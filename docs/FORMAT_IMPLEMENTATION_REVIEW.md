# Format implementation review

Date: 2026-08-10

Revision under review: committed tree `5baeba5e73d2145989f2aedca2f141049e1156cf`
(`feat: complete non-iwork interoperability roots`).

This is an independent review of that exact commit. The tree was clean before this
document-only update. Earlier score tables, test counts, and conclusions are superseded.

## Scope and rubric

This review covers exactly 19 formats: DOCX, XLSX, PPTX, DOC, XLS, XLSB, PPT, RTF, ODT, ODS,
ODP, OpenDocument Formula, ODC, ODG, ODI, ODM, OTH, ODB, and Markdown. All iWork formats and
shared IWA code are excluded.

The audit compared public APIs and implementations against
`docs/CRUD_Scenario_Checklist.md`, ADRs 0001, 0003-0006, 0008, and 0023-0027, format feature
matrices, checked-in specification material, lint policy, fixture provenance, and executable
evidence. Scores use the repository's existing 100-point judgment rubric:

- **Functional/Completeness** measures public read/create/opened-file edit breadth, lossless or
  explicit-refusal behavior, validation, genuine producer coverage, native changed-file evidence,
  full reopen/readback, and practical production readiness.
- **API/Quality** measures immutable snapshots, checked selectors and values, failure-atomic
  commits, reversible and durable patches, stale checks, deterministic joins and three-way
  planning, bounded history, dependency-aware transfer, bounded I/O, security disposition, XML
  publication integrity, documentation accuracy, and release-gate hygiene.

`95` requires release-grade evidence across the ordinary format root. Strong common machinery or
an excellent narrow owner is insufficient by itself. Completion requires both scores to be at
least 95.

## Authoritative scores

| Format | Functional/Completeness | API/Quality | Decisive finding |
|---|---:|---:|---|
| DOCX | 96 | 98 | Atomic canonical hyperlink batches now span multiple body paragraphs or one bounded block-control/table-cell owner, while retaining the complete relationship graph. The checked LibreOffice 26.2.5 changed-save/readback chain and current generator replay pass; crossing ranges, cross-story owners, arbitrary descendants, and Microsoft Word resave remain outside the evidence. |
| XLSX | 96 | 94 | Validated page-margin CRUD, wider bounded classic-chart relationship closure, and an exact-reproducible LibreOffice changed-save/readback close the prior functional evidence gap. API quality remains below 95 because crate-wide cast, ignored-error, `unwrap`/`expect`, and wildcard correctness allowances still override strict Clippy. |
| PPTX | 96 | 97 | Transfer now classifies table, classic-chart, DiagramML, and inert OLE graphic frames and rewrites all relationship-namespace attributes, with typed refusal for unclassified payloads. A current LibreOffice changed-save/readback is checked and exact-reproducible; Microsoft PowerPoint resave is not claimed. |
| DOC | 94 | 93 | Picture transfer now tolerates unrelated bounded shapes, groups, text boxes, reordered slots, and shared BStore resources while proving the selected singleton graph. Selected nested/group-owned or noncanonical graphs remain refused, no current native changed-save exists, and the broad DOC correctness-lint quarantine remains. |
| XLS | 94 | 93 | Existing canonical `ExtSST` buckets can be updated and reversible row/column shifts cover reference-free formulas, selections, merged ranges, external hyperlinks, and simple drawing anchors. Reference-bearing/shared formula owners and harder drawing dependencies remain refused, no current Excel/LibreOffice resave exists, and broad crate-wide correctness allowances remain. |
| XLSB | 94 | 94 | The durable root now transfers ordinary chart anchors, shapes, nested groups, and connector closure into an existing drawing with collision-safe identity/resource remapping. Active OLE, unknown/MCE graphs, package-global chart dependencies, and mixed conformance remain refused; LibreOffice is import-only for XLSB, no current Excel resave exists, and the largest legacy correctness-lint quarantine remains. |
| PPT | 94 | 96 | Slide visibility and closed `OfficeArtFConnectorRule` transfer are durable and identity-remapped with distinct typed refusals for external shape references. The attempted LibreOffice resave did not preserve the tested order/visibility/anchor semantics, so no native artifact is claimed; broader animation, BLIP, and shape-reference graphs remain unsupported. |
| RTF | 95 | 96 | Root shape-text editing/transfer now joins comments and note stories on the durable ordinary root, and a LibreOffice changed-save/readback preserves the sentinel. The current generator produces an equivalent but not byte-identical pre-native RTF because surrounding character controls differ, so the score stays at the threshold rather than receiving full provenance credit. |
| ODT | 95 | 95 | Advanced layout/protection durability, narrower correctness lint exceptions, manifest-version-safe writing, and a current Writer changed-save/readback meet the threshold. The current generator changes only the regenerated manifest versus the checked pre-native package, remaining layout/security breadth is bounded, and one broad legacy XML-position cast allowance remains. |
| ODS | 95 | 96 | Automatic styles now cover common data-style families, controls and geometry are wider, terminal password encryption/reopen is implemented, and current Calc changed-save/readback succeeds. Encrypted-source transactional re-encryption, signing, structured grid controls, full style/geometry breadth, and byte-exact replay of the checked pre-native manifest remain absent. |
| ODP | 95 | 96 | Source-backed stories, cells, forms, and extension owners now have granular namespace-aware edits, stable crypto refusals, and a successful current Impress changed-save/readback. Crypto authoring and broader producer extensions remain absent; current generator replay differs only in the regenerated pre-native manifest. |
| OpenDocument Formula | 95 | 97 | A crate-local libFuzzer harness plus reproducible nine-seed replayer complements the independent schema/property corpus, and current Math changed-save/readback preserves both MathML and StarMath semantics. `cargo-fuzz` was unavailable and not run, the checked pre-native manifest is not byte-identical to current regeneration, and the feature matrix's “native resave unavailable” sentence is stale. |
| ODC | 95 | 96 | Exact-span edits now cover plot label source, axis categories/grids, legend expansion, and wall/floor styles. A genuine ODFDOM 0.13.0 standalone ODC was independently created, changed, saved, semantically reopened, and ODF-validated; current LibreOffice `chart8` still cannot provide changed-save evidence. The crate feature matrix and older producer-evidence document incorrectly still say no standalone producer. |
| ODG | 96 | 98 | Durable source-backed geometry now covers endpoints, points/view boxes, and transforms; existing-package crypto dispositions are explicit. The current Draw changed-save/readback and current generator replay both succeed, alongside fresh encryption/signing evidence. Existing encrypted rewrite/rekey/re-sign remains final-scoped unsupported. |
| ODI | 95 | 97 | Transfer now closes named/automatic style parent/next/linked dependencies; password opening, signature verification, protected-member dispositions, and exact forms/extensions inventories are public. Genuine ODFDOM create/change/save evidence is strong enough for this niche root, but LibreOffice has no ODI filter and direct form/extension mutation plus non-style resource closure remain absent. |
| ODM | 93 | 97 | Common master lists, tables, and generated-index children receive schema checks, and direct body-item removal is durable and conflict-aware. The builder and broader mixed-content schema remain partial, crypto authoring is absent, and no current Writer native changed-save was attempted. |
| OTH | 94 | 97 | The builder now authors rich text, forms, metadata/styles, embedded files, and object sets; nested-list paragraph transfer and machine-readable security/validation refusals are added. Full Relax NG validation, encryption/signature lifecycle, remaining unprojected inline edits, and current Writer/Web native resave are absent. |
| ODB | 93 | 98 | Bounded query parameter/join inventory, richer constraint/index/relation models, and transitive linked-component payload closure materially deepen the inert database root. It remains a catalog rather than a database runtime, broader SQL/schema/component semantics are unmodeled, the LibreOffice filter is import-only through the CLI route, and no current Base/UNO changed-save exists. |
| Markdown | 98 | 98 | All 652 CommonMark 0.31.2 and 670 pinned GFM examples, exact-range reversible edits, release hashes, formatting, all-target tests/Clippy, and rustdoc pass. |

## Threshold result

Eleven formats meet both required scores: DOCX, PPTX, RTF, ODT, ODS, ODP, OpenDocument Formula,
ODC, ODG, ODI, and Markdown.

The exact sub-95 pairs are:

- XLSX: API/Quality `94`.
- DOC: Functional/Completeness `94`; API/Quality `93`.
- XLS: Functional/Completeness `94`; API/Quality `93`.
- XLSB: Functional/Completeness `94`; API/Quality `94`.
- PPT: Functional/Completeness `94`.
- ODM: Functional/Completeness `93`.
- OTH: Functional/Completeness `94`.
- ODB: Functional/Completeness `93`.

Because eight formats have at least one score below 95, the all-format completion threshold is not
met.

## Cross-cutting evidence

### Build, lint, and documentation

The exact commit passes an all-target/all-feature check for all 19 format crates. Warning-denied
library Clippy passes for all 19 plus `litchi-core`, `litchi-opc`, `litchi-ooxml-common`,
`litchi-odf-common`, and `xml-minifier`. Markdown's complete release gate passes.

Passing Clippy does not neutralize explicit crate-level `allow` attributes. XLSX, DOC, XLS, and
XLSB retain correctness-relevant broad suppressions; XLSB remains the widest. ODT removed
`absurd_extreme_comparisons`, `cast_sign_loss`, `let_underscore_must_use`, and `map_err_ignore`,
but retains broad legacy XML-position truncation allowance. These facts are reflected in the API
scores.

Feature matrices were treated as claims, not proof. Two checked-in claims are stale at this
revision: ODC's matrix and `crates/litchi-odc/docs/PRODUCER_EVIDENCE.md` still deny the standalone
producer now established under `test-data/odf/odc-producer-evidence/`, and the Formula matrix still
says native Math resave is unavailable despite the committed successful chain. This review uses the
artifacts and executable evidence, while recording the documentation inconsistency as quality debt.

### Authored XML minimality and exact-source provenance

The XML publication boundary remains intact. `xml-minifier::audit::verify_authored` rejects
indentation, padded markup, whitespace before tag closes, DTD/custom entities, and ambiguous
all-space nodes outside `xml:space="preserve"`, under finite limits. OPC audits generated content
types/relationships and every authored or changed XML-bearing part; only byte-identical XML from
the opened package receives exact-source exemption.

ODF's source-part/range/splice types still bind retained ranges to one exact archive and part,
reject foreign/stale/overlapping ranges, audit authored fragments, and fully parse the assembled
candidate. The new generic `xml_splice_publication` applies the same single-contiguous-delta proof
to any selected XML part; it is not a bypass. The writer now inherits validated ODF manifest
versions 1.0-1.4, writes the root entry version, omits invalid self/`META-INF` entries, and audits
the regenerated manifest.

Publication evidence passes: XML audit 14; static OOXML assets 4 with one explicit regeneration
test ignored; ODF exact-source splice 5; ODF writer 24; and OPC package writer 9. Direct
`zip::ZipWriter` use in reviewed format crates remains confined to ODS and ODM test directories for
malformed/prettified/producer-ingress negatives. No reviewed format production source bypasses the
shared audited writers.

### Durable API, composition, transfer, and security

Every reviewed root retains source-checked durable replay, exact or semantic inverse, stale
rejection, failure-atomic commit, deterministic join/three-way planning, bounded history, full
candidate reopen, and a bounded transfer disposition for its claimed owners. The commit expands
real dependency closure rather than merely relabeling refusals: DOCX batched relationships; OOXML
graphic frames/charts; legacy picture, cell, drawing, connector, and story graphs; ODF style,
geometry, forms/extensions, schema, query, and component closures.

Security remains explicit and format-specific. ODS supports terminal fresh encryption but not
encrypted transactional re-encryption. ODP and ODM publish closed typed refusals. ODG supports
fresh encryption/signing and exact existing-package dispositions. ODI opens encrypted input for
inert inspection and verifies signature math but refuses protected mutation. OTH exposes a closed
non-execution/refusal contract. ODB inventories protected state and refuses unsupported lifecycle
operations. Explicit refusal improves API quality but is not counted as implemented functionality.

### Producer and native-application evidence

The new LibreOffice 26.2.5.2 evidence follows genuine source -> public Litchi semantic change ->
isolated-profile LibreOffice save -> final Litchi semantic readback for DOCX, XLSX, PPTX, RTF, ODT,
ODS, ODP, Formula, and ODG. All nine checked resaved artifacts have pinned hashes and pass the
current readback executable. Seven provenance/harness tests validate artifact hashes, logs, filter
registry mappings, isolated profiles, and the import-only/no-filter boundaries.

The current generator reproduces the checked DOCX, XLSX, PPTX, and ODG pre-native artifacts
byte-for-byte. RTF remains semantically equivalent but differs in surrounding character controls;
ODT, ODS, ODP, and Formula differ only in regenerated `manifest.xml` after the current manifest
correctness change. Their native outputs and final semantic readbacks remain valid, but this review
does not call those five pre-native artifacts exact-current-generator reproductions.

The evidence is LibreOffice interoperability, not Microsoft Office evidence. DOC and XLS were not
attempted. XLSB and ODB have import-only LibreOffice filters; ODI has no registered filter. PPT was
attempted, but LibreOffice restored or canonicalized tested visibility/order/anchor semantics, so
no successful PPT artifact is retained. ODM and OTH were not attempted.

ODC separately has a genuine standalone ODFDOM 0.13.0 producer chain. Recorded library/source,
artifact, validator, and license hashes match; creation and changed resave ran in separate JVMs;
both packages have the correct stored-first MIME and clean validator transcripts. This is genuine
producer changed-save evidence, not current LibreOffice interoperability.

### Focused executable evidence

All independently selected high-risk suites passed:

| Target | Result |
|---|---:|
| DOCX document transaction | 13 passed |
| XLSX workbook edit/page margins | 59 passed |
| PPTX opened presentation | 13 passed |
| DOC body/resource transaction | 5 passed |
| XLS cell/resource and genuine fixtures | 22 passed |
| XLSB workbook structure/resources | 7 passed |
| PPT slide root | 15 passed |
| RTF dependency transfer/transactions | 27 passed |
| ODT package/layout/real corpus | 32 passed |
| ODS advanced transaction | 12 passed |
| ODP rich/unified transaction | 12 passed |
| OpenDocument Formula capability/schema | 14 passed; 9 fuzz seeds replayed |
| ODC transaction | 11 passed |
| ODG package/capability transaction | 31 passed |
| ODI semantic planning | 14 passed |
| ODM advanced/raw ingress | 14 passed |
| OTH semantic API | 32 passed |
| ODB advanced transaction | 16 passed |
| Markdown release gate | 49 passed, including 1,322 normative examples |
| Native evidence | 9 semantic readbacks; 7 provenance/harness tests |
| XML/package publication | 14 XML, 4 static assets, 5 ODF splice, 24 ODF writer, 9 OPC |

These focused runs do not substitute for Microsoft Office resaves, unexecuted coverage-guided fuzz
campaigns, or the missing native routes named above.

## Concrete remediation for every sub-95 format

| Format | Required remediation to reach both scores |
|---|---|
| XLSX | Localize or remove the crate-wide cast, ignored-error, wildcard, and `unwrap`/`expect` correctness suppressions, with narrow proofs beside schema-generated or bounded codec sites; keep the current drawing/native regression gate green. |
| DOC | Extend selected picture closure to the next common nested/group/text-box/noncanonical cases, eliminate or tightly localize the broad DOC correctness-lint quarantine, and add a current Word or compatible changed-save/readback artifact. |
| XLS | Support the next common reference-bearing/shared-formula and drawing shift closures, localize the broad BIFF correctness suppressions, and add current Excel or compatible changed-save/readback evidence. |
| XLSB | Close the remaining common chart/shape/OLE/MCE/package-global dependencies or final-scope them with corpus evidence, remove the legacy correctness quarantine, and obtain current Excel same-format changed-save/readback evidence. |
| PPT | Add safe closure for the next common BLIP/animation/shape-reference families and demonstrate a native changed-save that preserves the selected semantic operation; the failed LibreOffice result cannot be counted. |
| ODM | Complete common mixed-content/index-child schema and editing breadth, add a richer opened-master builder/security lifecycle or final scope, and prove a changed `.odm` through current Writer save/reopen. |
| OTH | Complete remaining inline/non-top-level editing and full schema-validation scope, add or final-scope crypto lifecycle, and prove a rich resource-bearing `.oth` through current Writer/Web save/reopen. |
| ODB | Deepen typed SQL/schema/component semantics beyond lexical inventory, add or final-scope the protected write lifecycle, and run a resource-bearing changed `.odb` through live Base/UNO `store()` followed by Litchi reopen. |

## Bottom line

The commit materially improves ordinary-root breadth and adds credible current LibreOffice
interoperability for nine formats plus genuine standalone ODC producer evidence. The shared XML
minimality and provenance rules remain enforced, raw direct ZIP negatives remain test-only, and
the selected build, lint, durability, security, publication, and readback tests are green.

The all-format threshold is nevertheless not met. Eight formats retain the explicit sub-95 pairs
listed above; eleven formats meet both required scores.
