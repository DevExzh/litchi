# Format implementation review

Date: 2026-08-10

Revision under review: committed tree `edff2304400da52c310de35bc8621a1a1b931bab`
(`feat: close remaining non-iwork format gaps`).

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
| XLSX | 96 | 96 | Validated page-margin CRUD, wider bounded classic-chart relationship closure, and an exact-reproducible LibreOffice changed-save/readback retain the prior functional result. The crate-wide cast, ignored-error, `unwrap`/`expect`, and wildcard correctness suppressions are now removed; all 720 library tests and strict warning-denied Clippy pass. |
| PPTX | 96 | 97 | Transfer now classifies table, classic-chart, DiagramML, and inert OLE graphic frames and rewrites all relationship-namespace attributes, with typed refusal for unclassified payloads. A current LibreOffice changed-save/readback is checked and exact-reproducible; Microsoft PowerPoint resave is not claimed. |
| DOC | 95 | 95 | Canonical picture transfer now proves the selected graph while tolerating unrelated main/header-story shapes, groups, text boxes, reordered slots, and shared BStore resources. The ordinary body-text root and critical picture-graph boundaries override the legacy lint quarantine with strict correctness denies, and an exact-reproducible LibreOffice 26.2.5 DOC changed-save/readback preserves the paragraph sentinel. Nested/group-owned and other noncanonical selected graphs remain typed refusals, and Microsoft Word resave is not claimed. |
| XLS | 95 | 95 | Reversible row/column shifts now rewrite bounded same-sheet `PtgRef`/`PtgArea` formula tokens while preserving relative/absolute flags, alongside selections, merged ranges, external hyperlinks, and simple drawing anchors. The crate-wide production `unwrap`/`expect` quarantine is replaced by reasoned module-local exceptions, and an exact-reproducible LibreOffice 26.2.5 XLS changed-save/readback preserves `42.25`. Names/3-D/shared-relative tokens and harder drawing owners remain refused; Microsoft Excel resave is not claimed. |
| XLSB | 95 | 94 | The durable transfer root covers every anchor in the checked four-workbook/six-anchor standard-drawing corpus, including shape hyperlinks and chart-part MCE, with compact chart XML and collision-safe graph remapping. Active OLE, other relationship-bearing shapes, drawing-level MCE, mixed conformance, and package-global chart dependencies remain refused; LibreOffice is import-only and no current Excel resave exists. API quality remains below 95 because crate-wide cast, `expect`, ignored-error, and wildcard correctness suppressions still cover the legacy codec outside the strict new transfer module. |
| PPT | 94 | 96 | The durable root now edits slide visibility and canonical fixed-width manual/automatic advance timing while retaining effects, sound, hidden/cursor flags, reserved bits, and framing; genuine producer PPT cases reopen and durable/inverse/history/merge behavior passes. The attempted LibreOffice resave still failed to preserve the tested order/visibility/anchor semantics, so no native artifact is claimed; broader BLIP, animation, and shape-reference closure remains unsupported. |
| RTF | 95 | 96 | Root shape-text editing/transfer now joins comments and note stories on the durable ordinary root, and a LibreOffice changed-save/readback preserves the sentinel. The current generator produces an equivalent but not byte-identical pre-native RTF because surrounding character controls differ, so the score stays at the threshold rather than receiving full provenance credit. |
| ODT | 95 | 95 | Advanced layout/protection durability, narrower correctness lint exceptions, manifest-version-safe writing, and a current Writer changed-save/readback meet the threshold. The current generator changes only the regenerated manifest versus the checked pre-native package, remaining layout/security breadth is bounded, and one broad legacy XML-position cast allowance remains. |
| ODS | 95 | 96 | Automatic styles now cover common data-style families, controls and geometry are wider, terminal password encryption/reopen is implemented, and current Calc changed-save/readback succeeds. Encrypted-source transactional re-encryption, signing, structured grid controls, full style/geometry breadth, and byte-exact replay of the checked pre-native manifest remain absent. |
| ODP | 95 | 96 | Source-backed stories, cells, forms, and extension owners now have granular namespace-aware edits, stable crypto refusals, and a successful current Impress changed-save/readback. Crypto authoring and broader producer extensions remain absent; current generator replay differs only in the regenerated pre-native manifest. |
| OpenDocument Formula | 95 | 97 | A crate-local libFuzzer harness plus reproducible nine-seed replayer complements the independent schema/property corpus, and current Math changed-save/readback preserves both MathML and StarMath semantics. `cargo-fuzz` was unavailable and not run, and the checked pre-native manifest is not byte-identical to current regeneration; the feature matrix now accurately records the successful native chain. |
| ODC | 95 | 97 | Exact-span edits cover plot label source, axis categories/grids, legend expansion, and wall/floor styles. A genuine ODFDOM 0.13.0 standalone ODC was independently created, changed, saved, semantically reopened, and ODF-validated; current LibreOffice `chart8` still cannot provide changed-save evidence. The feature matrix and producer-evidence document now accurately distinguish ODFDOM ODC evidence from unavailable LibreOffice/FODC evidence. |
| ODG | 96 | 98 | Durable source-backed geometry now covers endpoints, points/view boxes, and transforms; existing-package crypto dispositions are explicit. The current Draw changed-save/readback and current generator replay both succeed, alongside fresh encryption/signing evidence. Existing encrypted rewrite/rekey/re-sign remains final-scoped unsupported. |
| ODI | 95 | 97 | Transfer now closes named/automatic style parent/next/linked dependencies; password opening, signature verification, protected-member dispositions, and exact forms/extensions inventories are public. Genuine ODFDOM create/change/save evidence is strong enough for this niche root, but LibreOffice has no ODI filter and direct form/extension mutation plus non-style resource closure remain absent. |
| ODM | 95 | 98 | The fresh builder and opened-master transaction now author all bounded common direct-body kinds: paragraphs, level 1-10 headings, nonempty lists, rectangular named tables, and generated indexes. Dependency-free item transfer carries exact provenance, dependent fragments fail closed, and genuine LibreOffice masters pass compact publication, durable/inverse/history/merge, and full reopen. Inline mixed content remains deliberately open, crypto authoring is final-scoped unavailable, and no current Writer native resave was attempted. |
| OTH | 95 | 98 | Complete inline replacement plus exact-boundary prepend/append now works for nested-list paragraphs while preserving unknown producer inline bytes; the path is durable, reversible, merge/history-aware, transferable, and fully reopened on a resource-enriched genuine Writer/Web template. Full Relax NG validation and destructive replacement of unknown inline markup remain refused, crypto is final-scoped unavailable, and the producer fixture edit is not a native changed-save. |
| ODB | 95 | 98 | Bounded SQL analysis now classifies common `SELECT`/`INSERT`/`UPDATE`/`DELETE` structure and dependency support, query transfer can cascade local table/FK closure, schema ambiguity fails before mutation, and component inventories are exact. A live isolated UNO `XStorable.store()` changed-save/readback preserves the inert `SELECT 424242` query and original table without opening a connection or executing content. The permanent database-runtime boundary and typed refusals for CTEs/subqueries/set operations/table functions/multiple statements remain explicit. |
| Markdown | 98 | 98 | All 652 CommonMark 0.31.2 and 670 pinned GFM examples, exact-range reversible edits, release hashes, formatting, all-target tests/Clippy, and rustdoc pass. |

## Threshold result

Seventeen formats meet both required scores: DOCX, XLSX, PPTX, DOC, XLS, RTF, ODT, ODS, ODP,
OpenDocument Formula, ODC, ODG, ODI, ODM, OTH, ODB, and Markdown.

The exact sub-95 pairs are:

- XLSB: API/Quality `94`.
- PPT: Functional/Completeness `94`.

Because two formats have one score below 95, the all-format completion threshold is not
met.

## Cross-cutting evidence

### Build, lint, and documentation

The exact commit passes an all-target/all-feature check for all 19 format crates. Warning-denied
library Clippy passes for all 19 plus `litchi-core`, `litchi-opc`, `litchi-ooxml-common`,
`litchi-odf-common`, and `xml-minifier`. Markdown's complete release gate passes.

Passing Clippy does not neutralize explicit `allow` attributes. XLSX has removed its crate-wide
correctness suppressions. XLS no longer grants crate-wide production `unwrap`/`expect` exemptions,
although many bounded legacy modules retain individually reasoned exceptions. DOC still has a broad
legacy codec quarantine, but its ordinary body-text root and critical picture-graph functions now
override it with strict correctness denies. XLSB still grants crate-wide cast, `expect`,
ignored-error, and wildcard exceptions outside its strict new drawing-transfer module; that is the
remaining lint-based threshold blocker. ODT's broad legacy XML-position truncation allowance is
unchanged and remains reflected in its threshold API score.

Feature matrices were treated as claims, not proof. The ODC feature matrix and producer-evidence
document now accurately record the genuine standalone ODFDOM chain, and the Formula matrix now
records its current Math native chain. The changed DOC/XLS/ODB behavior and explicit refusal scopes
match their implementations and tests; native evidence is credited only from the separately pinned
artifacts and executable readback, not from matrix wording.

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

Every reviewed ordinary root retains source-checked durable replay, exact or semantic inverse, stale
rejection, failure-atomic commit, deterministic join/three-way planning, bounded history, full
candidate reopen, and a bounded transfer disposition for its claimed owners. The commit expands
real ordinary-root behavior rather than merely relabeling refusals: DOC picture graph coexistence;
XLS same-sheet reference-token shifts; XLSB drawing/chart/hyperlink closure; PPT slide-advance
timing; ODM typed common body authoring/transfer; OTH provenance-safe inline boundary edits; and ODB
typed SQL, dependency, schema, and component closure.

Security remains explicit and format-specific. ODS supports terminal fresh encryption but not
encrypted transactional re-encryption. ODP publishes closed typed refusals. ODG supports fresh
encryption/signing and exact existing-package dispositions. ODI opens encrypted input for inert
inspection and verifies signature math but refuses protected mutation. ODM and OTH expose closed
final-scoped crypto/non-execution refusals. ODB returns typed verification/removal support and
distinct re-signing/re-encryption refusals. Explicit refusal improves API quality but is not counted
as implemented functionality.

### Producer and native-application evidence

The LibreOffice 26.2.5.2 evidence follows genuine source -> public Litchi semantic change ->
isolated-profile LibreOffice save -> final Litchi semantic readback for DOC, DOCX, XLS, XLSX, PPTX,
RTF, ODT, ODS, ODP, Formula, and ODG. ODB uses a fresh isolated profile and loopback-only live UNO
`XStorable.store()` because its CLI filter is import-only. All 12 checked resaved artifacts have
pinned hashes and pass the current readback executable. Seven provenance/harness tests validate
artifact hashes, logs, filter registry mappings, isolated profiles, and unavailable routes.

The current generator reproduces the checked DOC, DOCX, XLS, XLSX, PPTX, ODB, and ODG pre-native
artifacts byte-for-byte. RTF remains semantically equivalent but differs in surrounding character
controls; ODT, ODS, ODP, and Formula differ only in regenerated `manifest.xml` after the manifest
correctness change. Their native outputs and final semantic readbacks remain valid, but this review
does not call those five pre-native artifacts exact-current-generator reproductions.

The evidence is LibreOffice interoperability, not Microsoft Office evidence. XLSB has an
import-only LibreOffice filter and no current Excel resave; ODI has no registered filter. PPT was
attempted, but LibreOffice restored or canonicalized tested visibility/order/anchor semantics, so
no successful PPT artifact is retained. ODM and OTH were not attempted. The ODB helper loads hidden
with macros disabled, requests no connection/query/form/report, and therefore proves same-package
persistence only, not database execution.

ODC separately has a genuine standalone ODFDOM 0.13.0 producer chain. Recorded library/source,
artifact, validator, and license hashes match; creation and changed resave ran in separate JVMs;
both packages have the correct stored-first MIME and clean validator transcripts. This is genuine
producer changed-save evidence, not current LibreOffice interoperability.

### Focused executable evidence

All independently selected high-risk suites passed. Changed roots were rerun at this revision;
source-identical roots retain the previously selected counts after the full all-target check:

| Target | Result |
|---|---:|
| DOCX document transaction | 13 passed |
| XLSX library | 720 passed |
| PPTX opened presentation | 13 passed |
| DOC body/resource transaction | 6 passed |
| XLS cell/resource and genuine fixtures | 24 passed |
| XLSB workbook structure/resources | 9 passed |
| PPT slide root | 17 passed |
| RTF dependency transfer/transactions | 27 passed |
| ODT package/layout/real corpus | 32 passed |
| ODS advanced transaction | 12 passed |
| ODP rich/unified transaction | 12 passed |
| OpenDocument Formula capability/schema | 14 passed; 9 fuzz seeds replayed |
| ODC transaction | 11 passed |
| ODG package/capability transaction | 31 passed |
| ODI semantic planning | 14 passed |
| ODM advanced/raw/semantic/typed authoring | 22 passed |
| OTH semantic API | 33 passed |
| ODB advanced transaction | 19 passed |
| Markdown release gate | 49 passed, including 1,322 normative examples |
| Native evidence | 12 semantic readbacks; 7 provenance/harness tests |
| XML/package publication | 14 XML, 4 static assets, 5 ODF splice, 24 ODF writer, 9 OPC |

These focused runs do not substitute for Microsoft Office resaves, unexecuted coverage-guided fuzz
campaigns, or the missing native routes named above.

## Concrete remediation for every sub-95 format

| Format | Required remediation to reach both scores |
|---|---|
| XLSB | Remove or tightly localize the remaining crate-wide cast, `expect`, ignored-error, and wildcard correctness suppressions, with strict warning-denied coverage across the ordinary workbook root. Current Excel same-format changed-save/readback would strengthen functional evidence but is not the scored API blocker. |
| PPT | Demonstrate a same-format native changed-save that preserves one claimed durable-root semantic operation (order, visibility, anchor, or advance timing), or narrow/final-scope the interoperability claim if compatible applications canonicalize it. The recorded failed LibreOffice result cannot be counted. |

## Bottom line

The commit materially improves the eight former sub-95 roots and adds credible current LibreOffice
interoperability for 12 formats plus genuine standalone ODC producer evidence. The shared XML
minimality and provenance rules remain enforced, raw direct ZIP negatives remain test-only, and
the selected build, lint, durability, security, publication, and readback tests are green.

The all-format threshold is nevertheless not met. XLSB retains API/Quality `94`, and PPT retains
Functional/Completeness `94`; the other 17 formats meet both required scores.
