# Format implementation review

Date: 2026-08-10

Revision under review: committed tree `5e5256af1faf05c7392d0c54ed329510fa045513`
(`feat: complete non-iwork release roots`).

This is an independent review of that exact commit. The tree was clean before this
document-only update. Earlier score tables, test counts, and conclusions are superseded.

## Scope and rubric

This review covers exactly 19 formats: DOCX, XLSX, PPTX, DOC, XLS, XLSB, PPT, RTF, ODT, ODS,
ODP, OpenDocument Formula, ODC, ODG, ODI, ODM, OTH, ODB, and Markdown. All iWork formats and
shared IWA code are excluded.

The audit compared public APIs and implementations against
`docs/CRUD_Scenario_Checklist.md`, ADRs 0001, 0003-0006, 0008, and 0023-0027, each format's
feature matrix, checked-in specification material, fixture provenance, and executable tests.
Scores use the repository's existing 100-point judgment rubric:

- **Functional/Completeness** measures public read/create/opened-file edit breadth, lossless or
  explicit-refusal behavior, validation, genuine producer coverage, full reopen/readback, and
  practical production readiness.
- **API/Quality** measures immutable snapshots, checked selectors and values, failure-atomic
  commits, reversible and durable patches, stale checks, deterministic joins and three-way
  planning, bounded history, dependency-aware transfer, bounded I/O, security disposition, and
  release-gate hygiene.

`95` requires release-grade evidence across the ordinary format root. Strong common machinery or
an excellent narrow owner is insufficient by itself. Scores are deliberately conservative where
the host cannot supply a current native-application changed-file resave.

## Authoritative scores

| Format | Functional/Completeness | API/Quality | Decisive finding |
|---|---:|---:|---|
| DOCX | 94 | 97 | Complete bounded internal relationship-subgraph transfer now includes direct hyperlinks in path-addressed nested controls/tables. Crossing scopes, arbitrary descendants, cross-paragraph selection, and current Word changed-file resave evidence remain outside the claim. |
| XLSX | 94 | 94 | Rich shared strings, full ordinary defined-name replacement/clear, page setup/print options, and selected image plus relationship-free classic-chart leaf transfer are durable. General drawing/ChartEx graphs, recalculation, page-margin writes, current Excel resave, and broad correctness lint allowances remain. |
| PPTX | 94 | 96 | Top-level connector endpoint closure and typed refusals for nested, identity-less, unknown, or unresolved shapes make transfer safer. Common unsupported shape/extension families and current PowerPoint resave still cap completeness. |
| DOC | 93 | 93 | Main-story picture transfer now validates all markers and the complete shared Dgg/BStore, rehomes the selected graph, and has durable blob-backed replay/inversion. Groups, text boxes, noncanonical stores, transforms, extension/delay BLIPs, auxiliary stories, collision cases, broad lint debt, and current Word resave remain. |
| XLS | 93 | 93 | Ordinary plain/rich text, formulas, authored SST/XF data, and any `MulRk` member deletion with split/remerge now reopen durably. Unsafe row/column shifts, formula-owner forms, phonetic SST synthesis, duplicate rich identities, formula/range/drawing crossings, broad lint debt, and current Excel resave remain refused or absent. |
| XLSB | 94 | 94 | Durable formula dependency transfer now remaps uniquely equivalent names, external names, sheet/XTI, and structured-table dependencies; images can append while retaining an existing drawing. General chart/shape/group transfer, ambiguous dependencies, current Excel evidence, and substantial legacy crate-wide lint quarantine remain. |
| PPT | 93 | 95 | Ordinary shape IDs are rehomed into target-owned identities and bounded hyperlink/comment/media/sound references are rewritten. BLIP-bearing properties, connector rules, animation/build references, active OLE, broader presentation mutation, and current PowerPoint resave remain outside the supported root. |
| RTF | 94 | 96 | Comment bodies and footnote/endnote stories join the existing dependency-complete ordinary transfer root. Positioned or opaque dependencies, broader structural editing, and current Word/LibreOffice changed-file evidence remain gaps. |
| ODT | 94 | 94 | Header/footer/master/page-layout and protection operations now participate in durable transactions, merge, and transfer, while strict library Clippy passes. Advanced layout/security breadth, current Writer resave, and remaining broad correctness-related lint allowances keep both dimensions below release threshold. |
| ODS | 94 | 96 | Same-family automatic-style replacement/removal, wider form controls and drawing geometry, and an explicit typed decrypt/re-encrypt refusal close prior ambiguity. Full style/form/geometry breadth, encryption writes, and current Calc resave remain incomplete. |
| ODP | 94 | 96 | Source-backed whole-owner edits preserve arbitrary producer story, nested-list, table, and form markup; named drawing resources transfer with payload closure. Producer-extension breadth, crypto lifecycle operations, and current Impress resave remain incomplete. |
| OpenDocument Formula | 94 | 98 | The model now covers substantially richer Content MathML and is checked by an independent W3C schema/signature oracle, 1,024 generated trees, one-rule breakers, 4,096 arbitrary byte inputs, and two changed/reopened producer packages. Current Math resave and coverage-guided fuzz evidence remain absent. |
| ODC | 92 | 97 | Exact-span editing extends to title, subtitle, footer, and legend, but no genuine standalone producer ODC/FODC exists. Apache OpenOffice 4.1.16 and LibreOffice 26.2.5 chart-filter attempts produced neither a changed standalone save nor usable resave evidence. |
| ODG | 94 | 97 | Named gradient, hatch, fill-image, marker, opacity, and stroke-dash resources are durable; fresh packages support password encryption, signing, or both with signature verification. Existing encrypted-package rewriting and password change/re-encryption/re-signing remain unsupported, and no current Draw resave exists. |
| ODI | 94 | 96 | Genuine ODFDOM 0.13.0 original and independently load/change/save ODI packages close the synthetic-only producer gap; active-content policy and exact style-dependency classification are tested. Automatic/transitive style transfer, crypto APIs, broader forms/extensions, and native-suite resave remain absent. |
| ODM | 91 | 96 | Generated-index rename and the closed crypto capability model are durable and merge/history aware. Broader paragraph/list/table/index-child schema validation, signing/encryption writes, and current Writer resave remain incomplete. |
| OTH | 92 | 96 | Resource references/payloads and object-directory closure now join exact rich/nested durable transfer and history. The fresh builder is partial, full Relax NG coverage and security lifecycle are absent, some inline/non-top-level selectors remain refused, and no current Writer/Web resave exists. |
| ODB | 91 | 97 | Richer query/table filter, ordering, update metadata, component dependencies, and typed protected-operation dispositions improve the database root. Broader database semantics, re-sign/re-encrypt, and current Base resave remain incomplete. |
| Markdown | 98 | 98 | The full pinned CommonMark 0.31.2 and GFM corpora, exact-range reversible edits, complete release script, all-target Clippy, and rustdoc are green. This is the only reviewed format meeting both thresholds. |

## Cross-cutting findings

### Build, lint, and documentation evidence

The exact committed tree passes the all-target/all-feature check for all 19 format crates. Strict
warning-denied library Clippy also passes for all 19 formats plus `litchi-core`, `litchi-opc`,
`litchi-ooxml-common`, `litchi-odf-common`, and `xml-minifier`. Markdown's checked-in release
script passes its pinned hashes, formatting, 49 tests, all-target Clippy, and rustdoc steps; its
earlier test-field lint failure is closed. ODT's earlier 334 `unwrap`/`expect` diagnostics are also
closed.

```text
cargo check -p <all 19 format crates> --all-targets --all-features
# exit 0

cargo clippy -p <all 19 formats and reviewed common crates> \
  --lib --all-features --no-deps -- -D warnings
# exit 0

sh crates/litchi-markdown/scripts/release-gate.sh
# exit 0
```

A passing command is not proof that every lint is enabled. XLSX, DOC, XLS, XLSB, and ODT still
retain broad crate-level allowances for combinations of narrowing/sign casts, ignored results,
wildcard handling, parser-state assumptions, or related correctness families. XLSB has the largest
legacy quarantine, although its newly added dependency paths locally deny relevant lints. PPTX
retains mostly schema/API/style-oriented allowances, and shared OOXML code retains inherited
parser/style exceptions. These suppressions are reflected in the quality scores rather than being
treated as green evidence.

The reviewed feature matrices now describe the principal additions, including XLSX's formerly
stale ordinary-write rows. The matrices, ADRs, and checked-in specifications were treated as claims
to verify against code and tests, not as evidence by themselves. No conflict was found that would
justify a higher score than the implementation; several matrices appropriately record explicit
refusal and producer/native-evidence limits.

### Authored/referenced XML minimality and exact-source provenance

The authored XML publication contract remains sound. `xml-minifier::audit::verify_authored`
rejects indentation, padded markup, whitespace before tag closes, DTD/custom entities, and
ambiguous all-space character nodes outside `xml:space="preserve"`, under finite byte, event,
attribute, token, text, and depth limits. The read-side verifier remains permissive enough to
preserve genuine source whitespace.

OPC publication identifies XML through path/media-type rules, audits generated content types and
relationships, and audits every authored or changed XML-bearing part. Exact-source exemption is
limited to byte-identical XML captured from the opened package. ODF's `XmlSourcePart`,
`XmlSourceRange`, `AuthoredXmlFragment`, and `XmlSplicePublication` bind retained ranges to one
archive and part, reject foreign/stale/overlapping ranges, audit every authored tag/text fragment,
and fully parse the assembled result. The bounded contiguous-delta path preserves unrelated
producer bytes only when it can prove one exact source splice; otherwise it falls back to strict
whole-part authored verification. This is the required minimality/provenance boundary, not a
blanket exemption for referenced XML.

Focused publication evidence passes: XML authored audit 14, static OOXML assets 4 with one explicit
regeneration test ignored, ODF splice provenance 5, and OPC package writer 9. A fresh source search
finds direct `zip::ZipWriter` use in reviewed format crates only under ODS and ODM test directories.
Those call sites construct malformed, prettified, or producer-ingress negatives; no reviewed
format production source bypasses the shared audited writers.

### Durable APIs, composition, history, transfer, and security

Commit `5e5256af1` materially extends ordinary-root durability without weakening ADR-0003. Every
reviewed root keeps source-checked replay, exact or semantic inverse, stale rejection,
failure-atomic commit, deterministic join/three-way planning, bounded commit-coupled history, and a
bounded transfer disposition for its claimed owners. The most consequential additions are:

- OOXML and legacy Office roots add nested hyperlink, selected drawing/chart, formula/name/table,
  connector, picture-store, packed-cell, media/comment/sound, and additional story closure, with
  typed refusal where identity or dependency ownership is not provable.
- ODT adds advanced layout and protection transactions; ODS adds automatic-style writes and a
  final encryption disposition; ODP preserves arbitrary source-backed whole owners and transfers
  named drawing resources; ODG adds named-resource and fresh-package signing/encryption paths.
- ODI now has a genuine producer chain and bounded active-content/style policy; ODM, OTH, and ODB
  add durable index, resource/object, and query/table/component semantics. OpenDocument Formula
  adds an independent oracle/property corpus. Markdown's normative release gate is complete.

Security claims remain deliberately format-specific. ODS reports decrypt/re-encrypt as typed
unsupported rather than silently weakening protection. ODP identifies signed/encrypted packages
and refuses mutation before stage/apply. ODG supports fresh encryption/signing and verifies the
signature math but does not claim arbitrary existing-package rekey/re-sign. ODI has active-content
policy but no encryption/signature API. ODM exposes a closed capability model, and ODB has typed
protected-operation support/refusal. These explicit dispositions are quality-positive, but they do
not count as missing functional capability being implemented.

### Producer and native-application evidence

Genuine producer fixtures remain available for the established Office and ODF roots described by
their provenance files. The exercised corpus includes Office/Open XML SDK DOCX, Apache POI and
LibreOffice XLSX, PowerPoint PPTX, multi-generation Word DOC, Apache POI-corpus XLS,
producer/third-party XLSB, real PPT, and Word/LibreOffice RTF artifacts. ODT, ODS, ODP,
OpenDocument Formula, ODG, ODM, OTH, and ODB use
recorded LibreOffice-family inputs where their feature matrices claim them. Fixture identity and
lineage, not a producer-like `meta:generator` string alone, are the evidentiary boundary.

ODI now adds a pinned ODF Toolkit ODFDOM 0.13.0 producer chain using the
official `OdfImageDocument.newImageDocument()` API: an original ODI and a separately loaded,
renamed, and saved ODI have recorded artifact/library checksums, correct media type, license, member
diffs, and clean ODF Validator results. The inherited generator string is explicitly not treated as
producer identity. The synthetic FODI remains separately labeled.

ODC is the remaining standalone-producer exception. Its genuine chart XML comes from chart
subdocuments inside producer FODS/FODT packages and is not relabeled as standalone ODC/FODC.
Recorded Apache OpenOffice 4.1.16 and LibreOffice 26.2.5 attempts through `chart8` failed to produce
a changed standalone artifact: `storeAsURL()` disposed the bridge, while the in-place LibreOffice
store remained byte-identical.

Producer-created input is distinct from current application interoperability. The audited Linux
host has no Microsoft Office, LibreOffice/OpenOffice, or compatible desktop suite installed, so it
produced no Litchi-change -> native-save -> Litchi-reopen fixture for any Office or ODF format.
Registry filter declarations are recorded only as route information; XLSB and ODB are import-only
in the inspected LibreOffice declarations, ODI has no registered filter, and the ODF harness
correctly refuses ODI/ODB. The native-resave probe's four tests pass, but that validates the harness
and registry mapping, not a resave. These honest limits cap the affected completeness scores.

### Focused executable evidence

All independently selected high-risk suites passed:

| Target | Result |
|---|---:|
| DOCX document/package root | 13 passed |
| XLSX durable workbook and cell dependencies | 22 passed |
| PPTX opened presentation | 13 passed |
| DOC body/resource transaction | 5 passed |
| XLS cell/resource transaction and genuine fixtures | 21 passed |
| XLSB workbook structure/resources | 5 passed |
| PPT slide-order/media root | 12 passed |
| RTF transaction and dependency transfer | 25 passed |
| ODT package, layout, and protection transactions | 33 passed |
| ODS advanced document transaction | 11 passed |
| ODP unified and rich-content transactions | 11 passed |
| OpenDocument Formula capability/property corpora | 19 passed |
| ODC transaction | 9 passed |
| ODG package/capability transactions | 30 passed |
| ODI semantic planning and producer round trip | 11 passed |
| ODM advanced transaction | 9 passed |
| OTH semantic API | 30 passed |
| ODB advanced transaction | 15 passed |
| Markdown release tests | 49 passed, including 1,322 normative examples |
| XML authored audit | 14 passed |
| Static OOXML asset audit | 4 passed, 1 explicit regeneration test ignored |
| ODF exact-source splice provenance | 5 passed |
| OPC package writer publication | 9 passed |
| Native ODF harness/provenance probe | 4 passed |

These runs verify the scored high-risk seams; they do not substitute for current native
application resaves, continuous fuzzing, or an external full release matrix.

## Smallest actionable remediation to reach the 95 threshold

Every format below 95 in either column has a concrete remediation below. Markdown already meets
both thresholds and therefore has no score-blocking remediation.

| Format | Smallest defensible next remediation wave |
|---|---|
| DOCX | Add non-crossing cross-paragraph/composite selectors for the next common hyperlink-bearing owners, then run a complex Litchi edit through current Word save and Litchi reopen. |
| XLSX | Generalize selected leaf transfer to the supported drawing/chart relationship graph, add page-margin writes or explicitly final-scope them, localize correctness lint allowances, and validate a changed workbook in current Excel. |
| PPTX | Support or precisely classify the remaining common shape/extension families, narrow host/shared lint exceptions, and prove the changed root through current PowerPoint save/reopen. |
| DOC | Extend durable picture closure to common shared/group/text-box cases, localize broad lint assumptions, and validate the changed binary in current Word. |
| XLS | Add dependency-safe formula/range/drawing shifts and remaining common formula/SST owners, narrow broad lint allowances, and validate the changed workbook in current Excel. |
| XLSB | Extend dependency closure to ordinary chart/shape/group graphs, remove or tightly localize the legacy correctness quarantine, and validate resource-bearing changed workbooks in current Excel. |
| PPT | Add the next safe connector/BLIP/animation reference classes or final typed dispositions, broaden ordinary presentation mutation, and add current PowerPoint changed-file evidence. |
| RTF | Extend checked editing to the remaining common positioned/structural owners while preserving opaque/active refusal, then add current Word and LibreOffice save/reopen evidence. |
| ODT | Localize remaining correctness-related lint allowances, complete the next common advanced layout/security owners, and prove a changed file through current Writer save/reopen. |
| ODS | Complete common automatic-style/form/geometry families, implement or final-scope encryption writes, and prove a changed file through current Calc save/reopen. |
| ODP | Add semantic granular editing for the next producer-extension/story/form owners, implement or final-scope crypto lifecycle operations, and run current Impress save/reopen. |
| OpenDocument Formula | Add coverage-guided fuzzing around the independent oracle and prove a changed package through current LibreOffice Math save/reopen. |
| ODC | Obtain a genuine standalone producer ODC/FODC and a changed native round trip, retaining exact-span provenance for any newly encountered producer layout. |
| ODG | Complete existing-package password/signature lifecycle and remaining common advanced drawing/style semantics, then validate a changed drawing in current Draw. |
| ODI | Add automatic/transitive style closure and final encryption/signature and form/extension dispositions; native-suite evidence is only possible if a real ODI-capable producer/filter is identified. |
| ODM | Broaden child-schema validation/editing for common paragraph/list/table/index content, add or final-scope signing/encryption writes, and validate current Writer save/reopen. |
| OTH | Complete the fresh builder and common non-top-level selector coverage, add Relax NG/security evidence, and validate a resource-bearing template in current Writer/Web. |
| ODB | Deepen typed schema/query/component semantics, implement or final-scope re-sign/re-encrypt, and prove a resource-bearing changed database through current Base/UNO save and Litchi reopen. |

## Bottom line

The exact commit is materially stronger and its selected build, lint, XML, package-publication,
durability, security-disposition, and format-root tests are green. Authored XML remains compact and
audited; exact-source exemptions are provenance-bound; raw direct ZIP construction in reviewed
format crates remains test-only. ODI now has genuine programmatic-producer and changed-save
evidence, while ODC remains without a standalone producer artifact. Current native-application
changed-file evidence remains unavailable for the Office and ODF families on this host.

Markdown is the only reviewed format scoring at least 95 in both dimensions (`98/98`). The other
18 formats have at least one score below 95, so the all-format terminal threshold is not met.
