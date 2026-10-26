# Unified operations API — progress log

Driven by `docs/UNIFIED_OPS_GOAL.md`; specification `docs/UNIFIED_OPS_API_DESIGN.md`.

## Records

| No. | Title | Status |
|-----|-------|--------|
| 0349 | [Workspace gate hygiene and canonical durable JSON](0349-workspace-gate-hygiene.md) | Merged (R19 `uo2/ws-hygiene`, `55517a5d02`): feature-unification determinism fix, workspace clippy/test/doc clean |
| 0348 | [PPTX batch budget](0348-pptx-batch-budget.md) | Merged (R19 `uo2/pptx-batch`, `b70dd3e182`): every PPTX batch within 2× with batch equivalence; bench TSV split |
| 0347 | [Numbers swap, Python Commit.patch, XLSX picture single read](0347-numbers-swap-python-patch-xlsx-picture.md) | Merged (R18 `uo2/cross-small`, `0a2c9d0314`): Numbers sheet/table swap (no whole-format verb refusal left), Python `Commit.patch`, XLSX picture `set src` 1.27 → 1.07 |
| 0346 | [PPTX P0 residue, final](0346-pptx-p0-final.md) | Merged (R19 `uo2/pptx-p0b`, `8ba0c2260f`): PPTX comment fields, modern comments, notes/placeholder/slide sets; P0 planned rows 0 in every format; cold-handle cell batch miss to 0348 |
| 0345 | [DOCX P0 residue, final](0345-docx-p0-final.md) | Merged (R19 `uo2/docx-p0b`, `302f7c0e49`): add cell, the table `col` element, hyperlink tooltip, comment threads via `commentsExtended` |
| 0344 | [Closure tooling](0344-closure-tooling.md) | Merged (R18 `uo2/closure-tools`, `42c1c78ffd`): bench matrix reads every record with evidence, XLS/ODS P1 rows deferred, milestone follow-ups mapped, owner memo 25–29 |
| 0343 | [ADR 0034 glossary closure and docs hygiene](0343-glossary-closure.md) | Merged (R18 `uo2/glossary-closure`, `926cbde41e`): every 🔴 glossary cell renamed or deferred (owner memo 14–24), records index 0249–0343, facade/README/feature-matrix docs |
| 0342 | [v1 corpus closure: DOCX table borders, row properties, list styles on add and application metadata; PPTX table-level set](0342-v1-closure.md) | Merged (R18 `uo2/v1-closure`, `82b0d19475`): DOCX table borders, row properties, list styles on add, application metadata; PPTX table-level set |
| 0341 | [ADR 0008 rows in the CRUD checklist, the unreached cells, and the files that do not open (audit WP9)](0341-adr0008-rows-unreached-cells-open-failures.md) | Merged (R17 `uo2/acceptance-adr8`, `7ea8acee2a`): ADR 0008 checklist rows, unreached cells, open-failure triage |
| 0340 | [Keynote paragraph-set batches and ODC series batches scale linearly](0340-keynote-odc-batch-scaling.md) | Merged (R18 `uo2/keynote-batch`, `350404bda6`): Keynote paragraph-set and ODC series batches scale linearly |
| 0339 | [Parity residue: corpus detail, capability rows for unlisted elements, SmartArt decisions (record 0332 F1/F2, audit WP12)](0339-parity-residue.md) | Merged (R17 `uo2/parity-residue`, `69b2af2c77`): full corpus skip lists, 42 capability gap rows, SmartArt decisions |
| 0338 | [XLSX P0 residue, root-level metadata writes, exact JSON numbers and a PathArg display fixed point (WP16; record 0334 F1, F2)](0338-xlsx-p0-root-metadata-json-paths.md) | Merged (R17 `uo2/xlsx-p0-meta`, `4282e69000`): XLSX runs/cell add/type/picture src, root metadata writes, JSON float round trip, PathArg display |
| 0337 | [Performance closure: 1k-batch evidence for every format, follow-ups for every open miss, the `Path` size](0337-perf-closure.md) | Merged (R16 `uo2/perf-closure`, `2083926091`): 1k-batch evidence for 21 formats, XLS file-backed batch fix |
| 0336 | [`dump` and the raw layer: parts, raw, add-part, raw-set, validate (audit WP5)](0336-dump-and-raw-layer.md) | Merged (R17 `uo2/dump-raw`, `ed832736c3`): File::dump, raw layer (parts/raw/add_part/raw_set/validate), CLI and Python |
| 0335 | [DOCX P0 residue: picture, comment and SmartArt writes; chart series and axes (audit WP4)](0335-docx-p0-residue.md) | Merged (R16 `uo2/docx-p0`, `92d622b38f`): DOCX picture/comment set, SmartArt nodes, chart series/axis |
| 0334 | [Facade feature gates, feature-matrix check, batch-payload fuzz targets](0334-facade-features-and-batch-fuzz.md) | Merged (R16 `uo2/facade-fuzz`, `d165759e4a`): facade features from Cargo, feature matrix, batch fuzz targets |
| 0333 | [PPTX P0 residue (audit WP3)](0333-pptx-p0-residue.md) | Merged (R16 `uo2/pptx-p0`, `e39d0a75a4`): PPTX paragraph/run/row CRUD, cell fill, comment remove, picture set, coalesced batches |
| 0332 | [OfficeCLI parity dispositions and the v1 corpus subset (audit WP1 + WP13)](0332-officecli-parity-dispositions.md) | Merged (R15 `uo2/dispositions`, `0056b9cccd`): OfficeCLI parity dispositions, v1 corpus subset |
| 0331 | [Owner decision memo (2026-10-17)](0331-owner-decision-memo.md) | Confirmed by owner; item 22 redaction audit alone is deferred; items 14–20, 23, 24, 26 and 29 remain required engineering work (0350/0351) |
| 0330 | [Registry self-consistency and the XLSX cell link bug (audit WP2)](0330-registry-self-consistency.md) | Merged (R15 `uo2/registry-consistency`, `d761892de3`): element claims match property flags, XLSX batched cell links |
| 0329 | [Keynote and Pages table creation, Numbers `#REF!`, shared-formula splits, structural clears, comment edge cases, Keynote notes run styles](0329-iwork-table-creation-and-ref-errors.md) | Merged (R15 `uo2/iwork-tables2`, `84ef2cd65e`): Keynote/Pages table creation, Numbers #REF!, notes run styles |
| 0328 | [ODT, ODP and RTF element gaps against the DOCX and PPTX vocabularies](0328-odf-rtf-gaps.md) | Merged (R15 `uo2/odf-rtf-gaps`, `3bec513e6a`): ODT/ODP/RTF element gaps |
| 0327 | [XLS and ODS worksheet features (the nodes of records 0320 and 0322)](0327-xls-ods-worksheet-features.md) | Merged (R15 `uo2/xls-ods-features`, `243aab0e0b`): XLS and ODS worksheet feature nodes |
| 0326 | [PPTX query regression, path-text parser, typed coercion, element read misses](0326-pptx-regression-path-parser-typed-coercion-read-misses.md) | Merged (R14 `uo2/perf-r13`, `593d1b9fcd`): PPTX query-node size regression, path fast path, typed coercion, read misses |
| 0325 | [OOXML write follow-ups of records 0319–0322 (tier 1)](0325-ooxml-write-follow-ups.md) | Merged (R14 `uo2/ooxml-followups`, `16b800350b`): DOCX story paragraphs and marks, PPTX links, XLSX note runs, shared spreadsheet drawing |
| 0324 | [DOC element gaps against the DOCX vocabulary (tier 2)](0324-doc-element-gaps.md) | Merged (R13 `uo2/doc-gaps`, `34cb2c3a5f`): DOC styles, numbering, hyperlinks, bookmarks, fields, sections, tabs |
| 0323 | [PPT element gaps against the PPTX vocabulary (tier 2)](0323-ppt-element-gaps.md) | Merged (R13 `uo2/ppt-gaps`, `69fd64aacd`): PPT comments, hyperlinks, table cells, runs, theme/chart reads |
| 0322 | [XLSB worksheet features (the nodes of record 0320) and XLSX follow-ups](0322-xlsb-worksheet-features.md) | Merged (R13 `uo2/xlsb-features`, `a4a045b48a`): XLSB worksheet feature nodes and chart reads, XLSX fixes |
| 0321 | [PPTX element gaps for OfficeCLI tier-1 parity](0321-pptx-element-gaps.md) | Merged (R12 `uo2/pptx-p1`, `ba7669242a`): PPTX theme, columns, hyperlinks, media, equations, master/layout writes |
| 0320 | [XLSX P0/P1 worksheet features (OfficeCLI parity, tier 1)](0320-xlsx-worksheet-features.md) | Merged (R12 `uo2/xlsx-p0p1`, `931def6027`): XLSX notes, hyperlinks, drawings, CF, validation, autofilter, breaks |
| 0319 | [DOCX P0 and P1 elements for OfficeCLI parity (tier 1)](0319-docx-p0-p1-elements.md) | Merged (R12 `uo2/docx-p0p1`, `708e88727c`): DOCX styles, numbering, headers/footers, notes, tabs, breaks |
| 0318 | [Numbers formula-coordinate shift, Numbers cell comments; iWork table creation not started](0318-numbers-formula-shift-and-comments.md) | Merged (R12 `uo2/iwork-numbers2`, `3fa22ff1a4`): Numbers formula-coordinate shift and cell comments |
| 0317 | [Audit-proof follow-ups: retained proofs for DOCX commits, derived indexes for chains, depth-3 slide indexes](0317-opc-audit-proof-follow-ups.md) | Merged (R11 `uo2/opc-proofs`, `e81a4661f9`): proof hand-off, chained successor proofs, depth-3 PPTX indexes |
| 0316 | [XLS and ODP structural follow-ups (record 0308)](0316-xls-odp-structural-follow-ups.md) | Merged (R11 `uo2/xls-odp-gaps`, `588007818c`): AnchorBehavior names, XLS CF/DV/Array/Table shifts, ODP shape formatting |
| 0315 | [Query walker and read-path misses (record 0310 groups G4–G9)](0315-query-walker-read-path-misses.md) | Merged (R13 `uo2/query-walker`, `71212c8beb`): walker fast step, DOCX fenced children, XLS sheet-step cache |
| 0314 | [iWork batches: equivalence with one-at-a-time application, batch costs, and the open element gaps](0314-iwork-batch-equivalence.md) | Merged (R11 `uo2/iwork-batch`, `cbbdcd1fde`): iWork batch equivalence fixes, Pages paragraph-style loss |
| 0313 | [DOCX insertion audit proof: successor proofs from recorded content boundaries](0313-docx-insertion-audit-proof.md) | Merged (R10 `uo2/docx-insert`, `09a57ed610`): xml-minifier successor proofs for DOCX insertions |
| 0312 | [Lazy iWork tables, direct cell writes, Pages sections and Keynote placeholder styles](0312-iwork-lazy-tables.md) | Merged (R10 `uo2/iwork-tables`, `1100472b75`): lazy iWork tables, Pages sections, Keynote placeholder styles |
| 0311 | [Remaining single-op misses: the OPC double audit, the DOC adapter, the ODP and ODT wall-time puzzles](0311-single-op-performance-misses.md) | Merged (R10 `uo2/perf-opc`, `aefbcf2689`): verified OPC publication, DOC body snapshot |
| 0310 | [Open benchmark misses: root causes, re-measurement, cheap fixes](0310-miss-sweep.md) | Merged (R10 `uo2/miss-sweep`, `a4a09b30b8`): bench-miss triage, XLSB/XLSX query freshness, XLS axis writes |
| 0309 | [iWork element gaps (rows and columns, character styling, table cells)](0309-iwork-elements.md) | Merged (R9 `uo2/iwork-elements`, `a41adc99cf`): Numbers rows/columns, iWork character styling and table cells |
| 0308 | [Remaining non-iWork structural gaps (XLS, RTF, ODP, DOC, ODG)](0308-office-structural-gaps.md) | Merged (R9 `uo2/office-gaps`, `dd9ab70af5`): XLS producer shifts, RTF splice, ODP character formatting |
| 0307 | [DOCX run-level markup compatibility](0307-docx-run-markup-compatibility.md) | Merged (R8 `uo2/docx-mc`, `bcdcadadf0`): DOCX run-level markup compatibility |
| 0306 | [`litchi::iwork::Document` retired; iWork package directories and open options through `File`](0306-iwork-facade-retirement.md) | Merged (M6 `uo2/iwork-facade`, `18513c994c`): retire litchi::iwork::Document; Keynote/Numbers bundle directories |
| 0305 | [Lazy iWork trees, the DOCX single-set cost, the RTF `&Query` route](0305-lazy-iwork-trees-and-docx-single-set.md) | Merged (R9 `uo2/perf-lazy`, `5b8ebff53f`): lazy iWork parts, DOCX single set and source-backed snapshot memo |
| 0304 | [No silent content loss on text sets](0304-no-silent-loss-on-text-sets.md) | Merged (R8 `uo2/text-loss`, `123f07ac3e`): no silent content loss on text sets |
| 0303 | [PPTX/DOCX structural gaps on files from other producers](0303-pptx-docx-producer-structure-gaps.md) | Merged (R8 `uo2/pptx-gaps`, `c9017fc9e9`): PPTX owned-part slide removal, sections, F3 OLE removal |
| 0302 | [Residual single-op and batch performance misses](0302-residual-performance.md) | Merged (M6 `uo2/perf-residual`, `94f8047602`): residual single-op and batch performance |
| 0301 | [iWork structural leaf APIs (Numbers and Pages)](0301-iwork-structure.md) | Merged (M6 `uo2/iwork-structure`, `f24868dfb5`): Numbers sheet/table lifecycle, Pages directory save |
| 0300 | [iWork and tooling hygiene](0300-iwork-tooling-hygiene.md) | Merged (M6 `uo2/iwork-hygiene`, `420a04842f`): iWork and tooling hygiene |
| 0299 | [Pages to parity in the unified layer (M6)](0299-pages-parity.md) | Merged (M6 `uo2/iwork-pages`, `b44c03fdd3`): Pages W1 and unified parity |
| 0298 | [Keynote W1 and unified Keynote parity (M6)](0298-keynote-parity.md) | Merged (M6 `uo2/iwork-keynote`, `8cf481f5fd`): Keynote W1 and unified parity |
| 0297 | [Numbers to parity in the unified layer (M6)](0297-numbers-parity.md) | Merged (M6 `uo2/iwork-numbers`, `aabd78ac5f`): Numbers W1 and unified parity |
| 0296 | [iWork joins `File`: Pages, Keynote and Numbers (M6 foundation)](0296-iwork-joins-file.md) | Merged (M6 `uo2/iwork-found`, `52eca6c117`): iWork joins File; iWork-free gate removed |
| 0295 | [M4 engines on more formats, DOCX SmartArt, XLSX rich-text find](0295-m4-engines-on-more-formats.md) | Merged (R7 `uo2/m4-ext`, `8a6e6968e1`): find/merge/refresh on more formats, ODS import, DOCX SmartArt |
| 0294 | [Acceptance follow-ups: Python mirror, fixtures for unreached cells, benchmark matrix through 0302](0294-acceptance-follow-ups.md) | Merged (R7 `uo2/acceptance2`, `4139b28e13`): Python CRUD/fidelity mirror, fixtures for unreached cells, bench matrix through 0297 |
| 0293 | [Structural edits on files from other producers](0293-producer-structure.md) | Merged (R7 `uo2/producer-structure`, `873cbb31e0`): structural edits on producer XLS/PPTX/RTF files |
| 0292 | [ODF object-format verbs and flat ODF](0292-odf-verbs-and-flat-odf.md) | Merged (R7 `uo2/odf-verbs`, `622bf1ed98`): ODF object-format verbs and flat ODF |
| 0291 | [Shared target, path and plan performance](0291-shared-target-path-plan-performance.md) | Merged (R7 `uo2/perf-path`, `173b1f9fc4`): shared target/path/plan performance |
| 0290 | [Phase G checklist failures and metadata `language`](0290-checklist-failures-and-metadata-language.md) | Merged (R7 `uo2/checklist-fixes`, `dbd990cad1`): Phase G failures F1/F2/F4/F6/F7, ODF metadata language |
| 0289 | [Phase G acceptance evidence: CRUD checklist, corpus fidelity, benchmark matrix](0289-phase-g-acceptance-evidence.md) | Merged (R6 `uo2/phase-g`, `474810b463`): Phase G acceptance evidence |
| 0288 | [ODF follow-ups: shared metadata family, producer drawings, chart and image writes, object-format coalescing](0288-odf-follow-ups.md) | Merged (R6 `uo2/odf-gaps`, `742420f6f9`): ODF metadata family, ODG/ODC/ODI writes, tree coalescing |
| 0287 | [Remaining query performance misses](0287-remaining-query-performance.md) | Merged (R6 `uo2/perf-query2`, `e304079b78`): node size, walker moves, ODS/RTF/XLS query costs |
| 0286 | [OOXML adapter performance misses](0286-ooxml-adapter-performance.md) | Merged (R6 `uo2/perf-ooxml`, `9f63c06d63`): OOXML adapter performance |
| 0285 | [Find/replace fixes and consolidation with template merge](0285-find-merge-consolidation.md) | Merged (R6 `uo2/find2`, `f5918148a8`): find/merge consolidation, TOC field fix, story coverage |
| 0284 | [Typed SmartArt CRUD through the unified API](0284-typed-smartart-crud.md) | Merged (R5 `uo2/smartart`, `d9260e3692`): typed SmartArt CRUD for PPTX |
| 0283 | [DOCX charts, extended chart types and chart `set` coverage](0283-docx-charts-and-chartex.md) | Merged (R5 `uo2/charts-docx`, `055d6f6237`): DOCX charts, chartex types, chart set coverage |
| 0282 | [The remaining ODF formats in `File`: OTH, ODM, ODG, ODC, ODI, ODB and ODF formula](0282-remaining-odf-formats.md) | Merged (R5 `uo2/odf-leaves`, `dbb3d6f260`): remaining ODF formats into File |
| 0281 | [Per-format write performance misses](0281-per-format-write-performance.md) | Merged (R5 `uo2/perf-writes`, `bca745f391`): per-format write performance |
| 0280 | [Shared query and transaction performance](0280-shared-query-transaction-performance.md) | Merged (R5 `uo2/perf-query`, `b70335dc61`): shared query/transaction performance |
| 0279 | [Field and table-of-contents refresh (plan step 20c, M4 engine 3 of 3)](0279-field-refresh.md) | Merged (R5 `uo2/s20-refresh`, `491ca96c2f`): field/TOC refresh |
| 0278 | [ODP unified adapter: reads and writes over a scene index](0278-odp-unified-adapter.md) | Merged (R4 `uo2/s23-odp`, `223948be32`): unified ODP adapter |
| 0277 | [Unified ODS adapter (read + write)](0277-ods-unified-adapter.md) | Merged (R4 `uo2/s23-ods`, `f7e3ee888b`): unified ODS adapter: reads, writes, flat ODS read |
| 0276 | [Unified ODT adapter: body index, reads and writes](0276-odt-unified-adapter.md) | Merged (R4 `uo2/s23-odt`, `fcc4fe5266`): unified ODT adapter |
| 0275 | [RTF to parity (plan step 21)](0275-rtf-parity.md) | Merged (R4 `uo2/s21-rtf`, `3ea25405dc`): RTF tables, stories, links, pictures, styles, lists; source splicing |
| 0274 | [Template merge (plan step 20b, M4 engine 2 of 3)](0274-template-merge.md) | Merged (R4 `uo2/s20-merge`, `3cb803e29c`): template merge for DOCX/PPTX/XLSX |
| 0273 | [Formatting-aware find/replace (plan step 20a, M4 engine 1 of 3)](0273-find-replace.md) | Merged (R4 `uo2/s20-find`, `10f44c5a13`): formatting-aware find/replace on set |
| 0272 | [The perf-baseline harness test suite is green on trunk again](0272-perf-harness-green.md) | Merged (`uo2/harness-green`, `22d2ea5595`): perf-harness tests green on trunk, PPT aggregate routes v3, CI lint |
| 0271 | [Dotted chart series keys and honest warn-mode failures](0271-chart-series-keys-and-warn-failures.md) | Merged (`uo2/chart-series`, `92b0023604`): OfficeCLI dotted series keys, warn-mode failure honesty |
| 0270 | [Typed chart CRUD through the unified API](0270-typed-chart-crud.md) | Merged (R3 `uo2/s19-charts`, `33de040170`): typed chart CRUD for PPTX and XLSX |
| 0269 | [Closing the largest OfficeCLI corpus gaps](0269-officecli-corpus-parity.md) | Merged (R3 `uo2/corpus-parity`, `4a5f259a18`): apply-the-rest-and-warn, cell text alias, PPTX shape formatting and transitions |
| 0268 | [ODF W3: package split, atomic save and limits](0268-odf-w3-package-split.md) | Merged (R3 `uo2/s22-odf-w3`, `fd49666493`): ADR 0034 wave W3 owner split for ODT/ODS/ODP |
| 0267 | [Streaming worksheet cell cursors, small XLSB nodes, fence analysis](0267-streaming-cell-cursors.md) | Merged (R3 `uo2/stream-cursor`, `48bd7bedc4`): streaming cell cursors for early-stopping queries |
| 0266 | [CSV/TSV import and formula evaluation (plan step 18, M4)](0266-csv-import-and-formula-evaluation.md) | Merged (R3 `uo2/s18-csv-eval`, `047e8ff775`): CSV/TSV import and formula evaluation |
| 0265 | [The OfficeCLI example-corpus harness](0265-officecli-example-corpus.md) | Merged (R3 `uo2/s17-corpus`, `f592aecf28`): OfficeCLI example-corpus harness |
| 0264 | [PPT shape-text writes without whole-slide re-parses](0264-ppt-scale.md) | Merged (R2 `uo2/ppt-scale`, `4bcc2668b7`): incremental PPT commit verification and slide-memo reuse |
| 0263 | [Registry schema export and golden snapshots](0263-registry-schema-export.md) | Merged (R2 `uo2/p16-schema`, `3973a88a7e`): registry schema export and goldens |
| 0262 | [Query walker performance, eager PPTX scenes, source PPTX metadata](0262-query-walker-performance.md) | Merged (R2 `uo2/p8-perf`, `80915b091b`): query walker speedups and the 0252/0253 write misses |
| 0261 | [XLSB W2 residue, source-backed metadata and P0 writes](0261-xlsb-w2-residue-and-p0-writes.md) | Merged (R2 `uo2/p7d-xlsb`, `ab7e4b7480`): XLSB W2 residue, source-backed metadata and P0 writes |
| 0260 | [PPT reaches the P0 read and write surface (plan step 13)](0260-ppt-p0-surface.md) | Merged (R2 `uo2/p13-ppt`, `d6252a2411`): PPT P0 surface and W4 residue |
| 0259 | [XLS unified P0 read and write](0259-xls-unified-p0.md) | Merged (R2 `uo2/p12-xls`, `a804529f04`): XLS P0 surface and W4 residue |
| 0258 | [DOC P0 read+write surface through the unified API](0258-doc-p0-surface.md) | Merged (R2 `uo2/p11-doc`, `0cf3c61f43`): DOC P0 surface and W4 residue |
| 0257 | [No code references the untracked upstream checkouts](0257-no-3rdparty-references.md) | Merged (R2 `uo2/fix-3rdparty`, `e072df4db4`): fixtures vendored into test-data, guard |
| 0256 | [The `litchi` command line](0256-cli.md) | Merged (R1 `uo2/p15-cli`, `8288737b90`): litchi CLI |
| 0255 | [Python M3 surface (plan 0248 step 14)](0255-python-m3-surface.md) | Merged (R1 `uo2/p14-python`, `8def70fbcf`): Python M3 surface |
| 0254 | [Legacy OLE metadata for DOC, XLS and PPT](0254-legacy-ole-metadata.md) | Merged (R1 `uo2/p10-ole-meta`, `3f0df8f7e4`): legacy DOC/XLS/PPT metadata |
| 0253 | [Shared OOXML metadata](0253-ooxml-metadata.md) | Merged (R1 `uo2/p6-ooxml-meta`, `e81374672d`): shared OOXML metadata |
| 0252 | [PPTX ADR 0034 W2 residue and P0 writes](0252-pptx-w2-residue-and-p0-writes.md) | Merged (R1 `uo2/p7c-pptx`, `c3f3f11a8c`): pptx W2 residue and P0 writes |
| 0251f | [DOCX 10,000-paragraph fanout and the single-set overhead](0251f-docx-fanout-and-single-set.md) | Merged (R1 `uo2/p7a-docx`, `14acad03b6`): docx W2 residue and P0 writes |
| 0251e | [DOCX list definitions and `listStyle`](0251e-docx-lists.md) | Merged (R1 `uo2/p7a-docx`, `14acad03b6`): docx W2 residue and P0 writes |
| 0251d | [DOCX inline pictures: read and add](0251d-docx-pictures.md) | Merged (R1 `uo2/p7a-docx`, `14acad03b6`): docx W2 residue and P0 writes |
| 0251c | [DOCX comments: read, add and remove](0251c-docx-comments.md) | Merged (R1 `uo2/p7a-docx`, `14acad03b6`): docx W2 residue and P0 writes |
| 0251b | [DOCX package-graph changes and hyperlink writes](0251b-docx-hyperlink-writes.md) | Merged (R1 `uo2/p7a-docx`, `14acad03b6`): docx W2 residue and P0 writes |
| 0251a | [DOCX element-path writes: nested paragraphs/runs, hyperlink text, table structure](0251a-docx-element-writes.md) | Merged (R1 `uo2/p7a-docx`, `14acad03b6`): docx W2 residue and P0 writes |
| 0251 | [DOCX metadata transaction, Builder and stream output (ADR 0034 W2 residue)](0251-docx-metadata-builder.md) | Merged (R1 `uo2/p7a-docx`, `14acad03b6`): docx W2 residue and P0 writes |
| 0250 | [XLSX W2 owner reshape (ADR 0034 §2–§4)](0250-xlsx-w2-owner-reshape.md) | Merged (R1 `uo2/p4-xlsx-w2`, `b3e04abd1b`): xlsx Package owner and Workbook projection |
| 0249 | [Glossary gate hardening and truthful capabilities](0249-glossary-gate-and-truthful-capabilities.md) | Merged (R1 `uo2/a23-gate`, `3c374cea6d`): glossary gate hardening and truthful capabilities |
| 0248 | [Coordinator restart plan](0248-coordinator-restart-plan.md) | Claude resumes as sole coordinator (owner decision 2026-10-13). Survey of trunk 97ba48dcd7: 8 of 21 formats read+write, 13 still whole-format refused; trunk gates restored; prioritized steps P1–P10 |
| 0247 | [RTF direct-set coalescing](0247-rtf-direct-set-coalescing.md) | Bounded native batch planning and sequential failure/axis semantics; 1,764 full-facade / 629 RTF-only tests, tooling and gates pass; three trials pass 90 output checks and all 1k cases meet 2×, while native paragraph-text scaling, tiny-input overhead and full goal remain open |
| 0246 | [Native-backed RTF run insertion](0246-rtf-run-add.md) | Typed native insertion, unified placements/formatting, exact opaque/history preservation; 1,525 native / 1,746 full-facade / 611 RTF-only tests pass; three paired trials pass 108 output checks, with tiny-input overhead, scaling and full-goal gaps open |
| 0245 | [Opaque body anchors through RTF run edits](0245-rtf-run-opaque-anchors.md) | Native UTF-8 anchor projection, terminal-empty delimiter handling and exact history; 1,515 native / 1,738 full-facade / 603 RTF-only tests pass; six paired trials cover 10k opaque anchors, with query/scaling and full-goal gaps still open |
| 0244 | [Native-backed RTF run removal](0244-rtf-run-remove.md) | Atomic path/query removal, retained empty paragraphs and exact history; before-text staging admission fixed; 1,513 native and 1,735 facade tests/doctests pass plus late opaque-anchor regression; three paired trials retain query/scaling gaps; broader CRUD, Python and full goal remain open |
| 0243 | [Atomic RTF run text replacement](0243-rtf-run-text.md) | Native-backed text plus compound formatting, typed fan-out, Unicode coverage and empty-paragraph preservation; 1,513 native and 1,727 facade tests/doctests pass; three paired trials retain query/scaling gaps; broader CRUD, Python and full goal remain open |
| 0242 | [Atomic RTF run character formatting](0242-rtf-run-format.md) | Ten compound native-backed run setters, typed atomic fan-out, exact history and Unicode preservation; 1,504 native and 1,718 facade tests/doctests pass; three paired trials expose query overhead and scaling gaps; broader CRUD, Python and full goal remain open |
| 0241 | [Native-backed RTF run reads](0241-rtf-run-reads.md) | Borrowed run paths, local character properties, reusable plans and lexical formatting fixes; 1,497 native and 1,705 facade tests/doctests pass; three paired trials show bounded first reads and linear traversal, with generic-query overhead, run writes and full goal open |
| 0240 | [Typed RTF body-paragraph selectors](0240-rtf-body-positions.md) | Primary body edits and paragraph transfer use Position; split alias removed, consumers and glossary gate migrated; 1,495 native and 1,697 facade tests/doctests pass; wider glossary, coverage, performance and full goal remain open |
| 0239 | [Unified RTF paragraph layout](0239-rtf-unified-layout.md) | Eight native-backed layout properties, typed positions and atomic alignment/layout fanout; 1,492 native and 1,697 facade tests/doctests pass; small-document overhead, scaling and full goal remain open |
| 0238 | [Unified RTF paragraph alignment](0238-rtf-unified-alignment.md) | Typed native alignment batches, read/query/set and atomic mixed text updates; lossless alignment-only publication and boundary readback fix; 1,490 native and 1,688 facade tests/doctests pass; small-query overhead, scaling and full goal remain open |
| 0237 | [Unified RTF paragraph swaps](0237-rtf-unified-swap.md) | Atomic typed native swap, unified history/batches/save and exact no-ops; 1,485 native and 1,680 facade tests/doctests pass; changed swaps within 10% reference, no-op overhead and durable replay remain open with full goal |
| 0236 | [Unified RTF paragraph insertion](0236-rtf-unified-add.md) | Positioned native insertion, unified add/created paths/history/batches, empty and Unicode preservation, exact durable admission; 1,481 native and 1,671 facade tests/doctests pass; tiny-document overhead, broader CRUD and full goal remain open |
| 0235 | [Unified RTF paragraph moves](0235-rtf-unified-move.md) | Native placements, exact history/batches and empty-boundary preservation; durable boundary refusal; 1,476 native and 1,660 facade tests/doctests pass; paired changed moves within 10% reference, no-op overhead and full goal remain open |
| 0234 | [Unified RTF paragraph removal](0234-rtf-unified-remove.md) | Atomic original-position removal, edit/batch/history and source-preserving empty/boundary handling; 1,474 native and 1,650 facade tests/doctests pass; paired fanout measured through 10k targets; tiny-document overhead, coalescing, structural CRUD and full goal remain open |
| 0233 | [Unified RTF paragraph text writes](0233-rtf-unified-write.md) | Native paragraph set/query fanout, edit/batch/history, typed bounds and source policy; inline-break preservation and semantic patch weights fixed; structural CRUD, coalescing, performance qualification and full goal remain open |
| 0232 | [Unified RTF paragraph reads](0232-rtf-unified-read.md) | Native borrowed root/body/paragraph paths, lazy queries, children and views; truthful element gaps; RTF writes and full goal remain open |
| 0231 | [Checked RTF tail policy](0231-rtf-tail-policy.md) | Six checked append/wire bounds, EditOptions for retained/source-backed edits, and typed invalid-limit errors; publication-window policy and unified CRUD remain open |
| 0230 | [Canonical RTF tail family](0230-rtf-tail-family.md) | Focused Edit/Commit/Patch/Diagnostics/Error/LimitKind, edit_tail_append entries and typed limit refusals; redundant aliases removed; checked scoped policies and unified CRUD remain open |
| 0229 | [RTF logical-tail commit and replay](0229-rtf-tail-contract.md) | Canonical package accessors, owner replay, exact no-ops and recipient-policy admission; remaining scoped vocabulary and unified CRUD stay open |
| 0228 | [RTF semantic replacement admission](0228-rtf-replacement-budget.md) | Checked independent payload budget, admission before incoming copies, consistent composition accounting; unified CRUD and full goal remain open |
| 0227 | [Checked primary RTF semantic policy](0227-rtf-semantic-policy.md) | Checked operation-count profile and EditOptions replace unchecked Limits/edit_with_limits; typed construction/admission errors; 1,446 native and 1,620 facade tests/doctests plus lint/docs/glossary/harness checks pass; semantic payload budgets, unified CRUD and full goal remain open |
| 0226 | [RTF atomic publication](0226-rtf-atomic-publication.md) | Format-neutral atomic publisher, native RTF save policies and unified save/origin/stream dispatch; core/native/full facade suites pass; RTF CRUD, semantic profiles and full goal remain open |
| 0225 | [Native RTF Builder and unified creation](0225-rtf-builder.md) | Checked plain-text Builder returns Package; unified creation adopts native owner with explicit option gaps; 1,441 native and 1,613 facade tests/doctests pass; atomic save, unified CRUD and full goal remain open |
| 0224 | [RTF primary transaction contract](0224-rtf-primary-contract.md) | Canonical commit parts, Package::apply and exact no-op introspection; replay preserves recipient policy and identity; 1,430 native and 1,610 facade tests/doctests pass; finite semantic profiles, construction/save, unified CRUD and full goal remain open |
| 0223 | [Immutable RTF package and projection](0223-rtf-owner.md) | Concrete Package owns lifecycle; shared read-only Document and no edit::Snapshot alias; facade retains package; 1,427 native and 1,610 facade tests/doctests plus consumer/lint/docs/glossary checks pass; primary vocabulary, creation/save, unified CRUD and full goal remain open |
| 0222 | [RTF source limits and open policy](0222-rtf-source-limits.md) | Checked native SourceLimits/OpenOptions and typed ingress errors; unified options now forwarded; 1,423 native and 1,610 facade tests/doctests, minimal profile, lint/docs/glossary/boundaries pass; RTF owner/CRUD migration and full goal remain open |
| 0221 | [PPT planning over the retained scene](0221-ppt-planning.md) | Initial scene cache retained, text-only planning avoids nodes/body copies, duplicate source probes removed; 1,605 facade tests/doctests and 52 PPT-only checks pass; all 28 exact proofs unchanged; 1k ratios 1.396/1.619 meet 2×, positional single-op and rollover scaling remain open |
| 0220 | [Aggregate PPT text publication](0220-ppt-aggregate-text-publication.md) | Bounded native text chunks publish once per edit; 1,411 existing native plus eight aggregate tests and 1,603 facade tests/doctests pass; 28 v2 routes verify actual history links; unified 10k improves 19.7–20.0×, positional 1k ratio 2.080 and 26–30× growth leave budgets and full goal open |
| 0219 | [Grouped PPT drawing rewrites](0219-ppt-grouped-drawing-rewrite.md) | Strict framing/error order and exact publication bytes preserved; 1,411 native and 1,603 facade tests/doctests plus final focused checks pass; 28 paired routes verified; unified 1k improves 10.5–11.0× and 10k 5.2×, but 250–255× growth leaves scaling and the full goal open |
| 0218 | [PPT shape edit runs](0218-ppt-shape-coalescing.md) | Direct text/geometry runs share native primary edits; 1,594 existing facade tests/doctests plus nine coalescer tests, harness checks, lint/docs pass; twenty named paired routes pass exact checks; 1k overhead 2.0–2.3%, 10k growth 122–124× leaves native linear scaling and full goal open |
| 0217 | [Unified PPT slide structure](0217-ppt-slide-structure.md) | Native-backed remove/move/swap with history and capability dispatch; 1,585 existing facade tests/doctests plus nine structural tests pass; 18 paired local cases below 10% median overhead; insertion, shape CRUD, coalescing and full goal remain open |
| 0216 | [Unified PPT setters](0216-ppt-unified-set.md) | Native text/anchor and existing visibility-record writes; path/query, edit/batch/history and immutable publication; 1,585 facade tests/doctests pass; paired write overhead 6.4–39.2% (budget gap); structural CRUD, nested writes, batch coalescing and full goal remain open |
| 0215 | [Unified PPT navigation and views](0215-ppt-unified-read.md) | Native-backed get/find/query/views, typed anchor conversion, shared static shapes, source fences and capability golden; native/facade suites and lint/docs pass; native read scaling, unified PPT mutation and full goal remain open |
| 0214 | [Signed PPT primary transactions](0214-ppt-signed-noops.md) | Exact signed no-ops preserve source bytes; changed writes return typed Signed; retained leniency, durable replay and facade classification covered; native/facade suites, feature/lint/docs/glossary checks pass; unified PPT navigation/mutation and full goal remain open |
| 0213 | [Canonical PPT Builder and creation](0213-ppt-builder-creation.md) | Immutable build output, explicit encryption entropy, preserved speaker notes and unified create/publication; native/facade suites pass; signed no-op edits, unified navigation/mutation, performance and full goal remain open |
| 0212 | [Canonical PPT primary transaction](0212-ppt-primary-edit.md) | Native-backed immutable Edit/Patch/Commit, finite operation policy, source fences and retained leniency; native/facade tests and lint/docs checks pass; signed no-ops, Builder, unified PPT read/write, performance and full goal remain open |
| 0211 | [Immutable PPT owner and owned presentation](0211-ppt-immutable-owner.md) | Canonical shared Package/Presentation, explicit raw Reader, exact atomic publication and facade adoption; native/facade regressions pass; primary PPT transactions, Builder, unified read/write and full goal remain open |
| 0210 | [Checked PPT source policies](0210-ppt-source-limits.md) | Canonical checked SourceLimits/OpenOptions, typed errors and all-stream admission before payload reads; native/facade tests, lint/docs, feature/glossary/boundary checks pass; primary immutable PPT owner, unified PPT read/write and full goal remain open |
| 0209 | [XLS publication dependency scans and patch ownership](0209-xls-publication-work.md) | Profile-backed removal of scalar selection scans and deep native patch cloning; 3,342 tests/doctests and lint/docs/checks pass; measured 10k atomic sets ~17% faster, removals 33–36% faster; source 1k removal ratio 2.093× remains open; repeated bounded publication and complete performance/CRUD goal remain open |
| 0208 | [XLS edit inventory and staging indexes](0208-xls-edit-indexes.md) | Native indexed lookups/overlap checks and facade retained-source cell lookup; join/no-op correctness and source fences covered; 3,339 tests/doctests plus six final index regressions, lint/docs/glossary/boundaries pass; paired ascending 10k batches improve 32–44%, reverse-order cases exact; general scaling and full goal remain open |
| 0207 | [XLS cell transaction coalescing](0207-xls-cell-coalescing.md) | Distinct direct scalar sets/removals share native transactions, with semantic barriers and bounded native-limit rollover; 1k operations use one commit; paired runs exact, eager 1k batch ~1.15–1.16× native, source removal borderline at 2×; native indexing, linear scaling and full performance/CRUD goal remain open |
| 0206 | [Unified XLS cell removal](0206-xls-cell-removal.md) | Native-backed path/range/query removal, edit/batch/history, packed numeric/blank closure and corrected filesystem write coverage; 3,318 tests/doctests plus lint/docs/checks pass; paired removal exact, small-source timing qualification remains open; broader CRUD, streaming, coalescing and full goal remain open |
| 0205 | [Unified XLS scalar writes](0205-xls-unified-writes.md) | Native-backed scalar/clear/name setters, edit/batch/history, lossless SST closure and explicit inverse gaps; 3,595 tests/doctests plus final capability/write checks, lint/docs/glossary/boundaries/parity pass; paired eager numeric write overhead below 0.7%; broader CRUD, coalescing, source streaming and full goal remain open |
| 0204 | [Eager XLS query cursors](0204-xls-eager-query-cursors.md) | Direct native eager cell iteration and selected-tab metadata remove sheet/catalog-sized setup; bounded-allocation regressions pass; source-backed and axis streaming, unified XLS mutations and full goal remain open |
| 0203 | [Unified XLS reads](0203-xls-unified-reads.md) | Native-backed navigation/query/views, complete tab catalog, source fences and explicit gaps; tests/lint/docs/glossary pass; eager paired reads miss performance target and traversal early termination remains open; unified XLS mutations and full goal remain open |
| 0202 | [Explicit XLS encryption entropy](0202-xls-explicit-encryption-entropy.md) | Single-use explicit entropy removes ambient RNG from Builder build; 2,242 tests/doctests, lint/docs/glossary/boundaries pass; Package security publication, unified XLS CRUD and full goal remain open |
| 0201 | [XLS encryption profile](0201-xls-encryption-profile.md) | Retained exact encryption mode/key length and canonical password opening; 2,323 tests/doctests, lint/docs/glossary/boundaries pass; security publication verbs, Builder entropy, unified XLS CRUD and full goal remain open |
| 0200 | [Canonical XLS Builder](0200-xls-canonical-builder.md) | Concrete Builder and Package-returning build replace Writer publication; encrypted owners retain explicit output; 2,234 tests/doctests plus focused Builder checks, lint/docs/glossary/boundaries and harness compilation pass; encrypted-build determinism, unified XLS CRUD and full goal remain open |
| 0199 | [XLS creation and publication](0199-xls-creation-publication.md) | Native minimal Package, unified creation and atomic publication, origin retention, first-cell row-table creation and explicit option gaps; 2,229 combined tests/doctests plus final owner/publication checks pass; Builder, unified XLS CRUD and full goal remain open |
| 0198 | [XLS primary transactions](0198-xls-primary-transactions.md) | Package Edit/Commit/Patch/Diagnostics, retained profiles and positional reads, exact replay/inverse, typed staging limits and publication fences; 2,223 combined tests/doctests plus 1,747 final native checks, lint/docs/glossary/boundaries and consumer compilation pass; Builder, facade XLS CRUD/batching and full goal remain open |
| 0197 | [Immutable XLS owner](0197-xls-immutable-owner.md) | Native Package and owned Workbook, raw reader migration, facade delegation, exact publication and retained edit policies; 2,768 tests/doctests plus final owner checks, lint/docs/glossary/boundaries and consumer compilation pass; primary transactions, Builder, unified XLS CRUD and full goal remain open |
| 0196 | [XLS eager source limits](0196-xls-source-limits.md) | Checked SourceLimits, typed errors and pre-payload admission; canonical option setters and credential redaction; 2,199 tests/doctests, lint/docs/glossary/boundaries and consumer checks pass; positional policy preserved; canonical owner, unified XLS writes and full goal remain open |
| 0195 | [XLS retained limits](0195-xls-retained-limits.md) | Object/CFB policies retained across edits, replay, overlays, joins and history; shared candidate/layout gates; 5,844 combined tests/doctests plus final history checks, lint/docs/glossary/boundaries and consumer checks pass; canonical owner/API migration and full goal remain open |
| 0194 | [XLS cell-edit vocabulary](0194-xls-edit-vocabulary.md) | Concrete Edit, pair split, is_noop and exact Diagnostics::changed; 2,178 tests/doctests, lint/docs/glossary/boundaries and consumer checks pass; retained-limit gap identified; immutable owner, unified XLS writes and full goal remain open |
| 0193 | [DOC strict-owner admission](0193-doc-strict-owner-admission.md) | Typed preflight before strict payload capture across body opens/reopens; one parsed CFB retained; 5,276 combined + 1,913 final DOC/facade tests/doctests, lint/docs/glossary/boundaries pass; full goal remains open |
| 0192 | [DOC source limits](0192-doc-source-limits.md) | Canonical checked SourceLimits, inline typed errors, shared defaults and all-stream reader preflight; 2,713 tests/doctests, lint/docs/glossary/boundaries pass; focused strict-owner allocation ordering and full goal remain open |
| 0191 | [DOC primary transactions](0191-doc-primary-transactions.md) | Native Package Edit/Commit/Patch/Diagnostics, retained policies and validated owner handoff; facade adopts canonical package; 2,707 tests/doctests, lint/docs/glossary/boundaries pass; scoped ratios under 10%, absolute large-case slowdown and full goal remain open |
| 0190 | [Immutable DOC owner](0190-doc-immutable-owner.md) | Concrete immutable Package, cheap detached Document, merged OpenOptions and explicit raw reader; encrypted/exact publication checks; 2,696 combined tests/doctests, lint/docs/glossary/boundaries pass; repeated 100-paragraph overhead 13.6%, primary transactions and full goal remain open |
| 0189 | [DOC paragraph controls](0189-doc-paragraph-controls.md) | Opt-in policy and retained-reader diagnostics; repeated exact outputs; retained native reader reproduces similar large-corpus overhead; release build/Clippy and smoke pass; performance acceptance and full goal remain open |
| 0188 | [DOC paragraph loop](0188-doc-paragraph-loop.md) | Native Builder/Package publication and borrowed body traversal; unified paragraph text create/read/set/batch/history/save; guarded limits and atomic publication; 3,231 combined + 11 final focused tests, lint/docs/glossary/boundaries pass; exact paired outputs, up to 20.4% overhead leaves performance acceptance open; broader DOC CRUD and full goal remain open |
| 0187 | [XLSB sheet removal](0187-xlsb-sheet-removal.md) | Native source-catalog removal sets and durable tag 10; unified sheet paths/queries, explicit dependency refusals and exact history accounting; 484 shared + 1,225 native + 1,258 facade tests and final focused checks pass; lint/docs/glossary/boundaries/parity pass; repeated paired overhead 0.5–3.2%; disposition handling, batch scaling and full goal remain open |
| 0186 | [XLSB worksheet insertion](0186-xlsb-sheet-insertion.md) | Native staged insertion with dependency remapping, blank sheet/index authoring and durable tag 9; unified add for eager/source owners; 477 shared + 1,196 native + 1,250 facade tests and final focused checks pass; lint/docs/glossary/boundaries/parity pass; repeated paired overhead 1.0–2.5%; sheet removal, batch scaling and full goal remain open |
| 0185 | [XLSB sheet ordering](0185-xlsb-sheet-order.md) | Semantic native permutation with dependency remapping and durable tag 8; unified move/swap for eager/source owners; 1,645 shared/native + 1,243 facade tests, final focused checks, lint/docs/glossary/boundaries/parity pass; repeated paired overhead 1.3–4.8%; insertion/removal, broader CRUD and full goal remain open |
| 0184 | [XLSB name dependencies](0184-xlsb-name-dependencies.md) | Shared lexical dependency rewrite, native tag 7 with historical replay retained, typed facade gaps; 1,609 shared/native + 2,398 XLSX + 1,237 facade tests and final focused checks pass; lint/docs/boundaries/parity pass, repeated overhead below 5.4%; structural verbs and full goal remain open |
| 0183 | [XLSB sheet-name selectors](0183-xlsb-sheet-name-selectors.md) | Complete-catalog native/facade name selection, legacy patch normalization and atomic duplicate admission; 1,154 native / 1,234 facade tests, full/minimal lint, docs/glossary/boundaries/parity pass; repeated paired timings exact and within 10% overhead; rename dependency closure, structural verbs and full goal remain open |
| 0182 | [Indexed source materialization](0182-indexed-source-materialization.md) | Both OPC conversion paths preflight namespaces with an index and retain source/authored provenance; 1,162 OPC / 1,151 XLSB / 1,230 facade tests plus eight focused checks pass; repeated source 10k time improves about 55%, eager stable; lint/docs/boundaries/parity pass; full scaling qualification and goal remain open |
| 0181 | [XLSB batch payload publication](0181-xlsb-batch-payload-publication.md) | One OPC signature revocation per built-in payload batch, strict custom-part refusal and opaque-record coverage; OPC/native/combined facade suites and lint/docs/boundaries/parity pass; repeated 10k native medians improve about 90% eager / 75% source; residual scaling and full goal remain open |
| 0180 | [XLSB pending sheet admission](0180-xlsb-pending-sheet-admission.md) | Pending visibility runs avoid repeated catalog scans; 561 facade tests/doctests and Clippy pass; paired 1–1k matrices exact, repeated 1k facade overhead below 2%; native signature-scan scaling diagnosed and remains open; full goal open |
| 0179 | [XLSB sheet-command coalescing](0179-xlsb-sheet-command-coalescing.md) | Homogeneous positional visibility commands share one native edit with active-view and semantic barriers; 1,229 full facade tests plus 12 final focused checks, lint/docs/minimal/boundaries/parity pass; repeated 1–10k matrices exact and one-commit, measured 1k/10k below 1.18x native; linear scaling and general sheet-command coalescing remain open |
| 0178 | [XLSB catalog validation indexes](0178-xlsb-catalog-indexes.md) | Indexed native name/physical-target admission preserves Unicode/OPC equivalence; 1,161 native / 1,218 facade tests, Clippy and parity pass; repeated 1k query medians improve about 63%, remaining superlinear publication and 10k acceptance stay open; full goal remains open |
| 0177 | [XLSB sheet visibility](0177-xlsb-sheet-visibility.md) | Native complete-tab state edits and unified hidden/visibility properties; active/selected-view closure and typed gaps; 1,159 native / 1,217 combined tests plus 24 final focused checks, lint/docs/glossary/boundaries/parity pass; repeated 1–1k query measurements equivalent, linear scaling and 10k acceptance remain open; full goal remains open |
| 0176 | [XLSB cell removal](0176-xlsb-cell-removal.md) | Native idempotent removal and unified cells/ranges/queries with atomic command coalescing; indexed native lookup fixes quadratic scaling; 1,144 native / 1,207 combined tests plus 34 final focused checks, lint/docs/glossary/boundaries/parity pass; repeated paired matrices meet measured one-cell and batch budgets; full goal remains open |
| 0175 | [XLSB creation and canonical builder](0175-xlsb-creation.md) | Native `Builder::build` / `Package::new` and unified creation; repeated-build SST fix; 1,138 native / 1,197 combined tests/doctests, lint/docs/glossary and consumer checks pass; explicit creation-option gaps; full goal remains open |
| 0174 | [XLSB publication projection reuse](0174-xlsb-publication-projection.md) | Package-owned cached projection preserves exact admitted output and policies; 1,133 native / 1,194 combined tests/doctests, lint/docs/glossary pass; two after matrices meet measured one-cell 10% threshold, 1k/10k below 2x; full goal remains open |
| 0173 | [XLSB direct cell resolution](0173-xlsb-direct-cell-resolution.md) | Shared path resolution defers redundant read grids for direct writes; 1,194 combined tests/doctests, lint/docs/minimal/boundaries pass; 12 paired cases exact, large-batch ratios improve, eager small-edit 10% requirement remains open (1.109 set / 1.118 edit); full goal remains open |
| 0172 | [XLSB scalar cell writes](0172-xlsb-scalar-writes.md) | Native scalar assignment and shared-string closure; unified cells/ranges/queries, coalesced batches, exact no-ops/replay and retained quotas; 1,128 native / 521 XLSB facade tests, 1,190 combined before final guards; paired 1k/10k within 2x, eager one-shot 1.186 remains a performance follow-up; full goal remains open |
| 0171 | [XLSB package cell-value transactions](0171-xlsb-package-cell-values.md) | Package-owned Edit/Commit/Patch and diagnostics replace Workbook mutation entries; raw worksheet family is explicit; immutable exact replay, retained policies, bounded images/deltas and complete-tab selectors; final 1,109 native and 501 facade tests/doctests passed; unified cell writes and full goal remain open |
| 0170 | [XLSB publication leniency](0170-xlsb-publication-leniency.md) | Cell-value/watch/sparkline and structural target publication retain opt-in cached-number policy and raw bytes; legacy structural replay shares finite destination admission; 1,096 native and 501 facade tests/doctests plus seven final regressions passed; owner migration and full goal remain open |
| 0169 | [XLSB unified worksheet-name writes](0169-xlsb-unified-sheet-name.md) | Native package-backed sheet.name set/transactions/batches and atomic save; complete-tab Unicode admission, signed/source preservation, typed quotas and tolerated-record regressions; full suites plus final focused checks, lint/docs/parity/boundaries; eight paired cases exact, ratios 0.997–1.082; full goal remains open |
| 0168 | [XLSB primary package transactions](0168-xlsb-primary-package-transactions.md) | Package edit/apply and canonical Commit triad replace structural Workbook entry points; exact durable replay, atomic input-budget admission and bounded chaining covered; 1,087 native and 489 facade tests/doctests passed; Clippy/docs/glossary/boundaries and two harness checks passed; scratch cleaned; full goal remains open |
| 0167 | [XLSB package policy retention](0167-xlsb-package-policy-retention.md) | Consuming Workbook-to-Package conversion retains source, signatures and projection policies; package reconstructions preserve skipped drawings and opt-in cached-number reads; 1,084 full-feature native and 489 facade tests/doctests passed; Clippy/rustdoc/glossary/boundaries passed; scratch cleaned; structural owner/commit migration and full goal remain open |
| 0166 | [XLSB structural diagnostics](0166-xlsb-structural-diagnostics.md) | Canonical changed/no-op introspection and exact signed same-name replay; 1,067 default / 1,078 full-feature native and 489 facade tests/doctests passed; Clippy/rustdoc/glossary/boundaries passed; scratch cleaned; owner migration audited; full goal remains open |
| 0165 | [XLSB structural transaction names](0165-xlsb-structural-transaction-names.md) | Module-local structural transaction family without old aliases; 1,066 default / 1,077 full-feature native and 489 facade tests/doctests passed; Clippy, rustdoc, glossary, boundaries and parity passed; temporary artifacts cleaned; native owner/commit contract, staging costs and unified XLSB writes remain open |
| 0164 | [PPTX slide creation properties](0164-pptx-slide-creation-properties.md) | Slide add accepts name/hidden with strict aggregate admission, one native edit, source/eager replay and batch coverage; 1,125 facade tests/doctests and ten native checks passed; Clippy/docs/minimal/boundaries/parity passed; ten paired cases equivalent (source exact, eager decoded members), ratios 0.995–1.009; artifacts cleaned; full goal remains open |
| 0163 | [Combined PPTX slide properties](0163-pptx-combined-slide-properties.md) | Source primary combined setter and unified eager/source adoption avoid duplicate staging; 23 focused tests/doctests, Clippy and full-feature glossary passed; ten paired combined-native cases exact, one-shape ratios 1.002 eager / 1.029 source; artifacts cleaned; full goal remains open |
| 0162 | [PPTX slide properties](0162-pptx-slide-properties.md) | Eager/source native and unified slide name/hidden writes with strict atomic admission, exact no-ops and durable replay; 1,717 native and 1,122 facade tests/doctests plus final 34 focused checks passed; Clippy/docs/glossary/boundaries/parity passed; ten paired cases exact, one-shape ratios 1.042 eager / 1.014 source; 0159–0162 committed once Git metadata became writable; temporary artifacts cleaned; full goal remains open |
| 0161 | [PPTX shape removal](0161-pptx-shape-removal.md) | Native source/eager and unified top-level removal with reference guards, explicit durable v4 grammar and reversible topology; 1,117 facade and 1,708 native tests/doctests plus final 20 removal checks passed; Clippy/docs/glossary/boundaries/parity passed; ten paired cases exact, one-shape ratios 1.038 eager / 1.022 source; larger eager overhead remains a follow-up; artifacts cleaned; commit blocked by read-only Git metadata; full goal remains open |
| 0160 | [Publication resource errors](0160-publication-resource-errors.md) | XML publication budgets retain `limit_exceeded` and allocation failures retain `internal_error`; native source chain preserved; 1,111 facade and 164 XML tests/doctests, Clippy/rustdoc/minimal build/boundaries/parity passed; artifacts cleaned; commit blocked by read-only Git metadata; full goal remains open |
| 0159 | [PPTX fanout publication](0159-pptx-fanout-publication.md) | Same-slide selectors share native chunk publication with byte/target admission; 1,110 facade tests/doctests, Clippy/rustdoc/boundaries/parity checks passed; all 20 paired cases equivalent, 10k ratios 1.005–1.238 and unified 10k/1k growth 9.749–9.973x; artifacts cleaned; commit blocked by read-only Git metadata; full goal remains open |
| 0158 | [PPTX chunk publication](0158-pptx-chunk-publication.md) | Native multi-chunk publication and facade adoption; 1,688 native and 1,108 facade tests/doctests passed; 1k/10k ratios below 2x and unified 10k/1k growth 9.66–10.05x; Clippy/glossary/boundaries passed; implementation/evidence committed concurrently; artifacts cleaned; own staging blocked by read-only Git metadata; full goal remains open |
| 0157 | [PPTX slide-root admission](0157-pptx-slide-root-admission.md) | Distinct slide/notes node limits and preserved typed resource errors; 1,665 native and 1,107 facade tests/doctests passed; eager 10k workloads pass with 1.016–1.022 ratios; Clippy/glossary/boundaries passed; artifacts cleaned; commit blocked by read-only Git metadata; linear scaling and full goal remain open |
| 0156 | [PPTX command coalescing](0156-pptx-command-coalescing.md) | 1,000 direct shape commands share one native commit; 1,107 combined tests/doctests plus Clippy passed; 1k paired ratios 1.005–1.224; linear scaling fails and eager 10k exposes masked native node limit; artifacts cleaned; commit blocked by read-only Git metadata; full goal remains open |
| 0155 | [One eager edit per PPTX setter](0155-pptx-single-edit.md) | Removed discarded native edit; paired eager one-shape overhead is 4–5% in the measured generated corpus; 1,097 facade tests/doctests and Clippy passed; artifacts cleaned; commit blocked by read-only Git metadata; command coalescing, wider performance acceptance and full goal remain open |
| 0154 | [PPTX selector fanout](0154-pptx-selector-fanout.md) | Atomic text/geometry selection beyond 256 targets; 1,096 facade tests/doctests passed; paired 1–1,000-shape evidence finds eager small-write overhead above budget; Clippy passed; artifacts cleaned; commit blocked by read-only Git metadata; command batching, performance acceptance and full goal remain open |
| 0153 | [Atomic PPTX text and geometry writes](0153-shape-geometry-writes.md) | Native eager/source and unified combined text/bounds edits with durable replay; 1,663 native and 1,092 facade tests/doctests, Clippy/docs/glossary/boundaries passed; commit blocked by read-only Git metadata; broader CRUD/performance/full goal remain open |
| 0152 | [Text-box output admission](0152-text-box-output-admission.md) | Caller output limits checked before managed memory admission; 11 native and six facade tests plus Clippy passed; artifacts cleaned; commit blocked by read-only Git metadata; full goal remains open |
| 0151 | [Source-backed text-box addition](0151-source-text-box-addition.md) | Native and unified source text-box creation, canonical durable suffix proof and reference-aware IDs; 1,642 native and 1,090 combined facade tests/doctests passed; Clippy/docs/glossary/boundaries passed; commit blocked by read-only Git metadata; broader shape/format CRUD and full goal remain open |
| 0150 | [Source-backed slide insertion](0150-source-slide-insertion.md) | Native and unified source slide insertion, empty/disjoint durable topology and mixed lexical relationship edits; 1,144 OPC, 1,630 PPTX and 1,089 combined facade tests/doctests passed; Clippy/docs/glossary/boundaries passed; commit blocked by read-only Git metadata; source shape creation and full goal remain open |
| 0149 | [Unified source slide removal](0149-unified-source-slide-removal.md) | Source File/query/transaction/batch removal and effective catalog/related reads; 1,614 native and 1,081 combined facade tests/doctests passed; Clippy/docs/glossary/boundaries passed; commit blocked by read-only Git metadata; source insertion, broader CRUD and full goal remain open |
| 0148 | [Primary source slide removal](0148-primary-source-slide-removal.md) | Readable structural revisions and durable v3 replay/inverse after reopening; 1,614 native and 1,078 combined facade tests/doctests passed; Clippy/docs/glossary/boundaries passed; commit blocked by read-only Git metadata; facade source removal and full goal remain open |
| 0147 | [Exact prepared metadata images](0147-prepared-metadata-images.md) | Bounded source/changed/appended/omitted metadata handles retain managed admissions; 1,141 OPC, 1,589 native PPTX and 1,078 combined facade tests/doctests passed; OPC Clippy/docs and crate boundaries passed; commit blocked by read-only Git metadata; primary topology/durable integration and full goal remain open |
| 0146 | [Source-native slide removal plan](0146-source-slide-removal-plan.md) | Lazy OPC removal with strict ownership/metadata guards and exact root splicing; 1,589 native and 1,078 combined facade tests/doctests, Clippy/docs/glossary/boundaries passed; commit blocked by read-only Git metadata; primary durable source removal and full goal remain open |
| 0145 | [Guarded source topology metadata](0145-source-topology-metadata.md) | Source-native exact relationship/content-type XML restoration behind physical and typed guards; 1,135 OPC, 1,583 native PPTX and 1,078 combined facade tests/doctests passed; Clippy/docs/glossary/boundaries passed; commit blocked by read-only Git metadata; primary source removal and full goal remain open |
| 0144 | [Source PPTX removal prerequisites](0144-source-removal-prerequisites.md) | Shared lexical binding removal and source layout validation; physical XML limits enforced on master/layout/theme; 1,583 native and 1,078 combined tests/doctests plus Clippy/docs/glossary passed; boundaries passed and temporary artifacts cleaned, commit blocked by read-only Git metadata; primary source removal and full goal remain open |
| 0143 | [PPTX modeled order metadata](0143-pptx-modeled-order-metadata.md) | Stock source Move/Swap and typed presentation extensions preserve exact metadata; 1,578 native, 562 PPTX facade and 1,078 combined tests/doctests passed; Clippy/docs/glossary/boundaries passed; commit blocked by read-only Git metadata; source topology and full goal open |
| 0142 | [Primary source PPTX slide ordering](0142-pptx-primary-source-order.md) | Source Move/Swap composed with text, logical read mapping, stable bindings and durable v2 proofs; 1,572 native, 561 PPTX facade and 1,077 combined tests/doctests, Clippy/docs/glossary passed; commit blocked by read-only Git metadata; source topology, positional metadata and full goal open |
| 0141 | [Composable source PPTX slide ordering](0141-pptx-source-order-composition.md) | Native staged moves/swaps, checked positions, exact cancellation and managed root reservations; 1,557 native, 559 PPTX facade and 1,075 combined tests/doctests, Clippy/docs/glossary passed; commit blocked by read-only Git metadata; primary catalog/durable integration and full goal open |
| 0140 | [Unified PPTX slide removal](0140-pptx-unified-slide-removal.md) | Eager path/scoped-query removal, native dependency/final-slide guards and typed refusal mapping; 558 PPTX and 1,074 combined tests/doctests, Clippy/docs/glossary passed; commit blocked by read-only Git metadata; richer dependencies/source topology and full goal open |
| 0139 | [Unified PPTX slide ordering](0139-pptx-unified-slide-order.md) | Eager move/swap through File, transactions and batches; exact placement/no-op/replay/undo and pre-promotion refusal; 548 PPTX and 1,064 combined tests/doctests, Clippy/docs/glossary passed; commit blocked by read-only Git metadata; removal/source topology and full goal open |
| 0138 | [Unified PPTX text-box addition](0138-pptx-unified-text-box-addition.md) | Eager shape/textbox creation, strict geometry and native CR preservation; 987 native, 543 PPTX facade and 1,059 combined tests/doctests passed; Clippy/docs/glossary passed; commit blocked by read-only Git metadata; source structural creation and full goal open |
| 0137 | [Unified PPTX slide addition](0137-pptx-unified-slide-addition.md) | File/transaction/batch add for created and eager owners, native preservation/proofs and explicit source-owner capability gap; 538 PPTX and 1,054 combined facade tests/doctests, Clippy/docs/glossary passed; commit blocked by read-only Git metadata; source topology, broader CRUD and full goal open |
| 0136 | [Native PPTX slide insertion](0136-pptx-native-slide-insertion.md) | Checked position/exact layout insertion, durable replay/inverse and topology dependency guards; 1,001 native tests, 20 native doctests and 533 facade tests/doctests passed; Clippy/docs/glossary passed; commit blocked by read-only Git metadata; facade CRUD and full goal open |
| 0135 | [PPTX part relationship metadata resolution](0135-pptx-part-metadata-three-way.md) | Physical part choices retain matching graph/XML; unrelated content-type and owner edits remain automatic; 986 native and 533 facade tests/doctests, Clippy/docs/glossary passed; commit blocked by read-only Git metadata; type/topology projection and full goal open |
| 0134 | [PPTX package-root three-way merge](0134-pptx-root-three-way.md) | Compatible root IDs merge automatically; divergent graph/XML candidates resolve with Base/Left/Right and durable inverse preservation; 983 native and 533 facade tests/doctests, Clippy/docs/glossary passed; commit blocked by read-only Git metadata; part/content-type conflict resolution and full goal open |
| 0133 | [PPTX durable metadata replay](0133-pptx-durable-metadata-replay.md) | LPTX0003 per-key metadata replay, exact forward/inverse publication and deterministic disjoint joins; 978 native, 1,126 OPC and 533 facade tests/doctests passed; Clippy/docs/glossary passed; commit blocked by read-only Git metadata; three-way metadata conflict resolution and full goal open |
| 0132 | [Guarded OPC content-type fragments](0132-opc-content-type-fragments.md) | Native-key Default/Override transport, shared metadata planner and stricter no-op revalidation; 1,126 OPC tests/doctests, 967 native PPTX tests, Clippy/docs passed; commit blocked by read-only Git metadata; durable metadata wire integration and full goal open |
| 0131 | [Guarded OPC relationship fragments](0131-opc-relationship-fragments.md) | Bounded lexical entry transport with source/context guards and exact inverse restoration; full OPC suite, 13 final focused tests, Clippy/docs and native PPTX check passed; commit blocked by read-only Git metadata; content-type transport, native wire integration and full goal open |
| 0130 | [Semantic PPTX slide-order selectors](0130-pptx-slide-order-selectors.md) | Canonical slide selector, semantic move/swap/reorder and private raw-ID lowering; 982 native and 533 facade tests/doctests, harness, Clippy and glossary passed; commit blocked by read-only Git metadata; metadata transport and full goal open |
| 0129 | [Guarded OPC metadata restoration](0129-opc-guarded-metadata-restoration.md) | Exact-source raw metadata installation with bounded parsing/audit and physical member presence; 1,103 OPC tests/doctests, 967 native PPTX tests and Clippy passed; commit blocked by read-only Git metadata; native durable wire integration and full goal open |
| 0128 | [PPTX empty slide-list prerequisite](0128-pptx-empty-slide-list.md) | Empty-list lexical insertion and bounded IDs; 967 native tests plus metadata replay diagnostic and 99 integration tests and Clippy passed; durable forward metadata preservation remains prerequisite to public insertion; commit blocked by read-only Git metadata; full goal open |
| 0127 | PPTX canonical constructor | Public package writer/font setters removed; native 1,515 and combined facade 1,081 tests/doctests passed; direct new-package edits/replay/inverse verified; remaining mutators/full goal open; commit blocked |
| 0126 | PPTX builder consumer migration | Public authoring consumers migrated; 12 unified write + 10 harness tests, 21/24 root and both native smoke examples passed; 3 existing codec/adapter failures recorded; native constructor/full goal open; commit blocked |
| 0125 | PPTX native builder and unified creation | Direct canonical native owner, typed option gaps and creation capabilities; 1,081 combined facade tests/doctests and native builder regression passed; Clippy/both glossary profiles passed; commit blocked by read-only Git metadata; full goal remains open |
| 0124 | PPTX immutable encrypted publication | Retained-profile re-encryption integrated with common atomic save; 531 facade tests/doctests, 13 native encryption tests and final public regressions passed; Clippy/glossary passed; commit blocked by read-only Git metadata; full goal remains open |
| 0123 | PPTX unified shape-text writes and publication | Native immutable setters, transactions/batches and atomic save integrated; 962 native + 1,076 combined facade tests/doctests passed; Clippy/glossary/review passed; commit blocked by read-only Git metadata; broader writes and full goal remain open |
| 0122 | PPTX committed source revision readback | 26 native primary tests, 17 native doctests and 519 facade tests/doctests passed; Clippy/glossary/review passed; commit blocked by read-only Git metadata; facade writes and full goal remain open |
| 0121 | Shared immutable PPTX native promotion | 518 facade tests/doctests and final signature regression passed; Clippy/glossary passed; commit blocked by read-only Git metadata; facade writes and full goal remain open |
| 0120 | PPTX durable source shape-text patches | 960 encrypted native tests, 17 native doctests and 516 facade tests/doctests passed; Clippy/glossary/review passed; commit blocked by read-only Git metadata; full goal open |
| 0119 | PPTX immutable source package ownership | 92 source, 8 owned-scene, 17 native doctests and 516 facade tests/doctests passed; Clippy/glossary/review passed; commit blocked by read-only Git metadata; full goal open |
| 0118 | PPTX committed package and immutable patch replay | 210 encrypted native opened-editor, 22 policy/marked-XML, 16 native doctests and 516 facade tests/doctests passed; consumers/features/glossary and Clippy passed; commit blocked by read-only Git metadata; full goal open |
| 0117 | PPTX patch admission and visible payload ownership | 203 default + 203 encrypted native opened-editor, 22 policy/marked-XML and 516 facade tests/doctests passed; Clippy clean; commit blocked; package-carrying commits and full goal remain open |
| 0116 | PPTX primary Edit entry and drop semantics | 193 native opened-editor, 22 policy/marked-XML, one native example and 516 facade tests/doctests passed; glossary, Clippy, harness and public example checks passed; commit blocked; full goal open |
| 0115 | XLSX border writes and typed local formatting | 1,012 combined tests/doctests, 90 stylesheet codec, 64 native cell-edit and 3 state tests passed; native examples, capabilities, Clippy and XLSX glossary passed; commit blocked; full goal open |
| 0114 | XLSX gradient-fill writes | 1,001 combined tests/doctests, 81 stylesheet codec, 57 native cell-edit and 78 style tests passed; canonical native replacement, bounded multi-stop input and typed join-budget recovery; examples, capabilities, Clippy and XLSX glossary passed; commit blocked; full goal open |
| 0113 | XLSX solid and pattern-fill writes | 995 combined tests/doctests, 75 stylesheet codec and 50 native cell-edit tests passed; typed native fill changes, facade foreground/pattern/background policy, Clippy/native example, both capability profiles and XLSX glossary check passed; commit blocked; full goal open |
| 0112 | XLSX number-format writes | 988 combined tests/doctests, 66 stylesheet codec, 44 native cell-edit and 75 style tests passed; native resource reuse/allocation and unified aliases; Clippy/native example clean; commit blocked; full goal open |
| 0111 | XLSX protection and remaining alignment attributes | 982 combined tests/doctests, 56 stylesheet, 39 native cell-edit and 73 native style checks passed; five setters and optional justification readback; 50 final facade checks and Clippy clean; commit blocked; full goal open |
| 0110 | XLSX alignment writes | 975 combined + 45 final facade cell tests/doctests, 48 stylesheet and 37 native cell-edit checks passed; Clippy and native example clean; shared font/alignment lowering and seven setters; commit blocked; full goal open |
| 0109 | XLSX remaining typed font properties | 967 final DOCX/XLSX/XLSB/encryption tests/doctests, 38 stylesheet, 38 font-focused and 39 final cell-write checks passed; seven font read/write properties, checked integer bounds and explicit XLSB gaps; Clippy clean; commit blocked; full goal open |
| 0108 | XLSX font color writes | 929 combined + 34 final cell tests/doctests, 32 stylesheet and 32 font-focused checks passed; typed RGB/theme/tint, lexical preservation and ordered insertion; Clippy clean; full goal open; commit blocked |
| 0107 | XLSX font name, size and script writes | 924 combined + 29 final cell tests/doctests, 26 stylesheet and 28 font-focused checks passed; Clippy clean; owned names, checked size/script, bounded quote-safe writes; full goal open; commit blocked |
| 0106 | XLSX strike-through and underline writes | 919 combined + 49 final write tests/doctests, 20 stylesheet and 22 font-focused checks passed; Clippy clean; native strike/underline, aliases and strict XML envelope; full goal open; commit blocked |
| 0105 | Reuse explicit XLSX merge-list planning | 914 combined tests/doctests and 18 focused merge tests passed; Clippy clean; per-worksheet admission reuse retains selected-node limits and atomic companions; full performance/coverage goal open; commit blocked |
| 0104 | XLSX merge operations and range readback | 911 combined + 419 minimal tests/doctests, 37 native + 21 source merge tests and final 15 unified regressions passed; Clippy clean; bounded native merges, atomic companions and table guards; broader coverage/full goal open; commit blocked |
| 0103 | XLSX row and column layout | 895 combined + 419 minimal tests/doctests and 14 delta tests passed; Clippy clean; native height/width/visibility/outline/collapse assignments with bounded selectors; broader writes/full goal open; commit blocked |
| 0102 | XLSX worksheet lifecycle | 883 combined + 419 minimal tests/doctests and final 34 targeted regressions passed; Clippy clean; worksheet add/remove, tab rename/visibility/move/swap and native patch integration; broader coverage/full goal open; commit blocked |
| 0101 | XLSX distinct-cell batch coalescing | 870 combined + 419 minimal tests/doctests, 3 native state tests and 9 native doctests passed; Clippy clean; source/eager 1k scalar batches within 2× native in scoped ABBA runs; broader coalescing/full goal open; commit blocked |
| 0100 | Preserve ignorable XLSX stylesheet extensions | 16 codec + 11 native font + 859 combined tests/doctests and 14 delta tests passed; Clippy clean; bounded MCE admission, opaque preservation and font insertion order verified; broader formatting/full goal open; commit blocked |
| 0099 | XLSX native and unified font assignments | 2,254 native + 858 combined + 419 minimal tests/doctests, final 11 codec and 14 delta tests passed; Clippy clean; lossless bold/italic resources, range writes, joins, replay and inverse verified; stylesheet profiles/compaction and full goal open; commit blocked |
| 0098 | XLSX range and scoped-selector writes | 855 combined + 419 minimal tests/doctests, 14 target-write regressions and 14 delta tests passed; Clippy clean; one native commit, distinct-cell/staging budgets and early rectangle admission verified; structural/coalesced writes and full goal open; commit blocked |
| 0097 | XLSX cell writes and native transaction history | 2,233 native + 20 final native + 841 combined + 419 minimal tests/doctests and 14 delta tests passed; Clippy clean; direct cell writes, rollback, replay, opaque preservation and protection reports verified; range/structural/coalesced writes and full goal open; commit blocked |
| 0096 | Immutable XLSX source promotion | 561 XLSX encryption + 817 combined tests/doctests and final concurrent regression passed; Clippy clean; exact source, successful promotion reuse, retry and signature isolation verified; cell writes/full goal open; commit blocked |
| 0095 | XLSX native creation and publication | 814 combined DOCX/XLSX + 558 XLSX encryption + 418 minimal tests/doctests and 14 delta tests passed; strict output/origin/signature policies verified; XLSX mutations and managed publication/full goal open; commit blocked |
| 0094 | DOCX cached-result field invalidation | 2,282 native + 703 DOCX + 418 minimal tests/doctests and 14 delta tests passed; exact dirty-marker replay/inverse and atomic rollback verified; additional namespace replay regression and Clippy passed; full goal open; commit blocked |
| 0093 | DOCX command paragraph-add splitting | 700 DOCX + 418 minimal tests/doctests, 39 final batch and 14 delta tests passed; Clippy clean; atomic LF splitting/VT inline controls verified; full goal open; commit blocked |
| 0092 | Literal first-node DOCX command run text | 2,277 native + 697 DOCX + 418 minimal tests/doctests; 14 delta tests, Clippy and replay review passed; paragraph splitting/field cache/managed edits/full goal open; commit blocked |
| 0091 | Reuse DOCX body during write admission | 692 DOCX tests/doctests and Clippy passed; controlled 1k fanout at 1.13× eager / 1.10–1.11× source native; small-write/full goal open; commit blocked |
| 0090 | Borrow DOCX immediate write assignments | 691 DOCX tests/doctests and Clippy passed; deferred plans retain ownership, paired timing shows no clear improvement; full goal open; commit blocked |
| 0089 | Paired DOCX paragraph text fanout baseline | Two 9-round ABBA invocations, byte/ledger equivalence and Clippy passed; small-write overhead exceeds review trigger, 10k typed refusal; full goal open; commit blocked |
| 0088 | DOCX paragraph selector command text controls | 690 DOCX + 418 minimal tests/doctests, 31 batch and 14 delta tests passed; single-pass target admission, Clippy clean; broader text/full goal open; commit blocked |
| 0087 | Native bulk dispatch for DOCX paragraph text fanout | 688 unified DOCX tests/doctests and 3 final fanout regressions passed; Clippy clean; eligible fanout uses native bulk, performance/full goal open; commit blocked |
| 0086 | Native DOCX bulk paragraph inline controls | 1,397 native + 12 bulk integration + 685 unified tests/doctests passed; managed publication/inverse and resource release verified, Clippy clean; full goal open; commit blocked |
| 0085 | Direct DOCX paragraph inline text controls | 1,397 native + 9 managed + 685 DOCX + 418 minimal tests/doctests; Clippy/delta checks passed; direct paragraph command controls integrated, broader text/full goal open; commit blocked |
| 0084 | DOCX command run-add text controls | 683 DOCX tests/doctests, 32 focused batch/capability and 14 delta tests passed; Clippy clean; input run controls integrated, other text contexts/full goal open; commit blocked |
| 0083 | Decode typed DOCX scalar controls in scope | 1,397 native + 19 source/content + 680 unified tests/doctests passed; 3k typed controls need zero owner scans; malformed scalar payloads refused; Clippy clean; full goal open, commit blocked |
| 0082 | Reuse DOCX run namespace context | 1,395 native + 19 source/content + 680 unified tests/doctests passed; zero owner scans for 1k captured run text reads, Clippy clean; broader context/performance/full goal open; commit blocked by read-only Git metadata |
| 0081 | Contextual DOCX paragraph/run text reads | 1,393 native + 5 content + 10 tracker + 39 source/streaming + 680 unified tests/doctests passed; Clippy clean; context reuse, text policy and full goal open; commit blocked by read-only Git metadata |
| 0080 | Namespace-correct DOCX typed run content | 1,397 final native unit/integration + 679 unified tests/doctests; fixture targets recovered (36 tests), Clippy clean; text-policy completion remains open; commit blocked by read-only Git metadata |
| 0079 | Shared formatting runs for DOCX closure edits | 678 DOCX +418 minimal tests/doctests and final19 transaction tests passed; typed targets, staged reads, poison and best-effort isolation verified; wider coalescing/full goal open; commit blocked by read-only Git metadata |
| 0078 | Paired DOCX batch scaling and target reuse | 452 focused tests + final 25 batch regressions passed; repeated-target 1k batch measures 1.08–1.10× leaf across eager/source; broader performance acceptance open; commit blocked by read-only Git metadata |
| 0077 | Atomic DOCX direct-path formatting runs | 669 DOCX + 418 minimal tests/doctests, final 24 batch regressions and 2 bisection probes passed; native-capacity splitting and retained XML reduction verified; performance/wider coalescing open; commit blocked by read-only Git metadata |
| 0076 | DOCX atomic and best-effort command batches | 1,039 all-feature + 665 final DOCX tests/doctests and 26 final batch/capability regressions passed; atomic/best-effort core integrated; coalescing/full command coverage open; commit blocked by read-only Git metadata |
| 0075 | Bounded DOCX explicit transactions | 1,021 all-feature + 417 minimal tests/doctests and final 24 transaction/capability regressions passed; owner-bound replay/limits verified; coalescing/batch remain open; commit blocked by read-only Git metadata |
| 0074 | Complete internal DOCX operation commits | 1,002 all-feature + 413 minimal tests/doctests; patch/diagnostics retained across all five verbs, replay/inverse/no-ops verified and empty-target admission fixed; public transactions/coalescing remain open; commit blocked by read-only Git metadata |
| 0073 | Bounded OfficeCLI batch payload ingress | 994 all-feature + 413 minimal tests/doctests; complete transport vocabulary, caller-selected byte/item limits and typed redacted failures verified; executor/coalescing remains open; commit blocked by read-only Git metadata |
| 0072 | DOCX creation and atomic path saving | 970 all-feature ops + 596 DOCX-only + 389 minimal + 1,097 OPC tests/doctests; typed creation, origin retention, guarded overwrite policy and encrypted save round trips verified; broader writes/batches remain open; commit blocked by read-only Git metadata |
| 0071 | Immutable DOCX structural operations | 946 all-feature ops + 2,257 DOCX leaf + 384 minimal tests/doctests; paragraph/run add/remove/move/swap, exact slot preservation and add registry verified; batches/path saving remain open; commit blocked by read-only Git metadata |
| 0070 | Immutable DOCX set and stream publication | 917 all-feature ops + 2,252 DOCX leaf + 1,093 OPC tests/doctests; 68 setters and exact eager/source no-ops verified; structural writes/batches/path saving open; commit blocked by read-only Git metadata |
| 0069 | Property inputs, strict registry admission and error attribution | M2 W1/W2 foundation: 529 DOCX + 359 minimal tests/doctests; File setters/batches pending; commit blocked by read-only Git metadata |
| 0068 | XLSB unified read adapter and indexed scalar fallback | R4d integrated: 1,076 leaf + 418 ops + 468 XLSX regression tests/doctests; extended gaps/write/performance open; commit blocked by read-only Git metadata |
| 0067 | XLSB retained grids and checked sparse reads | L-XLSB-R2 verified: 1,072 leaf + 401 ops tests/doctests; R4d pending; commit blocked by read-only Git metadata |
| 0066 | XLSB semantic tab catalog | L-XLSB-R1 verified: 1,055 leaf + 400 ops tests/doctests; R2/R4d pending; commit blocked by read-only Git metadata |
| 0065 | XLSX trait read boundaries and capability publication | 468 tests/doctests and Clippy passed; infallible metadata migration open; commit blocked by read-only Git metadata |
| 0064 | PPTX unified read adapter | R4c integrated; 2,310 tests/doctests passed, 3 ignored; extended gaps/performance open; commit blocked by read-only Git metadata |
| 0063 | Source-backed PPTX modern comments | R9 read foundation verified; unified PPTX dispatch pending; commit blocked |
| 0062 | Source-backed PPTX presentation catalog | R4 read foundation verified; full goal active; commit blocked |
| 0061 | Source-backed PPTX related parts | R3 read foundation verified; modern comments and full charts pending; commit blocked |
| 0060 | Shared DrawingML text bodies | R7 codec/property projection verified; full policy migration pending; commit blocked |
| 0059 | PPTX lazy shape properties | R6 properties/raw identity verified; commit blocked |
| 0058 | Source-backed PPTX scenes and bounded retention | R1/R2 source read/cache verified; R3/R6/R7 and R4c pending; commit blocked by read-only Git metadata |
| 0057 | Independently retained PPTX scenes and token classification | R2 eager foundation and R10 verified; source memo and R4c pending; commit blocked by read-only Git metadata |
| 0056 | XLSX unified read adapter and File dispatch | R4b integrated and verified; broader M1 and performance acceptance pending; commit blocked by read-only Git metadata |
| 0055 | XLSX readable formatting and source table access | L-XLSX-R3/R4 integrated and verified; adapter and performance acceptance pending; commit blocked |
| 0054 | XLSX retained grid, metadata projection and early-stop leaf reads | L-XLSX-R1/R2/R6 integrated; formatting/tables, adapter and performance acceptance pending; commit blocked |
| 0053 | Capability reports and generated OfficeCLI comparison | DOCX discovery integrated; parity/registry incomplete; commit blocked by read-only Git metadata |
| 0036 | OPC source-limit kinds and canonical errors | physical kind/error fields migrated; broader limits policy pending |
| 0035 | OPC/OOXML source-limit vocabulary | shared physical profile and builder renamed; broader limits policy pending |
| 0034 | Effective source DOCX text streaming | retained main/donor output integrated; specialized routes pending |
| 0033 | Effective source DOCX metadata reads | shared property graph/decoder integrated; remaining specialized routes pending |
| 0032 | Typed primary DOCX transaction limits | typed kind/count fields migrated; configurable limits and broader policy pending |
| 0031 | Canonical source DOCX package-owning Edit/Commit | primary owner migrated; specialized families, limits and performance pending |
| 0030 | Immutable retained DOCX semantic apply | main-document apply integrated; independent managed donor snapshots refused |
| 0029 | Source DOCX effective owner and physical-operation boundary | effective reads/publication integrated; specialized routes pending |
| 0028 | Effective OPC owning conversion | explicit materialization integrated |
| 0027 | DOCX patch predicate and content-free Debug | primary patch migrated; broader diagnostics pending |
| 0026 | Owned reads and physical identity access for retained OPC candidates | effective read adapter foundation; integrated in 0029–0031 |
| 0025 | Immutable XML edits of retained effective OPC candidates | flat chained source-Part candidates; primary owner integrated in 0031 |
| 0024 | Revision-bound source XML edits over retained payloads | per-part chaining foundation; package overlay revisions pending |
| 0023 | Retained OPC prepared topology and shared ZIP appended members | source-overlay ownership foundation; chained commits integrated in 0031 |
| 0022 | Eager DOCX package-bound Edit/Commit and removal of mutating publication doors | eager primary commit migrated; source joined in 0031; history and full W2 pending |
| 0021 | Shared source-backed DOCX owners and borrowing stream publication | sharing prerequisite complete; primary package-bound triad integrated in 0031 |
| 0020 | Cheap eager DOCX package shares and immutable primary patch application | eager apply migrated; package-owning commit W2 work pending |
| 0019 | Owned eager DOCX read projection and shared physical graph | projection migrated; immutable package and commit W2 work pending |
| 0018 | DOCX primary edit entry and focused error/result names; minimal/full glossary bans | names migrated; package-owned commit/projection W2 work pending |
| 0017 | W0 rustdoc API glossary gate, scoped rules, six-type family checks and CI | complete; profiles for leaf migration waves pending |
| 0016 | Rebase onto current base; adopt API glossary as ADR 0034; preserve shared fingerprint memo | kickoff; W0 checker and leaf migrations pending |
| 0001 | Owner override: read + write for every claimed format; tiers OOXML → OLE → (RTF, ODF, iWork) | recorded in the goal (2026-10-04) |
| 0002 | Leaf write-capability inventory (deep: OOXML + OLE; brief: tier 3) | merged (`cfe7681627`) |
| 0003 | M0a facade migration plan (flat `File`, `litchi::open`) | merged (`45a932ea2e`) |
| 0004 | ADR 0033 draft + `litchi-ops` scaffold + gate registration | merged (`af70966a6d`) |
| 0005 | M0b path/selector grammar in `litchi-core::selector` | merged (`9f8e49b181`) |
| 0006 | Owner accepts ADR 0033 | done |
| 0007 | Coordinator defaults for 0003's open questions | adopted, awaiting owner |
| M0a | Steps S1–S10 on `uo/m0a` (plan 0003 §4) | merged (`1927763795`); S10 gate passes with E1/E2 awaiting owner (F1, F2) |
| 0008 | Diagnosis: source-backed DOCX edits refuse every Word-authored fixture (MCE) — M2 blocker | merged (`e9c3868109`); fix staged as 0011 |
| 0009 | Design: editing a shared snapshot (R4) — M2 prerequisite | done (`9b3cab60f6`); staged leaf plan S0–S5 |
| 0010 | M1 unified-read plan (node view, path resolution, query, view modes, capabilities) | done (`31bd660497`); steps R1–R11 with leaf waves L1/L2 |
| M1 | Steps on `uo/m1` (plan 0010 §10); merges after M0a | in progress |
| 0011 | DOCX MCE admission, stages 1–3 of 0008 | merged into `uo/m0a` (`53f938e83c`); 19/19 Word fixtures editable source-backed; eager owner gains the same refusals (breaking) |
| 0013 | Signature policy: `Signed::unsign`, builder refusals, pptx/xlsb silent-strip fixes (M0a S6b) | merged into `uo/m0a` (`f30c521c90`) |
| 0015 | M2 write-kernel plan (verbs, registry, per-format handlers, batch executor, leaf prerequisites, save/create) | done (`de8d557c56`); steps W1–W10 and leaf prerequisites L-*-W* |
| M2 | Steps on `uo/m2` (plan 0015 §11); merges after M0a and M1 | in progress |
| 0014 | Implicit `unsign()` sites: pptx and shared ooxml-common (0014a), xlsx/docx and XLSB task panes/Ribbon (0014b); 0013 F1, F1b, F2 | merged into `uo/m0a`; silent strips now refuse until explicit `unsign` (breaking) |
| 0012 | litchi-opc: cheap `Clone` for `SourceBackedPackage` (0009 stage S0) | merged into `uo/m0a` (`0cf1f43403`); clone 6 reference counts, no allocation; open unchanged |

## Awaiting owner confirmation

- Record 0248 decisions (provisional coordinator defaults, 2026-10-13):
  - D3: tiers stay strict. Phase F (RTF/ODF/iWork) waits until tiers 1–2 reach P0 CRUD; the RTF work already landed is kept.
  - D4: earlier provisional decisions remain provisional.
  - D5: budget misses get engineering first. Explanations come only after a fix attempt.

## Wave R1 (2026-10-13, coordinator restart)
Branches `uo2/*` cut from `908f9cccdd`; records 0249–0255:
- `a23-gate`: plan steps 2–3, the glossary gate and truthful capabilities.
- `p4-xlsx-w2`: step 4, the xlsx W2 reshape.
- `p7a-docx`: docx W2 residue and P0.
- `p7c-pptx`: pptx W2 residue and P0.
- `p6-ooxml-meta`: shared OOXML metadata.
- `p10-ole-meta`: legacy metadata.
- `p14-python`: the Python surface.

- Merged `a23-gate` (0249, `47a44b5506`):
  - The glossary gate now checks 21 profiles, 673 bans and 8 families.
  - DOC `capabilities()` lists 22 element gaps.
  - The delta table shows read and write status per property; 48 write-only properties are no longer counted as read-supported.
  - Not done: facade profile (unresolvable glob re-exports), core `Limits` ban (owner decision).

- Merged `p10-ole-meta` (0254, `3f0df8f7e4`):
  - DOC/XLS/PPT `Builder::set_metadata` and the `package::metadata` family;
  - create-time metadata, and a set-only `/metadata` path;
  - unified overhead +0.7–6.3%.
- Merged `p15-cli` (0256, `8288737b90`): new `litchi-cli` crate, the `litchi` binary.
- Merged `p14-python` (0255, `8def70fbcf`): the Python read and write surface; 242 tests.
- Merged `p4-xlsx-w2` (0250, `b3e04abd1b`):
  - xlsx `Package` owner, with `Workbook` as its projection; `workbook::edit` registered;
  - 81 new bans;
  - no regression above 5%.
- New owner rule (2026-10-13, AGENTS.md `e467bcf4e8`): code never references `3rdparty/`; fixtures are copied into `test-data/`. The trunk worktree's untracked partial `3rdparty/` directory, left by the Codex session, was moved aside to the session scratchpad.

- Merged `p7c-pptx` (0252, `c3f3f11a8c`):
  - pptx W2 residue: `OpenOptions`, `PatchConflict`, typed `EditLimit`, `opened` family registered, scoped `Transaction`/`rollback` removed;
  - P0 writes: rich text, cells, notes, comments, picture add, layout placeholders and sections;
  - eager run formatting at ≥100 shapes misses the budget (+17–22%), a follow-up for plan step 8;
  - picture `set` remains open.

## Wave R2 (2026-10-13)
Cut from `e467bcf4e8`:
- `fix-3rdparty` (0257): vendor the fixtures into `test-data/`, plus a guard.
- `p11-doc` (0258), `p12-xls` (0259), `p13-ppt` (0260): tier 2 P0.

- Merged `fix-3rdparty` (0257, `e072df4db4`):
  - no code references `3rdparty/`; 376 fixtures (1.29 MB) vendored with provenance;
  - the delta tool reads its vendored OfficeCLI schemas;
  - guard `tools/test_no_3rdparty_references.py`, allowlist now down to the 4 docx files owned by p7a.
- Merged `p6-ooxml-meta` (0253):
  - one `litchi_core::Metadata`, `Props` deleted;
  - a shared exact core/app/custom transaction;
  - pptx/xlsx/xlsb `metadata` families; `/metadata` set and create-time metadata.
  - Misses: sets on source-backed PPTX (×2.5) and XLSB (×1.4).
  - Conflicts with 0252 resolved in `pptx/write.rs` (both dispatches kept) and in the glossary files.
- Known pre-existing failure on trunk, not from this wave: `litchi-oth` `semantic_api::raw_zip_negatives_are_rejected_by_the_open_boundary` (a DTD package is accepted). Tier 3 ODF follow-up.
- The `ods_formula_conditional_native` test target cannot compile: its 5.5 MB fixture was purged on 2026-10-05. It needs a reduced fixture.

- Merged `p11-doc` (0258):
  - DOC reads: runs, tables, stories, notes and comments (element gaps 22 → 13).
  - DOC writes: run and paragraph formatting; paragraph add, remove, move and swap.
  - Writes cost 0–8.5% over the leaf. The 0190 13.6% read miss no longer reproduces.
  - W4: `open_with_limits*` and `rollback` removed and banned.
  - Golden rows merged by hand: 3 DOC rows carry both 0258's text and 0253's metadata front-matter.

- Merged `p13-ppt` (0260):
  - PPT readers ported; shape text can change length on 129 of 214 corpus shapes (was 0).
  - Notes text set, blank slide insert, name/alt reads.
  - Single-set overhead 1.39× → 1.00×.
  - 10k batches still scale 35×. The owner asked for a fix: commit verification without a full re-parse, step 0264.
  - DOC and PPT `add` arms were unioned in the ops dispatch.
  - `test_threads.test_long_writes_release_the_gil` flaked once under load (3/3 reruns pass).

- Merged `p16-schema` (0263, `3973a88a7e`):
  - `litchi_ops::schema(format)` with registry goldens;
  - the delta table now reads the registry and covers DOC, PPT and RTF;
  - exposed through the CLI and Python.
- Merged `p12-xls` (0259):
  - XLS sheet lifecycle, formula writes, row/column layout, XF formatting writes, read-only named ranges;
  - writes cost 1.00–1.07× (file-backed layout ~1.10×).
  - Read overhead of 1.6–7× is traced to the shared resolver and walker and handed to the p8 perf stream.
  - Feature-subset build fixes made at the merge: signature-test imports, an encryption-gated test, and the `round_trip` dead code in pptx-less builds.

- Merged `p7a-docx` (0251–0251f):
  - Leaf: docx metadata transaction, `Builder` and `write_to`.
  - Writes: nested paragraph/run/hyperlink/table; hyperlink add and retarget; comments add and remove; picture add; list definitions and `listStyle`.
  - A 10k fanout is one write (op cap 4,096 → 65,536), ~1.05×.
  - Misses: single set 1.18×; cell writes ~1.24×.

- Merged `p7d-xlsb` (0261, `ab7e4b7480`):
  - XLSB W2 residue: one `OpenOptions` replaces 37 routes; `raw::Encoder`; three families registered.
  - Source-backed metadata sets are ×1.02–1.04 (was ×1.37–1.39).
  - P0 formats, formulas, merges and named ranges, all ≤1.08.
- Fixed `ce003ff150`: the harness metadata test used the deleted `Props`. Harness gates now build `--all-targets`.
- Merged `p8-perf` (0262, `80915b091b`):
  - one freshness check per step; typed comparisons; smaller nodes; O(1) XLSX records;
  - `first()` in O(1);
  - the 0252 and 0253 write misses fixed (1.00–1.01×);
  - full queries still miss R4b (XLSX 10.5–13.7×, PPTX 6.7–11.8×).
  - The query-corpus golden went stale through corpus growth (more vendored fixtures). The pre-0262 and merged builds print byte-identical digests over today's corpus, so it was re-recorded.
- **Owner decision (2026-10-14): accepted.** A PPTX filter on one key reads only that key, so a malformed unrelated property no longer fails the filter; reading it directly still errors (record 0262).

Still running from R1: `p7a-docx` (0251), `p7c-pptx` (0252), `p6-ooxml-meta` (0253).

- Merged `ppt-scale` (0264, `4bcc2668b7`):
  - incremental commit verification, with the full re-parse as fallback and oracle (~10.9k differential cases);
  - the slide memo is seeded across chunked edits;
  - native 10k batch 108.6 → 50.6 ms; 1k→10k per-op cost 3.30× → 1.67×.
  - **Owner decision (2026-10-14):** raise the PPT primary-edit and aggregate-selection bounds from 4096 to 16384 (measured 1.07×). Being applied on `uo2/ppt-scale`.
- Merged `s17-corpus` (0265):
  - all 96 OfficeCLI example scripts replay in CI through the `litchi` CLI;
  - 6,431 calls: 2,003 pass, 3,616 gap, 139 delta, 673 cascade, 0 bugs; no panic or `internal_error`.
  - Top gaps: xlsx cell `text` alias (906), pptx shape `size` on add (461), chart add (763), shape fill/geometry (255), transition (161), apply-the-rest-and-warn (2,172).

## Wave R3 (2026-10-14)
Cut from `4d1afe92bc`:
- `s17-corpus` (0265): the OfficeCLI example-corpus harness.
- `s18-csv-eval` (0266): CSV/TSV import and formula evaluation (M4).
- `stream-cursor` (0267): streaming cell cursors and the remaining query cost.
- `s22-odf-w3` (0268): ADR 0034 W3 for ODF (tier 3 starts now that tiers 1–2 have P0 read+write).

Still running: `ppt-scale` (0264, the owner-requested PPT 10k batch fix).

Added later in R3: `corpus-parity` (0269) and `s19-charts` (0270), then follow-ups `chart-series` (0271) and `harness-green` (0272), both cut from `fd49666493`.

- Merged `ppt-scale` bound change (0264 §4, `8de68ed60d`):
  - `MAX_OPERATIONS` and `MAX_SHAPE_TEXT_CHUNK_SELECTIONS` are now 16384;
  - a 10k batch publishes once, and 1k→10k per-op cost is 1.09× native and unified;
  - unified/native at 10k is 1.10 (was 1.25).
- Merged `corpus-parity` (0269, `4a5f259a18`):
  - command-channel `batch::Unsupported::Warn` (`apply_with`/`batch_with`, CLI `--unsupported warn`, exit 9); the library default stays strict;
  - cell `text` is an alias of `value`; PPTX shape formatting on add, fill and geometry, slide transitions;
  - X1/X2 refusal codes unified.
  - Misses: `Warn` batches are about 49× (no coalescing, P2); PPTX add 1.4× (existing overhead, P1).
- Merged `s19-charts` (0270, `33de040170`):
  - typed chart add/get/set/remove for PPTX and XLSX on a shared deterministic `chart::authoring::Spec`;
  - PPTX `add_chart` now takes a `Spec` (breaking);
  - new XLSX `chart::edit` family;
  - DOCX charts (0270a), chartex types and waterfall are still open.
- Merged `s18-csv-eval` (0266, `047e8ff775`):
  - `File::import` for CSV/TSV into XLSX/XLSB/XLS, as one atomic commit, within 2.1% of the leaf;
  - read-only `File::evaluate`/`recalculate`; new `litchi-ops → litchi-eval` edge;
  - `litchi-eval` fixes: a non-ASCII criteria panic, Excel coercion, and a depth limit;
  - 26.3% of the corpus's 12,082 cached formulas reproduce (79.4% of those compared). XLSX, at 70.2%, is limited by the engine's grammar (F7).
- Merged `stream-cursor` (0267, `48bd7bedc4`):
  - cell cursors for early-stopping queries; cold first match is 21× faster on XLSX, 17.7× on XLSB and 7.9× on XLS (50k rows);
  - 396 query digests byte-identical; the XLSB node is now 88 bytes.
  - **Owner decision (2026-10-14): accepted.** On a sheet broken past the rows read, an early query (`first()`, `limit(n)`) returns the leading matches, and the refusal is reported only if the walk reaches the defect.
- Merged `s22-odf-w3` (0268, `fd49666493`):
  - ODT/ODS/ODP each have an immutable `Package` plus read-only projections;
  - every ODF whole-package save is atomic; typed `SourceLimits`/`SemanticLimits`/`OpenOptions`;
  - the litchi-oth DTD test was changed to match the shared ingress policy (an ignored DOCTYPE is accepted; an entity-declaring DTD is still refused).
  - G1–G8 are open (other ODF crates, `Mutable*` writers, scoped names).
- Combined gate after these merges:
  - every `litchi-ops` single-feature build except XLSX was broken by two ungated 0269 test files; fixed in the 0271 merge.
  - XLSX chart add was refused by 0269's kind-first check; fixed in 0271.
  - Gate scripts now run clippy through `non_iwork_gate.py`, so iWork is excluded; facade tests with `--all-features`; harness tests with `--locked`.
- Merged `chart-series` (0271):
  - dotted `series{N}.name/.values/.color` on chart add;
  - `Error::warnings` names the properties a failed warn-mode call skipped;
  - corpus now 3,803 passed, 1,105 warned, 891 gap, 0 bugs.
- Merged `harness-green` (0272):
  - perf-harness tests are green on trunk:
    - the producer-scanner tests were re-derived for the intended `5fe77e0bb6` change;
    - case count set to 633;
    - lockfiles, formatting and 62 clippy findings fixed;
  - PPT aggregate routes renamed `_v2` → `_v3`;
  - the CI smoke job now lints the harness, and its path filters are pinned by policy tests.

## Wave R4 (2026-10-14)
Cut from `395d0ffa9f`:
- `s20-find` (0273): formatting-aware find/replace as `find=`/`replace=` on `set` (M4 engine 1).
- `s20-merge` (0274): template `merge` (M4 engine 2).
- `s21-rtf` (0275): RTF to parity (plan step 21).
- `s23-odt`, `s23-ods`, `s23-odp` (0276–0278): unified ODF adapters, read and write (plan step 23).

- Merged `s20-merge` (0274):
  - `File::merge`/`Edit::merge` fill OfficeCLI `{{key}}` placeholders and DOCX `MERGEFIELD`s in DOCX/PPTX/XLSX, keeping formatting across run splits;
  - CLI `litchi merge` with one output per record; Python `File.merge`;
  - 1.03–1.04× the leaf.
  - Open: F1–F10 (other formats, PPTX notes and tables, merge-field pictures, consolidation with find/replace).

- Merged `s23-ods` (0277):
  - ODS reads, views and queries over content cells; flat `.fods` opens read-only;
  - writes cover cells, named styles, merges, sheets, rows, named ranges, comments, `/metadata`, `create` and atomic save;
  - leaf: LibreOffice-produced ODS is now editable without losing untouched bytes;
  - 1.00–1.02× the leaf, and 1.42× for 10k batches (within 2×).
  - Owner decision (2026-10-14): the current ODS shape (`Edit::set_metadata`) stays; the glossary `metadata` family is recorded as follow-up M1.
  - Sheet and row structure edits are refused on LibreOffice-saved files (leaf rule R1).
- Merged `s20-find` (0273):
  - `find=`/`replace=` (with `matchCase`, `wholeWord`) on `set` for DOCX, PPTX, XLSX and RTF; regex refused (record 0010 Q3);
  - `Applied::matched`; new `Verb::Replace`.
  - Miss: DOCX 1.10–1.13× (leaf prefilter follow-up). PPTX ≤1.04×.
  - Merge fixes: `Verb::ALL` is now 17 with `Merge`; the catch-all arm is allowed for no-format builds; the PPTX-only prefilter is gated.
- Merged `s21-rtf` (0275):
  - leaf: story text edits are spliced into the retained source, so tables no longer move to the end; batches are linear;
  - unified reads: tables, header/footer, notes, comments, hyperlinks, pictures, styles, lists;
  - unified writes: story text and picture removal;
  - single inserts are now within budget.
  - Misses left: single layout set 1.15×, edit-closure singles 1.09–1.14×, single-target fan-out 1.33–1.53×, repeated direct removes.
  - Merge fix: RTF find refuses the new element kinds.
- Merged `s23-odt` (0276):
  - leaf `body` index with lossless typed body edits;
  - a `litchi_odt::metadata` family registered in the gate;
  - unified get/query/view plus set/add/remove/move/swap/create/save;
  - open +1 µs, set ±1%, 1k batch 1.16–1.39×.
  - Merge: 21 files conflicted with 0277's parallel dispatch arms. Files whose hunks split a block were rebuilt from trunk plus the branch's diff, and each file keeps all `ods`/`odt` cfg uses of both parents.
- Gate note: the non-encryption capability goldens come only from an all-features-but-`encryption` run; plain `cargo test -p litchi-ops` has no formats (`default = []`).
- Merged `s23-odp` (0278), after its agent merged trunk into the branch:
  - leaf scene index with checked edits on `Package::edit`;
  - unified reads over 11 node kinds; `.fodp` opens read-only;
  - set/add/remove/move/swap, `create`, atomic save and `/metadata`; batches make one leaf commit;
  - open +1.1%, set +0.3%, 1k batch 1.37×, 10k batch 1.30×.
  - Open: G-ODP-1 to G-ODP-8 (character formatting, non-text adds, find/merge, the shared master reader, flat writes).
- All three main ODF formats (ODT/ODS/ODP) now read and write through `File`. Python's `READ_ONLY` set is empty.

Queued next: field/TOC `refresh` (engine 3), DOCX charts (0270a), the other ODF leaves (0268 G1), the performance misses (full queries, DOCX set and cell writes, PPTX add, `Warn` batches) and iWork M6.

## Wave R5 (2026-10-14)
Cut from `80e24f8745`:
- `s20-refresh` (0279): field/TOC `refresh` (M4 engine 3).
- `perf-query` (0280): shared query and transaction misses (full queries, RTF fan-out and edit-closure singles, source freshness check).
- `perf-writes` (0281): per-format write misses:
  - DOCX set, cells and find;
  - PPTX add and metadata;
  - `Warn` batch coalescing;
  - XLSB metadata and XLS.
- `odf-leaves` (0282): the remaining ODF crates (OTH, ODM, ODG, ODC, ODF-formula, ODI, ODB) into `File`.
- `charts-docx` (0283): DOCX charts, chartex types and chart set coverage.
- `smartart` (0284): typed SmartArt CRUD (v1 scope).

- Merged `smartart` (0284):
  - PPTX `smartart` and nested `node` elements with add/get/query/set/remove;
  - leaf `diagram::authoring` (7 layouts, deterministic parts including Litchi's own drawing) and `diagram::edit`;
  - Litchi-made SmartArt is regenerated whole. Other producers' SmartArt takes in-place text edits (drawing patched or dropped); node add/remove there is refused until `layout` rebuilds it.
  - Checked in LibreOffice 26.2.6.3; 1.00–1.02× the native calls.
  - Open: DOCX (0284a), XLSX, node formatting; PowerPoint not verified.
- Merged `perf-writes` (0281):
  - `Warn` atomic batches coalesce again (49× → 1.00×);
  - PPTX text-box add builds its edit once (1.42× → 1.00–1.02×);
  - DOCX `Table::row`/`Row::cell` (cell writes 1.25× → 1.09×), find 1.12× → 1.07–1.09×;
  - `litchi_xlsb::Workbook::edit_metadata` (1.10× → 1.02×).
  - Misses left, with callgrind profiles:
    - DOCX single set 1.19× (F1: main-part name normalization in `litchi-opc`, the shared resolver, commit wrapping);
    - PPTX open + first set 1.31× (F2: promotion re-decompresses members);
    - XLS reads (F3, handed to 0280).
- Merged `odf-leaves` (0282): all 21 claimed formats except iWork now open as `File`.
  - W3 owner split for OTH, ODM, ODG, ODC, ODI, ODB and ODF formula, in seven glossary profiles.
  - `File::{Oth, Odm}` reuse the ODT adapter.
  - `File::{Odg, Odc, Odi, Odb, Odf}` share the new `format::odf_tree` adapter.
  - Open +1.7–2.5% (formula +11.6 µs); single set ≤1.06×.
  - Open: G1–G11. These include:
    - ODG inserts on LibreOffice files;
    - ODC chart writes;
    - ODI geometry;
    - no coalescing on the tree formats;
    - flat files;
    - the metadata family not adopted.
- Merged `s20-refresh` (0279), the last of the three M4 engines:
  - `File::refresh`/`Edit::refresh` and `Verb::Refresh`, with CLI `litchi refresh` and Python `File.refresh`;
  - DOCX TOC (`\\o \\u \\t \\h \\n \\p`) and computable fields; PPTX slide-number and date fields; XLSX cache write-back;
  - no layout engine: cached TOC page numbers are kept by anchor, and `PageNumbers::Require` refuses with `refresh_failed`;
  - no clock is read by default;
  - 1k-heading TOC 1.02×, current TOC 0.89×, 6k XLSX caches 1.005×.
  - Open: F1–F8.
  - Finding for 0273: a DOCX find over `/body` rewrites a TOC whose field spans paragraphs (fix queued).
- Merged `charts-docx` (0283):
  - DOCX charts at `/body/p[N]/chart[M]` (also `/chart[N]`) with add/get/query/set/remove;
  - deterministic chartex parts for waterfall, funnel, treemap, sunburst, histogram, Pareto and box-whisker on PPTX, XLSX and DOCX;
  - chart `set` covers type, data, geometry, axis range, series color and formatting keys, guarded by `Parsed::is_authored`;
  - corpus 3,847 passed, 779 gap, 0 bugs; new delta D10.
  - Miss: XLSX chart `set` 1.12–1.13× at 10 charts and 1.35× at 100 (an extra drawing-inventory read; cache it).
  - LibreOffice draws the extended charts as plain charts, and DOCX extended charts not at all.
- Merged `perf-query` (0280):
  - lock-free `fstat` source check in `litchi-core`; file-backed full queries are 16–23% faster;
  - the walker skips nodes with no wanted children (ODT 6× → 3.9×);
  - PPTX master/layout cache (full queries 26–83% faster);
  - direct RTF picture removals coalesce (64 removals 53–56× → 1.6×);
  - fair leaf baseline routes added.
  - Open: ODS query 35–41×, XLSX memory scans 4.4–4.8×, a 6% XLSX-memory 1k-row regression (layout), RTF single-target fixed costs. Handed to R6.
- Wave R5 complete.
- **Owner decision (2026-10-14):** keep the per-match file-source freshness check in queries (record 0280 G1, raised by 0281 F3). A once-per-walk check is rejected because it would break "the next pull after a source change fails". Per-check cost work (0280's lock-free `fstat`) stays.
- Environment: a session restart wiped `/tmp`, which held the gate scripts and `TMPDIR`. Gate scripts now live in the session scratchpad and create their `TMPDIR`.

Not dispatched: iWork M6. It waits for owner decisions D1/D2, and the iWork crates belong to their owners.

## Wave R6 (2026-10-14)
Cut from `9fe45629ab`:
- `find2` (0285): the find/TOC bug from 0279, consolidating find with merge's substitution, header/footer/notes/comment and table-cell search.
- `perf-ooxml` (0286): OOXML misses:
  - DOCX single set 1.19×;
  - PPTX open + first set 1.31×;
  - XLSX chart set up to 1.35×;
  - DOCX cells (borderline);
  - PPTX scene `Arc` clones.
- `perf-query2` (0287): ODS query 35–41×, node size and XLSX memory scans, RTF single-target costs, ODT, and redundant leaf probes inside the fence (the per-match check stays).
- `odf-gaps` (0288): the ODF `metadata` family, ODG producer inserts, ODC chart writes, ODI geometry, tree-format coalescing.
- `phase-g` (0289): Phase G acceptance — per-format CRUD checklist with a smoke matrix, full-corpus open → no-op → save regression, benchmark matrix.

- Merged `find2` (0285):
  - DOCX find and template merge share `document::substitute`; the 0273 text-node API is removed and banned;
  - the 0279 TOC bug is fixed — field state carries across paragraphs, results are searched as OfficeCLI does, instructions and markers are never touched;
  - DOCX find reaches headers, footers, notes, comments and text boxes; PPTX find and merge reach table cells and notes;
  - DOCX find is 23–31% faster; merge ≤1.05×; PPTX find 1.07× for one match.
  - New delta F5b: `/body` is the main story only.
  - Open: XLSX rich-text cells, `mc:AlternateContent` documents, PPTX masters and layouts.

- Merged `perf-query2` (0287):
  - unified `Node` 408 → 184 bytes (compact DOC/RTF/ODT paths); the walker moves each match once (XLSX query −8%, PPTX large query −16% in instructions);
  - ODS warm query 36× → 21–23×, batch 10k 1.39× → 1.30×;
  - RTF fan-out 1.27–1.48× → 1.22–1.37×;
  - XLS cell reads use 2 leaf source checks instead of 5, keeping the per-match check.
  - The 0280 6% regression is about 1% in work, the rest attributed to layout.

- Merged `perf-ooxml` (0286):
  - PPTX open + first metadata set 1.31× → 1.05× (preserving conversion defers payloads);
  - XLSX chart set 1.35× → 1.01× (chart inventory cache);
  - PPTX warm shape `get` 2.1× faster;
  - DOCX single set 1.21× → 1.12–1.13×, still a miss (shared target parsing and resolution, about 4.6 K instructions; follow-up F1).
  - Behaviour: corrupt parts on a source handle's preserving conversion now surface on first read.
- Merged `phase-g` (0289), with steps 25–27 as generated, golden-tested artifacts:
  - CRUD checklist over 18 formats;
  - corpus fidelity over 813 files, every no-op save byte-identical;
  - benchmark matrix of 1,016 rows from 0249–0284 (91 open misses, each with a cited follow-up).
  - Fixed: DOCX hyperlink/chart and DOC run sets inside `edit`/`batch`, the RTF body add claim, ODF manifest determinism.
- Merged `odf-gaps` (0288):
  - the ADR 0034 metadata family on all ten ODF leaves through one engine;
  - ODG structural edits on LibreOffice drawings;
  - ODC chart/axis/series writes and add; ODI geometry;
  - tree-format batches make one commit (1k 0.98–1.005×).
  - Fixed `meta.xml` edits after a BOM.
  - Miss: ODB single set 1.35× on 1,000 queries.
  - The `language` reporting fix was reverted by the agent over the Markdown golden. Coordinator: the golden is ours to update for intended changes; queued.
- After R6, the checklist has 7 claimed cells still failing (F1–F4, F6, F7; F5 now passes) and 70 unreached for lack of fixtures.
- Wave R6 complete.

Coordinator scripts and agent scratch now live in `/home/zhuhe/code/litchi-worktrees/.coord/` (outside the repository), so they survive `/tmp` wipes.

## Wave R7 (2026-10-14)
Cut from `99cc3138eb`. It targets what Phase G (0289) found still open for the Definition of Done:
- `checklist-fixes` (0290): checklist failures F1–F4, F6, F7, and metadata `language` reporting (the facade Markdown golden may change).
- `perf-path` (0291): shared target/path/plan costs:
  - DOCX single set 1.13×;
  - an RTF plan cache;
  - the walker fast path;
  - selector copies;
  - the ODB tree re-parse;
  - nested XLS checks (the per-match check stays).
- `odf-verbs` (0292): ODF object-format move/swap/add/remove; flat ODF open and writes.
- `producer-structure` (0293): structural edits on producer files — XLS rows/columns, PPTX slide remove, RTF paragraph lifecycle — and the three editor-vs-reader admission gaps.
- `acceptance2` (0294): Python mirror of the Phase G scenarios, fixtures for the 70 unreached cells, bench matrix extended to 0285+.
- `m4-ext` (0295): find/merge/refresh on more formats, CSV import into ODS, DOCX SmartArt (0284a), XLSX rich-text find.

## Owner decisions (2026-10-15)
- **D1 (record 0248), iWork placement: remove the iWork-free gate.** `litchi-ops` may depend on the iWork crates, so the single `File` enum gets `Pages`, `Keynote` and `Numbers` variants directly. There is no facade-level wrapper and no separate adapter crate. The provisional option (a) is withdrawn.
- **D2 (record 0248), iWork ownership: the iWork owner agent is paused.** The coordinator's agents may now modify the iWork crates, including the ADR 0034 W1 iWork renames and the `IWORK_API_REQUESTS.md` leaf APIs.

## Incident (2026-10-15): coordinator session lost
The coordinator's Claude process stalled and exited while seven agents (wave R7 plus `iwork-found`, 0296) were building.
- Evidence: no API error in the transcript and no reboot; no orphaned build processes. `/proc/pressure/memory` showed sustained full stalls, the host has 123 GB and no swap, and each agent ran cargo with the default 32 jobs. The kernel log is not readable, so an OOM kill is likely but unconfirmed.
- Mitigation:
  - every agent and coordinator build now runs with `CARGO_BUILD_JOBS=6–8` and `RUST_TEST_THREADS=8`, one cargo-heavy command at a time;
  - at most four agents run concurrently;
  - the coordinator watches `MemAvailable` (alert below 16 GB) and free disk.
- Resumed: `iwork-found` (0296), `checklist-fixes` (0290), `perf-path` (0291), `producer-structure` (0293). Held: `odf-verbs` (0292), `acceptance2` (0294), `m4-ext` (0295).

## Incident (2026-10-16): coordinator session lost to an OOM kill
The session exited at 07:46 while four agents (0305, 0308, 0309, 0310) were running.
- Cause, now confirmed: the user journal shows a kernel OOM kill in the session's tmux scope (memory peak 106.8 GB). That scope has `OOMPolicy=stop`, so systemd stopped the whole scope, coordinator included. The policy cannot be changed at runtime.
- Mitigation: a coordinator-side cargo shim (`.coord/bin/cargo`) runs every cargo command in its own `systemd-run --user --scope` with `OOMPolicy=continue` and `MemoryMax=26G`. An OOM now kills only that command. Every agent env, the gate and the rebless scripts put the shim first on `PATH`, and the agent rules require the same scoping for valgrind and perf.
- Resumed all four agents.

- Merged `office-gaps` (0308):
  - XLS producer row/column shifts byte-exact 176/165 → 227/227 of 305 sheets. The ClientAnchor "moves with cells" flags were inverted in the shift code and are fixed. Reserved Formula bits are ignored, so 4 worksheets load again (89 XLS query digests re-blessed).
  - RTF producer splice: paragraph insertion (add+remove 7 → 121), group split, replayed inherited formatting, table rows, removal with anchored content (60 → 110 of 147).
  - ODP G-ODP-1 run/paragraph character formatting through new automatic styles.
  - DOC `FloatingPictures.doc` gives a typed smart-tag refusal instead of `corrupt_file`; refused ODG text sets report `unsupported_capability`.
  - Open (to 0311): DOC unified is 1.56× the leaf (it predates 0308); ODP scan +3–7% wall time with −0.14% instructions.
  - Open (follow-ups): the public `AnchorBehavior` names are still inverted; XLS CF12, conditional-format/validation formulas with cell references, Array/Table records; RTF section breaks; ODP shape/cell/notes formatting.
- Merged `perf-lazy` (0305):
  - Keynote/Numbers trees are built lazily: Keynote open+get 1.10 → 1.03, Numbers 1.21 → 1.06–1.09 (within the +20 µs allowance).
  - DOCX single set 1.09 → 1.055 (`Docx::set_file`, direct `/body/p[N]` text route).
  - DOCX 1,000-paragraph set 2.5× → 1.56× and add 2.1× → 1.46× through a source-backed snapshot memo (the eager owner is at 1.00×).
  - RTF `&Query` route: a plan-cache hit costs ~0.05 µs, so no change.
  - Open (to 0311): litchi-opc audits the replacement part a second time after the commit's own audit.
- Merged `iwork-elements` (0309):
  - Numbers rows and columns (`table::topology`, request 6) through unified `add`/`remove` of `row[N]`/`col[N]`; 28 corpus tables round trip. Tables with formulas, merges, filters or pivots are refused, typed.
  - Character styling on one shared wire module: Keynote runs (request 10), Pages paragraph `font.*` (request 13), Numbers cell `font.*` (request 7).
  - Pages and Keynote table cells write `text`/`value`; Keynote media removal and placeholder hide/show; formulas on created Numbers workbooks (request 18).
  - Performance: Numbers within 1%; Keynote run/placeholder/media and Pages paragraph within 3%.
  - Open (to 0312): Keynote `table_cell` 1.99× and Pages 1.065× (+103 µs), both from the first-read tree build; requests 11 and 19; table creation; formula-coordinate shift for row/column edits.
- Merged `perf-opc` (0311):
  - litchi-opc audited the same DOCX part bytes twice per publish and discarded the commit's own proof. New `PreparedTopology::apply_verified_xml_bytes` takes the audit-minted `xml_minifier::audit::VerifiedSource` (private fields, unforgeable); a passing original audit is remembered per part and revision.
  - 1,000-paragraph DOCX on the default owner: set 1.57× → 0.54× the leaf, add 1.47× → 0.40–0.47×; second set on a committed handle 1,907 → 417 µs.
  - DOC: `litchi_doc::Package::edit` re-parsed the whole file per edit and now keeps its body snapshot. Unified is 1.03–1.04× the `Package::edit` leaf it wraps, and 1.12–1.27× the bare body-text leaf (the commit's `Document` validation).
  - The ODP scan (0308) and ODT plain set (0290) wall-time gaps are binary code layout: isolated benches show equal or faster times with fewer instructions. No code change.
  - Open: DOCX add vs the eager owner 1.05–1.22× (needs a deferred overlay audit or an `xml-minifier` insertion proof).
- Merged `iwork-tables` (0312):
  - Keynote slide tables and Pages body tables build lazily, and direct cell sets skip the table read: Keynote `table_cell` 1.99 → 1.02, Pages 1.065 → 1.02 (both met). A test proves that the direct and general routes give identical bytes or errors on 300 corpus cells.
  - Pages `section::lifecycle` (request 11): section add/remove; 31/31 corpus documents round trip.
  - Keynote title/body placeholder runs take character styles; placeholder `hidden` is readable.
  - Default-feature clippy fixed in `litchi-numbers` and `litchi-iwa-archive`.
  - Merge fix: one rustfmt hunk in `tests/iwork_unified.rs` (trunk gate fmt failure); the rebless run that followed was clean.
  - Open (to 0314): Numbers comments (request 19, diagnosed), the formula-coordinate shift, table creation, notes runs; batch misses Pages `batch_10000` 1.53×, Numbers `batch_1000` 1.32×, Pages `batch_1000` 1.08×, Keynote `batch_1000` 1.067×.
- Merged `docx-insert` (0313):
  - `xml-minifier` successor proofs: `VerifiedSource::verify_indexed` records content boundaries to depth 2, and `verify_successor` proves an edit between two boundaries of one parent by scanning only the new bytes; anything else gets the full audit. Every verdict is compared with the full audit (adversarial cases plus 3,000 random edits), and debug builds re-check every splice proof.
  - litchi-opc keeps the remembered original's indexed proof (at most 4 per package; managed packages remember nothing).
  - 1,000-paragraph DOCX add on the default owner: 442 → 74 µs (0.17× the eager owner, 0.07× the leaf); committed-handle add 454 → 270 µs; set legs unchanged.
  - Open (to 0317): chained far-apart edits on committed handles, depth-3 (PPTX shape) boundaries, the set's one original scan in the litchi-docx commit.
- Merged `miss-sweep` (0310):
  - The 102 open bench misses are triaged into twelve root-cause groups (G1–G12), each re-measured; 98 remain open, each with a measured cause.
  - XLSB queries skip nested freshness checks inside a fenced step and read the version once (file `cell[value>1000]` 2,117 → 414 µs); XLSX header rows no longer re-enter the fenced grid read (50,000 rows 32.8 → 11.7 ms); XLS positional row/column writes don't build the whole grid (1.11 → 1.03). Tests count source reads; outputs byte-identical.
  - Bench matrix parsers through 0313; the `unified-query` lockfile gains `memchr`.
  - Coordinator note: a leftover agent wait loop (`pgrep -f` matching its own command line) was killed; the agent rules now forbid that pattern.
  - Open (to 0315): walker and owned-node residual (G4–G6, about 50 misses), DOCX `Children::next` freshness (G7), XLS sheet-step cache (G8), RTF typed props (G9); Keynote bytes-open copy (G11).
- Merged `iwork-batch` (0314):
  - A corpus property test (`iwork_batch_equivalence.rs`) compares `File::edit`, atomic `File::batch` and chained `File::set`. It found real violations in all three formats, now fixed (about 8,000 sequences pass):
    - Pages, a leaf data-loss bug: several paragraph sets combined into one splice dropped the paragraph-style entries between them. The commit now combines only what it can prove equal (new `litchi-iwa-text-wire` `StorageBoundaries`).
    - Keynote and Numbers runs publish first wherever combining would differ. A failed Numbers run names the first failing item.
  - Batch misses remain: Pages `batch_10000` 1.47 and `batch_1000` 1.08; Keynote `batch_1000` 1.07 (meets the bar with paths parsed outside the clock); Numbers 1.31. Most of what remains is path-text parsing (about 0.96 K instructions per item), handed to 0315. The Pages leaf at 10,000 items is 3.5% slower, the cost of the correctness check.
  - Open (to 0318): Numbers comments (request 19), the formula-coordinate shift, Keynote/Pages table creation.
- Merged `opc-proofs` (0317):
  - The DOCX default-owner set takes litchi-opc's remembered proof: 409 → 23 µs (0.031× the leaf).
  - Successor proofs carry derived boundary indexes, and the overlay keeps each revision's proof, so chained edits on a committed handle are proved edit by edit: set 416 → 25 µs, add 268 → 82 µs.
  - Depth-3 indexes for PresentationML parts; PPTX set+write and add+write are 5% faster.
  - Debug builds re-check every splice and every derived index against a full audit; randomized chain tests run at depths 2 and 3.
  - Open: default-owner DOCX add +4.8% instructions (+1.0–1.6% wall); PPTX committed handles not chained; the first DOCX commit scans the original twice.
- Merged `xls-odp-gaps` (0316):
  - XLS `AnchorBehavior` is now `MoveAndResize`/`MoveWithoutResize`/`Absolute` (MS-XLS, POI). The duplicate writer enum is deleted and the old names are banned; written bytes are unchanged.
  - XLS shifts handle CF/CF12/CFEx/DV formulas, 3-D references, Array/Table records and position-locked anchors over uniform rows. Byte-exact row/column round trips 227 → 254/253 of 305 sheets.
  - XLS reader: WINDOW2 zoom, stale shared-formula flags and the AutoFilter "applied" flag are relaxed with typed defects. Embedded chart caches are no longer read as cells (94 XLS query digests re-blessed; fidelity structure byte-exact 80 → 92).
  - ODP: whole-shape/cell/notes formatting, link and field run formatting, reads through common styles.
  - Owner decision pending (not raised as blocking): the CONDFMT rule count > 3 and ShrFmla-order refusals are spec-strict and stay unless the owner relaxes them.
  - Open: 11 sheets the token reader can't decode, form-control/AutoFilter objects, anchors over mixed row heights, ODP graphic-style inheritance.

## Coordinator audit (2026-10-16): P0 surface against the delta table
Checked against design §13 and the generated `officecli-capability-delta.md`: several **P0** items are still element-level gaps on tier-1 formats. The definition of done requires docx/xlsx/pptx to cover the P0 surface, so these come before further performance work:
- DOCX: `styles`/`style` and `numbering`/`abstractNum`/`num`/`level` (L-DOCX-R6) → record 0319, together with P1 headers/footers and footnotes/endnotes (L-DOCX-R5).
- XLSX: `comment` (L-XLSX-R5), `hyperlink`, `picture`/`shape` (L-XLSX-R7) → record 0320, together with P1 conditional formatting, validation, autofilter and breaks.
- PPTX gaps are P1/P2 only (theme, media, equation, animation, zoom, model3d, hyperlink as an element).
- XLSB, XLS and ODS lag XLSX element coverage by about 30 families each; they are queued after 0320, reusing its element design.

- Merged `iwork-numbers2` (0318):
  - Numbers row/column edits on tables with formulas: a wire-level port of the host editor's shift (`table_topology/shift.rs`), extended to cross-table and cross-sheet readers, packed edges and whole-row/column sentinels. It is byte-equal with the host on 48 formula cases, 318 tables without formulas and 53 packed-edge cases, and found a host-editor bug on shared formula entries.
  - Numbers cell comments (request 19): per-table census, lists without FieldInfo, effective locators, a multi-member writer; all 35 corpus tables round trip. `File` reads, sets, adds and removes the cell `comment`.
  - New routes `row_insert_formulas` 1.011 and `comment_set` 1.002 meet the bar; `open_get` 1.063–1.069 (+6 µs, within the 20 µs arm).
  - Open: `#REF!` on deleted references, splitting shared formulas, cell-writer clears (21), comment edge cases (new member, `NeedSparse`, more than one author); Keynote/Pages table creation and notes run styles not started (plan in record §5).
- Merged `xlsx-p0p1` (0320) — **XLSX P0 surface closed**:
  - litchi-xlsx: five ADR 0034 edit families (conditional formatting, data validation, auto-filter, legacy notes, drawings) on one staged part engine with exact, reversible patches; note splices with VML shapes; fixed an empty `<cfRule/>` parse bug.
  - Unified nodes under a sheet: `comment[N]`, `hyperlink[N]`, `picture[N]`, `shape[N]`, `cf[N]` (12 rule kinds), `dataValidation[N]`, `autofilter`, `rowbreak[N]`, `colbreak[N]`.
  - OfficeCLI excel corpus: passed 1484 → 1552, gap 124 → 51, bug 0. Delta: XLSX elements supported 11 → 33 (gap 27 → 5: ole, pivottable, run, slicer, sparkline). CRUD checklist XLSX 92 → 191 cells, all ok.
  - Performance: every write within 1.011× the leaf; `comment_get` 1.10 (+4 µs, sort and clone of the sheet's notes).
  - Open (part to 0322): x14 CF styling, drawing effects, threaded comments, note runs, filter criteria, the prefixed-root validation/hyperlink writer conflict.
- Merged `docx-p0p1` (0319) — **DOCX P0 surface closed**:
  - Styles and numbering (`abstractNum`/`num`/`level`): read, set, add, remove; paragraph and run `style` set by id or display name (Δ4).
  - Headers and footers by type, footnotes and endnotes (anchored add, set, remove), tabs, breaks: editable. Bookmarks, fields, TOC, watermark and text boxes: read.
  - litchi-docx: style, numbering, header/footer and note catalogs with edit verbs; source-backed `bookmarks`/`text_boxes`. New fixture `textWatermark.docx` (18 KB, LibreOffice, provenance recorded).
  - Word corpus gap 248 → 221; DOCX delta elements supported 14 → 31, gap 21 → 4 (formfield, ole, revision, shape). New corpus deltas D1b (positional `num` path) and D11.
  - Writes within budget (max `footnote_add` 1.09). Reads `style_get_format` 1.22 (+1.1 µs) and `paragraph_list_format` 1.13 (+0.25 µs) miss.
  - Open: bookmark/field/TOC writes, paragraphs inside headers/footers/notes, `levelN.*` shortcut keys, the two read misses.
- Merged `pptx-p1` (0321):
  - `/theme` (colour/font scheme read and set, exact no-op on equal values); master/layout `name` and direct-shape `text` sets; table `col[C]` get/add/set width/remove; shape and run hyperlinks (URL, `slide[N]`, named actions; unsafe schemes refused); media get/remove/add with a poster; equation `formula` read.
  - New elements are reached by path or named query only, so views and `*` queries keep their shape.
  - Delta PPTX elements 25/8 → 30/3 (animation, zoom, model3d deferred as P2); corpus pptx gap 416 → 410, passed unchanged.
  - Writes 1.00–1.094× (column add on a one-slide deck); equation get 1.30× (+14 µs) is a recorded miss (F6).
  - Open: F1 `link`/`tooltip` as shape/run properties; F2 equation writes; F3 media timing; F4 master/layout other props; F5 colWidths, defaultFont, mediaType.
- Merged `xlsb-features` (0322):
  - XLSB gains the nine 0320 worksheet feature nodes, edited at the binary record level so unedited records keep their bytes. Chart, series and axis are read-only.
  - Delta XLSB elements supported 8 → 33 (the same 5 gaps as XLSX); property reads 78 → 295, writes 8 → 165; CRUD checklist XLSB 58 → 175 cells, all ok.
  - XLSX: the prefixed-root validation/hyperlink conflict is fixed with a regression test; `criteriaN.op`; threaded comments and people readable in the leaf.
  - Writes within 1.3% of the leaf; `comment_get` 1.21 (+1.3 µs) misses.
  - Open (part to 0325): XLSX note fonts/runs, XLSB chart writes, x14 CF/DV authoring, a shared spreadsheet drawing module, cross-sheet DV formulas, a unified threaded-comment node.
- Coordinator (2026-10-17): the query-walker agent (0315) hung for about 7 hours with no tool process; it was stopped and resumed from its transcript. Three 14-hour `pgrep -f` self-matching loops from 0310 were killed. `disk-guard.sh` now also reports any active agent whose transcript is idle for more than 60 minutes (`active-agents.txt`).
- Merged `doc-gaps` (0324):
  - DOC answers the DOCX paths and keys for styles, numbering, hyperlinks, bookmarks, fields, TOC, sections, tab stops, text boxes and shapes. Delta DOC elements 15 → 29, reads 39 → 156; CRUD checklist DOC 86 → 188 cells, all ok.
  - `doc::Edit` gains paragraph/run styles, tab stops, `HYPERLINK` insert/remove, field instruction/result targets, bookmarks and section page setup; unedited bytes kept, patches invert exactly. LibreOffice 26.2 reads the edited files as intended.
  - Writes within 1.02–1.07× of the leaf. Point reads miss by a fixed 0.3–0.6 µs (to 0326). Paragraph `format` is 5–9% slower because it returns seven more keys.
  - Deferred: style/numbering/field/TOC writes, revisions (P2), form fields, equations, pictures/OLE, watermark; `sdt` is DOCX-only.
- Merged `query-walker` (0315):
  - Single-subject queries take a fast walker step (XLSX bytes 5–10% faster, ODS warm query −21%, ODT `query_all` 3.0–3.3× → 2.4–2.9×). The general step is kept byte-for-byte, to avoid a measured layout penalty.
  - DOCX children share the fenced step's source check: `saut_page` 1.50× → 1.17×, `drawing` now meets. A per-handle XLS sheet-step cache cuts 26.7% of an eager `get` + text; writes invalidate it.
  - Fair XLSB/ODS leaf legs and RTF typed-length routes added. Open misses stay at 98 (six closed, six new rows).
  - Found: PPTX `*_all_shape` queries 15–35% slower on trunk since 0319–0321 (to 0326). The path parser and typed coercion are not started (to 0326).
  - Coordinator: the agent hung for about 7 hours mid-task and was resumed (see above).
- Merged `ppt-gaps` (0323):
  - PPT elements supported 15 → 26 (against the PPTX vocabulary). Comments, shape and text hyperlinks, table cell text and run formatting are writable; line breaks, theme, charts, OLE and media frames are readable.
  - New leaf mechanism `slide_order::records`: whole-record replacements that keep every other byte, with exact, inverse and durable patches. Table row/column add/remove is refused, typed (F1). Zoom, 3D models and modern comments are PPTX-only (wontfix).
  - Writes 1.009–1.054× of the leaf; comment get 1.44× (+0.6 µs) and run get 3.1× (+1.5 µs) miss (to 0326).
  - Merge fix: the appended glossary sections and test fixture/recipe lists conflicted with 0324; both sides kept, fidelity re-blessed.
- Merged `ooxml-followups` (0325):
  - DOCX stories: `/header|/footer|/footnote|/endnote[N]/p[M]` and their runs get/set/add/remove through the leaf `Edit::edit_story`.
  - DOCX marks: bookmark add/rename/remove; field add (OfficeCLI types or raw), instruction and result set; TOC add/switch set/remove. `level<N>.*` keys on `abstractNum` add; num "Mode A".
  - PPTX `link`/`tooltip` on shapes, placeholders, runs, pictures and groups. XLSX note `font.*` and `runs=` on add and set.
  - New crate `litchi-spreadsheet-drawing` (`objects`, `vml`), shared by XLSX and XLSB with byte-identical output.
  - Word corpus passed 645 → 661, gap 221 → 211, bug 0; DOCX property writes 227 → 250.
  - Writes within 1.07×; XLSX note set 1.48× (an existing extra parse from 0320; a per-handle notes cache is the follow-up); a rich-note read 1.10–1.12×.
  - Open: story paragraphs in views and unscoped queries, story tables/list styles, remaining mark edits, XLSB note formatting.
- Merged `perf-r13` (0326):
  - The PPTX query regression came from 0321: the hyperlink item grew every query node 168 → 176 bytes and added one mispredict per match. The new items are boxed and the two lazy frame cells merged, so the node is 144 bytes (a test guards it), and the large in-memory query is back to its pre-0321 time (1,676–1,709 µs).
  - The path fast path in `litchi-core::selector` now takes `[last()]`, `[k]`, `[@k]`, `[k=v]` and `[@k=v]`, differentially tested on 211,682 inputs. Attribute paths parse about 2× faster; `/body/p[N]` batches are unchanged in wall time (the parsed path is 616 bytes; a smaller predicate representation is a public-API follow-up).
  - Typed lengths skip the text round trip: RTF typed routes 1.30 → 1.22, the same as the text route.
  - Reads: PPTX equation 1.30 → 0.77 and XLSX note 1.10 → 1.02 now meet; XLSB note 1.22 → 1.15; DOCX unchanged; DOC is a ratio against a 1 ns struct copy; PPT left to its owner.
  - Open misses 98 → 101 (three newly tracked reads, still missing).

## Coordinator audit (2026-10-17): definition of done
A read-only audit of trunk `593d1b9fcd` against the goal's definition of done (`.coord/r9/dod-audit-2026-10-17.md`, outside the repository). Verdict: not met yet.
- Single entry point: mostly met. Bug: the facade feature lists omit odg/odc/odi/odb/odf-formula.
- 21 formats read and write, gaps listed: partly met. There is no whole-format refusal, but 54 delta `missing` element rows and thousands of property rows have no follow-up.
- P0: partly met. Open: PPTX paragraph/run/row/cell add and remove, comment remove; DOCX picture set/remove, comment set, SmartArt set/remove; XLSX SmartArt not recorded.
- Registry: 11 element rows contradict their property flags.
- Python/CLI: `dump` and the raw layer are missing.
- Acceptance: no OfficeCLI gap carries a wontfix/deferred disposition; no v1 corpus subset is defined; the checklist lacks the ADR 0008 rows.
- Performance: every miss is explained, 19 without a follow-up; 1k-batch evidence is missing for doc/oth/odm/odc/odi.
- Gates: `gate.sh` now also runs `litchi-core` tests, rustdoc with warnings denied for core/ops/facade, and the bench-matrix `--check`. Facade per-feature checks and fuzz builds come with WP6.

Work packages, in waves of up to four agents:
- **Wave 1:** WP2 registry consistency (0330, running), WP1 dispositions and the v1 subset, WP6 facade/fuzz, WP7 1k-batch evidence.
- **Wave 2:** WP3 PPTX P0 residue, WP4 DOCX P0 residue, WP5 `dump` and the raw layer, WP8 miss follow-ups.
- **Wave 3:** WP9 ADR 0008 checklist rows, WP12 SmartArt decisions. WP10/WP11 (XLS/ODS parity) are running as 0327.
- The owner decisions are collected in [0331](0331-owner-decision-memo.md).

- Merged `registry-consistency` (0330, audit WP2): the 11 contradicting element rows now match their property flags (only the claims had been too narrow). New test `element_operations_agree_with_property_flags` runs in every format and both feature sets. The XLSX batched cell-set path no longer refuses `link`/`tooltip`/`display`, so the corpus Bugs table is empty. A PPTX comment date that is not an XML dateTime is now `invalid_value`. Checklist pptx 279/279, xlsb 180/180 ok.
- Merged `xls-ods-features` (0327): XLS and ODS get the nine worksheet feature nodes with the XLSX vocabulary (XLS record edits in `litchi_xls::sheet_features`; ODS splices through `litchi_ods::advanced::sheet_features`). Delta elements XLS 7 → 29, ODS 8 → 29; checklist XLS 59 → 158, ODS 78 → 166, all ok. Feature read misses: XLS 1.50×, ODS 2.10× (not like-for-like). ODS writes re-open the package per helper (follow-up). Merge fix: litchi-ops `Cargo.toml` features combined (pptx chrono + xls base64).
- Merged `odf-rtf-gaps` (0328): ODT/OTH/ODM headers, footers (text set), numbering, fields, TOC, shapes, equations and columns read, tab stops editable. ODP hyperlinks and line breaks editable; columns, media, equations and chart series/axes read. RTF sections, bookmarks, fields/TOC instruction sets, tab stops, text boxes/shapes; paragraphs move across section breaks (0308 follow-up). Writes within 2%; ODT format reads 9% faster. Point reads +0.2–0.7 µs and child-listing views +3–19% are misses. Merge fix: appended glossary sections kept.
- Merged `dispositions` (0332, audit WP1+WP13): `tools/officecli_dispositions.json` classifies all 7,458 gap rows (P0: 89 planned, 4 out-of-scope; P1: 6,082 deferred). Both report tools render it and `--check` refuses untriaged rows, deferred P0 rows, milestone-only follow-ups and stale keys; both checks are now gate steps. The v1 corpus subset is **red**: 625 of 3,443 lines open (WP3 432, unverified 83, WP5 57, WP16 38, WP4 15). PPTX `diagram` (OfficeCLI's Mermaid renderer) is no longer a false match. Proposed: **WP16**, XLSX P0 residue plus root-level metadata writes. Merge fix: three stale P0 tier keys dropped after 0328 closed their rows.
- Merged `iwork-tables2` (0329): Keynote/Pages tables created by copying the file's first table (a file without one refuses typed: 53/56 corpus decks). Numbers removals write `#REF!`, shrink ranges and split diverging shared formulas (18/18 host cases). Comments on unstored rows and with several authors. Keynote notes run styles. All routes within budget.
- Coordinator: the XLS/ODS agent hung for about 70 minutes and was caught by the new stall check, then resumed.
- Merged `facade-fuzz` (0334, audit WP6): facade gates derived from Cargo features by `crates/litchi/build.rs` (`cfg(litchi_unified)`); odg/odc/odi/odb/odf-formula builds now have `litchi::open`, and unit tests pin the 21 formats. `tools/check_feature_matrix.py` checks the facade per format and per aggregate plus CLI and Python (31/31; 29/29 clippy `--all-targets`, after fixing 10 single-feature failures). New fuzz targets `batch_payload`, `batch_apply` and `smoke_batch_payload` with a 26-seed corpus ran clean for 300 s each, with libFuzzer on stable. Findings F1 (serde_json number drift) and F2 (`PathArg` display not a fixed point) went to 0338. iWork `semantic::{open, from_bytes}` deleted; `sheet::text` kept as non-Office tooling; duplicate glossary ADR 0033 file removed. `gate.sh` now runs the feature matrix, facade clippy over all targets, the fuzz builds and the smoke runs. Merge fix: a boundary-policy unit test still expected the deleted Pages reader and was updated; the agent rules now list the full tool test suite.
- Merged `docx-p0` (0335, audit WP4) — **DOCX P0 residue closed**: picture set (size, alt, name, `src`), remove and query; comment set (replies refused, typed); SmartArt `layout`/`text` set and remove, plus a `node` element with full CRUD; DOCX chart `series`/`axis` nodes. Fixed: a footnote/endnote text set dropped the reference mark. Repeated writes are within budget after per-file caches. The first write on a fresh file is 1.48–1.84× (the unified eager open is 2.7× the leaf open, an existing cost; follow-up F6).
- Merged `pptx-p0` (0333, audit WP3) — **PPTX P0 residue closed**: paragraph/run add/remove in shape, cell and notes bodies; table add; `tr` add/remove/height; merged cells kept consistent where safe; `tc` fill with theme colours and `none`, valign and padding; underline/strike; comment, modern-comment and reply removal (authors kept); picture frame/`alt`/`src` set. Formatting batches coalesce into one leaf commit per slide with a single-pass leaf staging (byte-identical on 166 probe rows): 1k batches 1.13–1.22×, linear to 10k (previously refused at 1k). Corpus PPTX passed 1286 → 1593; v1 open lines 625 → 197 (WP3-blocked 432 → 1). Policy: picture crop/effects and `table.style` moved to deferred-P1 under 0332's breadth rule; coordinator confirmed.
- Merged `perf-closure` (0337, audit WP7+WP8): new `unified-batch-closure` harness gives 1k-batch evidence for all 21 formats (outputs checked byte-equal); 20/21 meet 2× at 1k, and every other format stays within 2× at 10k. **Keynote paragraph-set batches are quadratic: 22× at 1k, 390× at 10k** (to 0340, a definition-of-done blocker). XLS file-backed batch runs drop their per-item source checks: 0209's 2.093× case is now 1.15–1.16×. Every open miss now has a follow-up or an explained residual (26 → 0 without one; `--check` enforces it); open misses 101 → 103. `Path` 616 → about 248 B design note, not implemented (public names would change). Merge fix: bench matrix regenerated after the 0333 + 0337 merge.
- Merged `parity-residue` (0339): the corpus harness records full skipped-property lists (86 unverified v1 lines accounted); 42 element gap rows added, so the delta has 0 `unlisted` element rows; the delta tool's element matching bug (slidemaster → shape) is fixed. XLSX SmartArt deferred-P1 (owner memo item 12), PPT SmartArt and masters given follow-ups; DOCX picture breadth deferred-P1 (item 13). v1 open 197 → 105.
- Merged `xlsx-p0-meta` (0338, WP16): root `set / title=…` writes document properties for DOCX/XLSX/PPTX (DOCX gains core-property writes); XLSX rich-text runs, cell `type` and `add`, picture `set src`, cross-sheet series values. F1: `serde_json` `float_roundtrip` workspace-wide (long-decimal decode +16%). F2: `PathArg` display is a fixed point of `parse`; fuzz tolerances removed. Misses: run set 1.32, picture `set src` 1.27 (follow-up 0338-P1).
- Merged `dump-raw` (0336, WP5): `File::dump` (OfficeCLI v2 batch, replay is a fixed point, typed warnings for anything skipped) for 9 formats; raw layer `parts`/`raw`, `add_part`, `raw_set` (XPath subset, audited publication) and `validate`; CLI and Python surfaces. Corpus missing-verb lines 138 → 0; v1 open 105 → 48. `dump` costs 1.45–2.3× a read-only traversal (no budget). Incident: the agent ran a host-wide `pkill -f 'cargo (doc|test)'`; the agent rules now forbid host-wide kills.
- Merged `acceptance-adr8` (0341, WP9): per-format ADR 0008 rows (atomic failure, patch inverse, composition), 63/63 ok, mirrored in pytest; unreached cells 13 → 0 (generated `video.ppt`, `shapes.key`); Word 2007 encrypted DOCX now opens as encrypted; flat ODP accepts predefined entities; the 37 files that do not open are classified and enforced by a test.
- Merged `keynote-batch` (0340): Keynote paragraph-set batches 21× → 1.09–1.11× at 1k and 380× → 1.72–1.77× at 10k, linear; the ODF object navigator resolves positional steps directly (ODC 10k 1.59 → 1.07). **The 1k-batch budget is now met in all 21 formats.** Open misses 103 → 101.
- Merged `v1-closure` (0342): DOCX app/custom metadata, table style and borders, row height and flags, `listStyle`/`start` on add; PPTX table-level set. **The v1 corpus subset is green: 3,443 lines, 2,274 passed, 1,169 accounted, 0 open.** Row set 1.13 is a miss (follow-up 0342-P1).
- Coordinator (2026-10-19): a second definition-of-done audit (`.coord/audit/dod-audit-2.md`, outside the repository) finds the remaining blockers: DOCX/PPTX P0 residue (26 planned P0 rows), Numbers `swap` (a whole-format verb refusal the first audit missed), five records absent from the bench matrix, Python `Commit.patch`, leaf-crate gates, and the glossary closure (0343). The 392 XLS/ODS P1 parity rows still marked `planned` will be relabelled deferred-P1 as an adopted default, listed in the owner memo.
- Merged `glossary-closure` (0343, WP14): every 🔴 ADR 0034 cell is closed. Five renames, each with the old name deleted and banned: OPC `write_to*`, DOCX `PatchConflict`, XLSB scoped `Edit`, flat ODF `from_bytes(&[u8])`, positional view reads take `Position`. Stale cells flipped; deferrals D1–D11 are owner-memo items 14–24 (D8 withdrawn, done by 0342). The PROGRESS records index now covers 0249–0343; the facade docs, README and FEATURE_MATRIX say every format reads and writes through `File`. The facade iWork fuzz targets compile again; `gate.sh` now builds `crates/litchi/fuzz` and runs clippy and tests for 19 leaf crates plus `litchi-ods --lib`.
- Merged `closure-tools` (0344, second-audit package D): the bench matrix now reads every record with evidence (18 more records; rows 2,377 → 2,737; open misses 101 → 142, each with a follow-up or explained residual), and `--check` fails on evidence without rows or any output of 1 MiB or more (per-format tables moved to `bench-matrix/`; the main file is 402 KB). The 392 XLS/ODS P1 rows are deferred-P1 (P1 planned 426 → 34; P0 still 26, to 0345/0346). 784 milestone-only follow-ups mapped to current ids. Owner memo items 25–29 added.
- Merged `cross-small` (0347, package C): Numbers `swap` of sheets and tables, so **no format refuses a definition-of-done verb outright any more** (rows and columns refused typed, F1). Numbers relocation into a sheet holding the same table name returns a typed `DuplicateTableName`. Python `Commit.patch` (`apply`, `inverse`, `is_noop`), so the ADR 0008 pytest mirror runs 21/21/21. XLSX picture `set src` reads the drawing once: 1.27 → 1.07. `Patch` docs name the content-checked ODF formats; RTF Mac code pages are `unsupported_capability`.
- Merged `docx-p0b` (0345, package A) — **DOCX P0 planned rows 10 → 0** (delta P0 planned 26 → 16, all PPTX): `add cell`, the table `col` element (width read/set/add, add/remove), hyperlink `tooltip`, and comment `parentId`/`done`/`anchoredTo`/`runStart`/`range` through `commentsExtended` (keyed by the last paragraph, per MS-DOCX). New writes 1.01–1.08× the leaf; plain comment add 1.65 → 1.02; column resolution 24 → 1.3 µs. Follow-ups F1–F5 (column move/swap, dump-replay comment controls, tracked `revision.*`, `commentsExtensible`, emptied `commentsExtended` part).
- Merged `pptx-p0b` (0346, package B) — **P0 planned rows 26 → 0 in every format**: PPTX legacy comment author/initials/date/direction; modern comment add and set (IDs derived from the edit, so no clock or randomness; `created` and `author` required on add); notes and placeholder add; slide title/text/layout. Cell, notes and paragraph-text batches coalesce into one commit (0333 F3); two quadratic leaf staging paths are linear (1k cells 1.24 s → 7 ms). CRUD checklist PPTX 308/308; v1 subset still green. **Miss:** eager fresh-handle cell batches are 2.04–2.16× at 1k/10k, over the 2× budget (to 0348).
- Merged `ws-hygiene` (0349) — **determinism fix**: `serde-saphyr` (the `yaml` feature) enables `serde_json/preserve_order`, so durable patch JSON kept insertion order and its bytes and digests depended on Cargo feature unification. The new `litchi_core::canonical_json::{to_writer, to_vec}` always sorts keys and is used by the core patch envelope and every other durable writer (ODT layout/protection, RTF tail-append, DOCX tail-append stream, ODF repair plan, which really differed between builds). Workspace clippy, tests and rustdoc are clean; `gate.sh` now runs all three workspace-wide plus a `litchi-ppt` build without `preserve_order`.
- Merged `pptx-batch` (0348): the PPTX leaf reports per-edit changes from batched staging, and cell/notes plans resolve against one-pass indexes. After review, the indexes validate every body with the reader, so batches stay equivalent to one-at-a-time sets (new corpus property test `pptx_batch_equivalence.rs`; it also caught an empty-notes divergence 0346 had). Notes staging is one pass (leaf 10k 511 → 22 ms). Cells 1.05–1.31×, warm notes 1.19–1.22×, cold notes 1.89–1.97× (follow-up 0348-P1); the bench-matrix TSV is split per format. **Every PPTX batch route is within the 2× budget.**

## Owner decision (2026-10-20): API debt is required work (record 0350)
- Memo items 14–20, 23, 24, 26 and 29 are overridden: immutable `Package` APIs, one Edit/Commit/Patch/Diagnostics contract, one error taxonomy, configurable limits through options, structured metadata, iWork and `litchi-iwa` naming, `sheet::text` module ownership, and static dispatch on the `litchi-core::sheet` traits. All consumers migrate in the same change; correctness, fidelity and performance are preserved.
- Item 22 (the comprehensive redaction audit) is deferred; its rules still apply.
- Every other memo item is confirmed (1–13, 21, 25, 27, 28, 30).
- Plan: an inventory and contract design (record 0351) first, then foundational changes, then per-crate migrations.

## API-debt program (record 0350), merges
- Merged `api-contract` (0351): the normative contract (immutable `Package`, one Edit/Commit/Patch/Diagnostics shape, `Refused(Refusal)` replacing `UnsafeEdit`, limits through `OpenOptions`/`EditOptions`), foundation `litchi_core::edit::{Commit, Diagnostics, Refusal, RefusalKind}`, gate rule kinds `immutable_owners`/`family_methods`/prefix-suffix bans, the debt inventory (642 package mutators, 309 commit types, 444 `*_with_limits`, 29 `rollback`) and waves W1–W8. Coordinator decision: a patch conflict is `invalid_input` in every format; `SourceChanged` stays `io_error`.
- Merged `iwork-names` (0352): iWork `CellPosition`/`CellRange` → `At`/`Area`; `litchi-iwa` `{pages,keynote,numbers}::Builder` and 168 unprefixed per-format names; Keynote native IDs are documented plain API. Follow-ups for W7: 18 deprecated raw-ID host items, root `IWork*` prefixes, the broken iWork public-API tool test.
- Merged `sheet-dispatch` (0353): the `litchi-core::sheet` traits use associated types; boxed iterator traits deleted; per-cell boxing gone from the evaluate path; ABBA 0.36–0.98 on all 15 scenarios. 0353a (concrete sheet error type, synchronous formula engine without boxed futures) is in progress.
- Merged `sheet-text-meta` (0354): `litchi::sheet::text` → crate `litchi-sheet-text` (facade feature `sheet-text`), `File::import` shares its tokenizer (parse 0.88–0.90×); `Metadata.keywords: Vec<String>` in every format, arrays in JSON/CLI/Python, unchanged lists keep source bytes.

- Merged `w1-docx` (0355): `litchi-docx` `Package`, `source_backed::Package` and `Document` immutable; outputs `&self` (repeated save byte-identical, ~70 µs both times); sign/unsign return packages; 23 edit families over one reversible patch; drafts in `docx::builder`; `UnsafeEdit` → `Refused`/`PatchConflict` (`invalid_input`)/`Internal`; `SemanticLimits` on `OpenOptions`, no public `*_with_limits`; DOCX routes within 1.2% of base. Sent back as 0355a (required, not follow-ups): restore the removed alt-chunk/revision-limit/output-limit capabilities, migrate the source-backed streaming families, typed `LimitKind`, classify refusals by kind, and the 1k-op DOCX batch on a 1,000-paragraph body (~300× the leaf, also at base).

- Merged `w3-ppt` (0357): 12 PPT scoped families commit a `Package` via `edit_x`/`apply_x`; `rollback`/`Transaction`/`is_changed`/`*_with_limits` gone; `raw::Reader`/`SourceBackedPackage` `*_with_options`; typed refusals classified by kind; 13 `Presentation` readers now read incrementally saved decks; ABBA 0.95–1.02. Sent back as 0357a: public routes for the text splice, slide-order history/composition/three-way, master reordering and table-style placement; animation no-op; signed refusal only at a changed commit; record-sized patches; the notes-master validator; signed tests.
- Merged `sheet-dispatch` 0353a: synchronous formula engine (no boxed futures; `WEBSERVICE` prefetched between passes through the caller's async `Fetch`); sheet traits carry each implementer's typed `Error`; `DepthLimit` folded into `Limit { kind: Depth }`; all scenarios 2–66% faster than trunk, formula-heavy recalculation 24–26%; residual F3 `cells/xlsx` +1.5% against the 0353 tip (−7% against trunk). `litchi-xlsx`'s local boxed alias goes in W4.
- Merged `w2-pptx` (0356): `UnsafeEdit` gone; immutable `Package`; 33 families publish a `Package`; no public `*_with_limits`; drafts in `pptx::builder`; routes within ±1.1% of base, every 1k batch within 1.97× of the leaf. Accepted: per-part patch conflict checks, encryption mode unchanged by saves. Sent back as 0356a: source-backed stream editors as families on the source-backed owner (the DOCX 0355a shape), the ~55 raw `OpcPackage` mutators, selector-typed family keys, the inventory script.

- Merged `w1-docx` 0355a: 13 source-backed DOCX families commit to a `Package` (`write_to` the only output; `publish_*_to_stream`/`*Publication` banned; the shape is the reference for PPTX 0356a); alt-chunk edits, revision budgets and `max_output_bytes` restored; typed `LimitKind`; refusals classified by kind (OfficeCLI corpus unchanged); 1k DOCX batch 1.03–1.05× the leaf (was 302×); the flaky OPC cancellation test is deterministic.

- Merged `w3-ppt` 0357a: public routes restored on the contract shape (primary-edit durable patches, `join`, three-way planning, master and table-style verbs; `text_edit` splice family); record-sized family patches (6.4× smaller, peak RSS 0.18×); animation no-op and signed rules; no-object decks editable; notes-master validator fixed; POI title-master fixture. 0357b (lazy target digest for the +10.7% commit, durable master/table-style verbs) in progress.
- Merged `w2-pptx` 0356a: source-backed PPTX editing on `source::Package::edit` and `source::svg` (deferred source-backed commits, the 0355a shape); 20 `SourceBacked*` editor types private and banned; no public raw `OpcPackage` mutators (11 new families); selector-typed family keys (`MasterId`, `LayoutId`, `Owner`; a string selector is a slide name); inventory regenerated. 0356b (retire `SourceBackedPresentationEditor`) in progress.

- Merged `w6-odf` (0358): the eleven ODF crates on the contract; immutable owners incl. `flat::{Document, Workbook, Presentation, Drawing, Chart, Image}` with `OpenOptions`; ODT `MutableDocument` crate-private behind `Package::edit` (121 verbs, parity-tested byte for byte); ODS `builder::Workbook`; `Refused`/`PatchConflict`/`Signed`; every route ≤1.027 of base. The purged ODS fixture is replaced by a 345 KiB tracked fixture, so the full `litchi-ods` suite runs again and the coordinator gate no longer excludes it. 0358a in progress (source-backed publishers, durable/merge parity of deleted part patches, `odf-objects` build, string refusals).
- Merged `w3-ppt` 0357b: block-hash commits (changed commit 0.970× the pre-0357a base; first commit on a never-hashed package 1.105×, paid once per package — reported to the owner); durable form and `join` for master and table-style verbs (`Refusal::JoinConflict`).
- Merged `w2-pptx` 0356b: `SourceBackedPresentationEditor` retired (crate-private, banned); `source::cross_slide_copy` family; inventory counts `builder::*` drafts (PPTX owner rows 0). Noted for W8: a cross-copy patch replays only onto the same opened package, unlike DOCX section-layout patches.

- Merged `w6-odf` 0358a: ODS `source_backed::{cells, dde, sheet_metadata}` and ODP `tail_append` families commit deferred owners (`write_to`/`save` on both owners, byte-identical to the old stream routes, time ≤1.032); every capability removed in 0358 has a tested package-level route (record §3); `odf-objects` builds alone; no string-carrying `Unsupported` in ODF. Residuals: ODP tail append +0.5 MiB peak heap (full owner open); ODT set 1.070 and flat ODT open 1.090 shown to be code-generation placement (same instruction counts; 1.014/1.032 with `codegen-units = 1`). Noted for W8: tail-append patches, like PPTX cross-copy, replay only onto shares of their own owners.
- Trunk fix: `tools/perf-baseline` and `tools/native-resave` lock files gain litchi-ppt's `sha2` (record 0357b).

- Merged `w5-legacy` (0359): DOC `raw::Reader`, RTF `raw::Tree`; 9 DOC and 7 XLS package families; `UnsafeEdit` → `Refused`/`PatchConflict`/`Internal` in XLS and ograph (codes unchanged, 20 XLS refusal messages reworded); no `Transaction`/`rollback`/`is_changed`/`_with_limits` in the six crates; three new gate profiles (ole-common, crypto, ograph); inventory 554 → 44 rows; every route ≤1.024 of base. Accepted under 0351 §2.1 rule 6: `raw::Reader`/`raw::Tree` keep `&mut self`. 0359a in progress (`body_text`/`cell_values` publication, source-backed commits, DOFR/saved-selection names, DOC 1k batch with a raised limit).

- Merged `w7-iwork` (0361): every iWork commit has the full triad; six single-operation lifecycle families gained an `Edit`; 106 families registered with `family_methods`; 105 family errors use `Refused(Refusal)` with `kind()` (litchi-ops classifies by kind; final-element removals stay `invalid_argument` as in PPTX); no public `*_with_limits` on leaf or host APIs; `litchi_iwa::Package` immutable, the host editors are `Builder` drafts, no `IWork` prefixes and nothing deprecated (`Borders` alias and a root re-export deleted; 16 native-ID items kept as documented API). Routes within budget (Numbers open+get 1.07, +7 µs; 10k batch ≤1.85). `tools/test_check_iwork_leaf_public_api.py` fixed and added to the coordinator gate. 0361a in progress (`litchi-iwa-common::wire` and `litchi-iwa-structured` limit twins).

- Merged `w5-legacy` 0359a: `doc::body_text` and `xls::cell_values` crate-private behind `Package::edit`, their extra capabilities (compose/join, three-way, history, durable apply, transfer) now `Package` routes; DOC `source_backed::paragraph`, XLS `source_backed::{cells, comments, sheet_visibility}` and RTF `source_backed::tail_append` are deferred families (byte-identical, 0.44–1.00× the old publishers); `parts::{dofr, saved_selection}`; DOC paragraph `bold`/`italic` join the body run, so a 1k DOC batch with a raised limit is 0.91–1.04× the leaf (was 6.0× and 15.7 GB). Noted for W8: XLS/RTF source-backed patches also replay only onto their own owner.

- Merged `w7-iwork` 0361a: the six `litchi-iwa-common::wire` operations take required `WireLimits` (`WireLimits::DEFAULT`; ~1,700 call sites, limits unchanged); `litchi-iwa-structured` conformed (`SemanticLimits` builder, `Error::Limit`/`InvalidLimit`, own glossary profile; unused in the workspace but kept per ADR 0028 for four capabilities only it has). iWork open routes within budget (callgrind ≤ +0.12%).
- 0361a coordinator gate closed at `f2bbbc8d07`: the isolated all-feature workspace rustdoc rerun exits 0 (one build job, fresh caches, 635.98 s, 4.095 GiB peak RSS); the other 51 recorded checks already passed. Receipt in `results/0361a/wsdoc-isolated.json`. W4/W4b/W8 remain required and in progress; no push authorized.
- Coordinator follow-up [0356c](0356c-pptx-manifest-fidelity.md): PPTX raw manifest edits preserve lexical content through independent-copy replay and inverse; private graph sharing and compact patch images avoid retaining whole graphs. Eleven fidelity tests and targeted validation pass. Three isolated ABBA trials cover 16 routes, all ≤1.010× baseline; 128-patch history peak heap falls from 2.98 to 2.93 MB. W4/W4b/W8 integration and the final workspace gate remain open.

- Coordinator follow-up [0359b](0359b-rtf-shared-refusals.md): RTF edit/tail refusals use shared `Refusal` and typed kinds; all native, facade and harness consumers migrate together. Full RTF suite (1,562 tests), facade layout/write tests and the 1,020-file corpus pass. The reviewed golden changes only fixed refusal text in 76 RTF rows. Three isolated ABBA trials cover eight routes, all ≤1.033× baseline; peak heap and allocations are unchanged. W4/W4b/W8 integration and the final workspace gate remain open.

- Coordinator follow-up [0362b](0362b-exact-limit-constructors.md), first slice: the inventory now detects exact `with_limits` names as well as the suffix. `RawWireFields::new` requires its limits and every iWork consumer migrates; native/focused tests, clippy and the exact-name gate pass. Other exact-name constructors remain required closure work, recorded in 0362b.

- Coordinator follow-up [0362b](0362b-exact-limit-constructors.md), Graph/XLS/PPT/ODS slice: owner ingress/history use options; raw chart scans and embedded payloads require limits. The PPT Graph draft moves to `builder::graph::Draft`, with semantic models under `chart::semantic`. All consumers migrate; 149 tests, all-target clippy, the locked harness and five glossary profiles pass. The remaining shared-helper constructors and final integrated gate remain open.

- Coordinator follow-up [0362b](0362b-exact-limit-constructors.md), BIFF slice: raw frame constructors and copies require limits; Graph traversal retains its admitted policy and every XLS/Graph/PPT consumer migrates. Redundant iterator error state is removed (104→72 bytes). All 1,416 native tests, 36 tooling tests, clippy, locked harness, five glossary profiles and the bench matrix check pass. Three isolated ABBA blocks show scan medians 0.986/0.794× and encoding medians 0.993/0.994×; complete samples and reproducer are retained. Shared-helper/typed-error audits and W4/W4b/W8 integration remain required.
- Coordinator follow-up [0359c](0359c-typed-legacy-edit-budgets.md), XLS slice: embedded-payload/publication bounds and cell staging now emit typed budget errors; canonical `LimitKind` replaces `SourceLimitKind` and `StagedChangeLimit` is removed with all consumers migrated. All 1,346 targeted tests, XLS/operations clippy, locked harness compilation and XLS glossary checks pass. Shared OLE/Graph migration and integrated corpus/performance acceptance remain required; this is not a new deferral.
- Coordinator follow-up [0359c](0359c-typed-legacy-edit-budgets.md), Graph slice: finite budget kinds, BIFF normalization and XLS/PPT unified classification migrate together. Reusing admission under unchanged policy removes the initial measured slowdown while preserving tighter-limit/output checks. All 1,446 targeted tests, clippy, the locked harness, four glossary profiles and crate boundaries pass. Three isolated ABBA blocks give native parse/encode 0.937×, edit/commit 1.002× and replay 0.994× median ratios; failed candidates and final samples are retained. Shared OLE typed errors, the Graph owner split and final integrated corpus/performance/gate remain required.

## Coordinator status (2026-10-19): code work for the definition of done complete
The third audit (`.coord/r9/dod-audit-2026-10-19-final.md`, outside the repository) listed three code items: PPTX batch budget (0348), the 34 DOCX chart rows (deferred-P1, owner memo 30) and docs hygiene. All three are done, and the workspace-wide gate now passes on trunk. What remains is the owner's:
- the 30 decisions in [0331](0331-owner-decision-memo.md); items 12, 27, 28 and 29 decide definition-of-done verdicts;
- merging and pushing `feat/unified-ops-api` (it is not pushed);
- the two JSON receipts under `docs/performance/results/change-0890/binaries/` (already on the base branch and origin; the AGENTS.md wording "anything under `binaries/`" covers them).

- Merged `checklist-fixes` (0290):
  - Phase G checklist failures F1 (bare `placeholder` query), F2 (negative `c:axId` read leniently), F4 (PPT `hidden` inserts the slide info atom), F6 (OTH/ODM `meta.xml` created on demand) and F7 (RTF picture remove on mismatched payloads) are fixed;
  - `File::metadata()` reports `dc:language` for every ODF format (six Markdown golden rows updated);
  - F3 (PPTX OLE remove under markup compatibility) remains, with the cross-layer cause recorded.
- Merged `producer-structure` (0293):
  - XLS row/column insert and remove on producer workbooks: 176 row and 165 column round trips are byte-exact;
  - the full XLS reader no longer stops at embedded charts, so 89 XLS query digests now see more cells;
  - PPTX slide removal admits common producer constructs; owned notes, pictures and diagrams stay refused, naming the dependency;
  - RTF paragraph remove/move/swap uses an exact source splice: removal works on 60 of 147 producer files;
  - empty-edit refusals 21 → 19; no regressions.
- Disk incidents (free disk at 22 GB and then 13 GB): agent and gate build directories are pruned aggressively. Most of the disk is held by the performance program's build directory in the main checkout, outside this program.

- Merged `perf-path` (0291):
  - ODB single set 1.35× → 1.00×;
  - compile-time path-lookup index;
  - per-thread RTF plan cache (fan-out 1.13–1.28×);
  - walker fast path and borrowed selectors: ODS warm cell −20%, ODT −15%, XLS get −23% instructions.
  - Miss: DOCX single set 1.11–1.12×. The remaining 7.7 K instructions are in `litchi-core` `Path::parse`, DOCX handle allocation and admission.
  - Fix at merge: the 0296 NUMBERS vocabulary gains the alias index.
- Merged `odf-verbs` (0292):
  - ODG/ODC/ODI/ODB/formula move/swap/add/remove are element-level (new leaf `move_shape_to_page`, `move_{table,column,query}`);
  - `.fodt`/`.fodg` open (46/46 and 3/5);
  - flat ODT/ODS/ODP/ODG edits write back only the changed elements; flat fidelity refusals 71 → 0.
  - Open: G-ODP-1, flat member-adding edits, flat ODC/ODI/ODB/formula.

## M6 — iWork (2026-10-15)
- Merged `iwork-found` (0296, owner decisions D1/D2):
  - the iWork-free gate is removed; the `pages`, `keynote` and `numbers` features carry their leaf edges (ADR 0024 amendment), and `litchi-ops`/`litchi-cli` leave the non-iWork bulk lane (48 → 46);
  - `File::{Pages, Keynote, Numbers}` through the iWork object-tree adapter, and `litchi::open` routes iWork through `File`;
  - writes: Pages paragraph/section text; Keynote title/body text and slide hidden/remove/move; Numbers cell values;
  - all 126 iWork corpus no-op saves are byte-identical; open/get/set are within 1.1% of the leaf.
  - Open:
    - `litchi::iwork::Document` is kept (package directories, tables and text roles not modelled);
    - `create`/structural verbs are typed gaps (`IWORK_API_REQUESTS.md`, requests 15–17);
    - Keynote slide remove has no multi-slide fixture (known failure F8);
    - `litchi-numbers` test-target lints predate this change.

- Merged `iwork-pages` (0299):
  - Pages ADR 0034 W1: focused families, `SourceLimits`/`SemanticLimits`/`OpenOptions`, `Package` owner with a read-only `Document`, `Package::new()` plus a `Builder` with explicit identity, a `pages` glossary profile;
  - leaf paragraph transactions, metadata writes and package-directory open;
  - unified paragraphs add/remove/move/swap/set, sections, headers/footers, footnotes, table read/remove, `/metadata`, `create(Pages)` and Pages directories;
  - 85/85 checklist cells, 32/32 files byte-identical; single ops within 1–2%.
  - Miss: 10k batch 2.64× (the fixed tree build).
  - Open: section add/remove, text runs, table-cell writes.
- Merged `m4-ext` (0295):
  - find/replace and template merge on XLSB, DOC, XLS, PPT, RTF, ODT/OTH/ODM, ODS and ODP through one shared engine;
  - XLSX rich-text cells rewritten run by run;
  - field refresh for DOC, ODT-family and ODP; XLSB/XLS formula cache write-back;
  - CSV import into ODS; DOCX SmartArt insertion;
  - ODF splice writes no longer quadratic.
  - Miss: ODT find with 1 match in 10k paragraphs is 1.12× (F7).
  - The branch's committed corpus-deltas table was stale and was regenerated.
- Merged `iwork-numbers` (0297), after its agent merged trunk:
  - Numbers ADR 0034 W1 (`Package` owner, `Workbook` projection, limits, `Package::new()`);
  - formula text parser;
  - unified create, renames, cell value/formula/number format, merges, clear, sheet/table moves;
  - cell sets coalesce per table.
  - Typed gaps: sheet/table/row/column add/remove (requests 4–6), formulas on created workbooks (18), comments (19), text styles.
  - Single ops 1.00–1.04×, open+get 1.17×, 1k batch 1.44–1.48×.

- Merged `iwork-hygiene` (0300):
  - the release-only Numbers builder failure was a `debug_assert!` around the name-index insert in `litchi-iwa-package`, fixed with release-mode tests;
  - Numbers/Pages merge owner UUIDs are derived, not random (Keynote to follow in 0298);
  - `litchi-iwa` chart examples compile (typed `ChartData` error) and clippy is clean on all targets;
  - the crate-boundary test derives its edges from the JSON;
  - the iWork leaf public-API checker exits 0 (Numbers `Position`/`Range` aliases deleted and banned; shared table modules allowed with reasons).
- Usage-limit pause (weekly limit hit 2026-10-15, reset 2026-10-18): `iwork-keynote` (0298), `iwork-structure` (0301) and `perf-residual` (0302) stopped mid-task and were resumed afterwards.

- Disk (2026-10-16):
  - free disk fell to 4–16 GB repeatedly with three agents rebuilding;
  - with the owner's approval, the main checkout's build directory (`/home/zhuhe/code/litchi/target`, about 965 GB of rebuildable output) was deleted, so free disk is about 1 TB;
  - memory (no swap) still caps concurrency at about four agents.

- Merged `perf-residual` (0302):
  - one-pass plain path parse in `litchi-core`;
  - ODT find (1 match in 10k paragraphs) 1.10× → 1.02×;
  - Pages coalesced sets skip the tree (1k 1.73× → 1.10×, 10k 2.45× → 1.61×);
  - Numbers 1k batch 1.46× → 1.33×.
  - Misses left: DOCX single set ~1.11× (cache pressure), Numbers open+get ~1.25× (tree build). Both go to 0305.
- Merged `iwork-keynote` (0298):
  - Keynote ADR 0034 W1 (`Presentation`, limits, 41 focused families);
  - wire-level structural edits (slide duplicate/add-from-layout, drawables, text boxes, shapes, images);
  - `Package::new`/`Builder`; F8 fixed (stale thumbnail record);
  - unified set/add/remove/move/swap/create with coalesced batches;
  - checklist 113/122 with 0 known failures; 56/56 no-op saves byte-exact.
  - Miss: open+get 1.085× (eager drawable tree; 0305).
- Merged `iwork-structure` (0301):
  - Numbers `sheet::lifecycle`, `table::lifecycle` and `table::order`; the multi-table cell writer fixed;
  - Pages `save_directory` and table-anchored paragraph removal;
  - structural ops 0.98–1.03×.
  - Open: Numbers rows/columns, formulas on created workbooks, styles, comments, Pages sections, runs and cells.
- Merged `acceptance2` (0294):
  - a Python mirror of the checklist (1,696 cells, 21 formats) and of no-op saves, agreeing with Rust;
  - 66 of 70 unreached cells reached (7 generated fixtures plus 1 POI fixture);
  - PPT builder connectors;
  - bench matrix through 0297 (1,420 rows, 101 open misses explained; parsers for 0298, 0301 and 0302 pending).
  - Finding: ODT-family paragraph text sets drop inline frames (fix in 0304).
- With 0296–0301, all 21 claimed formats read and write through the single `File` enum. `litchi::iwork::Document` remains until Keynote package directories open.

- Merged `pptx-gaps` (0303):
  - PPTX slides leave with their exclusively owned parts (notes, comments, pictures, diagrams, VML, OLE), reversibly, with durable patch wire v5;
  - `p14` sections stay consistent on slide remove/move/swap/add;
  - F3 closed (OLE frames under enclosed markup compatibility removed on the source bytes); the Phase G known-failure list is empty;
  - corpus PPTX removals 1/3 → 9/9; new paths ≤1.009×.
  - Open: DOCX run-level markup compatibility (0307).

- Merged `text-loss` (0304):
  - ODT/OTH/ODM text sets keep zero-text elements (frames, notes, bookmarks, marks) at their offset and refuse fields, inner links, rubies and nested blocks;
  - silent drops fixed (now kept or refused): DOCX cell text, PPTX fields, PPT metacharacter fields, ODP inline objects, ODS hyperlinks;
  - a format × content × behaviour audit test for every format;
  - `Leaf::{Pptx, Xlsx}` boxed (combined-feature clippy clean).
  - Open: ODT plain-paragraph sets are 3.4–4.7% slower in wall time with +0.22% instructions (unexplained, under the 5% trigger), and the F1/F3–F6 notes in the record.

- Merged `docx-mc` (0307):
  - DOCX documents with run-level `mc:AlternateContent` are admitted (new `admit_held` proof); block positions are proven, so set/remove/add/move work;
  - each wrapper is one opaque object to find, merge and text sets; a text set over a text box is refused with a typed error;
  - compactor whitespace-reference bug fixed; DOCX fidelity refusals 1 → 0.
  - Finding: unified DOCX text set and add on a 1,000-paragraph document are 2.5× and 2.1× the leaf, handed to 0305.
  - Merge fix: the boundary test scaffold now matches 0298's `pub(crate) mod` (failing on trunk since 0298; the gate now runs the boundary and non-iWork unit tests).
- Merged `iwork-facade` (0306) — **plan step 28 done**:
  - `litchi::iwork::Document` is deleted and every caller (facade, Python, CLI, examples, fuzz, docs) uses `File`; the glossary bans the old names;
  - Keynote and Numbers package directories open and save back as directories through the shared `litchi_iwa_archive::bundle`;
  - `OpenOptions::{pages,keynote,numbers}_options` reach the leaves;
  - bench matrix parsers through 0307 (1,522 rows, 102 open misses).
  - Open: Keynote bytes open 5.3% over a no-copy leaf; loose `Index/` bundles save in the `Index.zip` layout.

Old reference commits (not merged): `uo/m2-w3-executor` 5258f8b42d, `uo/m2-l-pptx-w4` 9474e2b6ce, `uo/m2-l-xlsb-w1` 876588a23c, `uo/m1-r4-perf` 012fd509ff.

- M0a S10 (`docs/unified-ops/s10-measurement.md` §9). M0a landed on the strength of two explanations. If either is rejected, the chosen alternative becomes a harness follow-up; M0a is not reverted.
  - F1: accept E1. Seven rows over their trigger time one call in a fresh child process, so they measure the harness executable. Standalone probes give the same first-call time at both revisions (XLSX open 187.6 vs 187.9 µs); Callgrind shows +0.75% instructions. Alternatives: an untimed warm-up call in the child (new route names), or standalone binaries for first calls.
  - F2: accept E2. The RTF open row depends on the allocator state the untimed projection leaves; with malloc trimming off, the unified open is 0.98× the pre-M0a one. Alternatives: fixed malloc tunables for RTF children, or bare-loop variants.

- Record 0015 decisions D1–D19 and open questions Q1–Q14 (provisional defaults). Highlights:
  - Write verbs take `&self` and return a new `File`; `edit` returns a `Commit`.
  - Batches run sequentially, grouped into same-family runs, and the batch result must equal its items applied one at a time.
  - Writes refuse invalid values and unknown keys, with suggestions. OfficeCLI's apply-the-rest-and-warn behaviour is only a command-channel option.
  - After `Signed::unsign`, a save writes the unsigned copy. Decrypted files save re-encrypted when given the password, or refuse; plaintext needs `into_plaintext()`.
  - XLSX moves to the eager engine on its first write; PPTX text-only runs stay on the source-backed editor.

- Record 0010 open questions (§11.3), the two that most need the owner:
  - Q3: regex values (`~= r"…"`) are refused with `unsupported_capability` until the owner approves adding the `regex` crate.
  - Q6: sheet names match case-insensitively in every workbook format. This changes XLSB and ODS.
  - The other eleven defaults are listed in §11.3.

- Record 0009 (R4): a shared handle shares the leaf package instead of moving it. Write verbs take `&self` and return a new handle; there are no consuming verbs on `File`. Option B (keep the reader and re-open) and option C (consuming verbs with a sole-owner fast path) are rejected.

- Record 0007 defaults: flat `File` (F2), `limit_exceeded`/`cancelled` codes (Q1), iWork → `unsupported_format` until tier 3 (Q2), `retain_source = true` (Q3), `Signed` wraps every signed input (Q4).

- `OpenOptions` shape: one struct with common knobs plus optional per-format sub-options.
- Verb names: `apply`; find/replace as `find=`/`replace=` props on `set`; `merge` = template merge.

## 0004 — log

- Drafted ADR 0033 (Proposed) and listed it in the ADR index; ADR 0024 inventory amended (66 packages / 245 declarations).
- Scaffolded inert `litchi-ops` (edge to `litchi-core` only); registered in `crate_boundaries.json` and `non_iwork_gate.py` (bulk 46 → 47); gate tests updated and green.
- Known: workspace `cargo fmt --all --check` fails on base-branch files (soapberry-zip, iwa-archive) unrelated to this program.
- Follow-up: the first iWork leaf edge from `litchi-ops` (tier 3) requires reclassifying it out of the bulk gate, like the facade.

## 0002 — log

- Code-verified inventory; facade is read-only for every format and drops/never holds the editable leaf owner for doc, ppt, rtf; xlsx/xlsb are held only in the non-editable lazy form.
- No `Locked`/`Signed` types exist. Changed commits silently strip signatures in pptx `opened::Transaction`, several xlsb families, odf-formula, ODS legacy mutators, ODP in-place mutators (ADR 0006 concern; fix in the signature-policy prerequisite).
- doc/ppt `Writer::set_property` never emitted; xls `Writer` writes no property streams. Feature-matrix overclaims: pptx hyperlink write, docx app.xml write.
- Stream order: T1 xlsx > docx > pptx > xlsb; T2 xls > ppt > doc; T3 per record. Cross-cutting prerequisites: facade atomic writer, signature/encryption policy, commit-shape adapter, editable owner retention in M0a.

## 0003 — log

- Plan for M0a: a flat `File` enum with one variant per format, plus `Locked`/`Signed`, each variant one `Arc` handle (16 bytes, Send + Sync). The editable owner is built at the first edit, not at open. Leaf constructors L-PPTX-1, L-XLSX-1, L-XLSB-1, L-DOC-1, L-XLS-1 and L-PPT-1 avoid indexing the container twice.
- Gate: `non_iwork_gate.py` forbids any edge from a bulk package to an iWork crate. So `litchi-ops` stays free of iWork until M6, and the gate must be redesigned first (R1).
- Findings against the design: F1 (§16.2's "mutable owners" for pptx/xlsx are read-only), F2 (flat enum), F14 (Python `FileFormat` drift, including Odb → Doc).

## 0005 — log

- `litchi-core::selector` grammar: paths (1-based → `Position`, `last()`, quoted segments), predicates with `and`/`or`/`not`, header-row predicates, CSS-like queries (`:contains`, `:has`, `:not`, combinators, unions, sheet scopes, ranges), story axes, typed `ParseError` with offset and code, explicit limits.
- 950 OfficeCLI paths and 132 selectors parse and round-trip. Typical paths and simple queries parse with zero allocations (~57 ns per path, ~120 ns per selector, release build, single machine). There are 30 recorded OfficeCLI deltas: element aliases and bare-key meaning move to M1 resolution, and the unscoped-mutation guard moves to the verb layer.
- Open: the `selector_grammar` fuzz target compiles but has not been run (no nightly toolchain or cargo-fuzz on the host).

## 0008 — log

- None of the 19 DOCX fixtures uses a content-changing MCE construct. 13 only declare the namespace (Word 2007 `xmlns:ve`) and 6 carry one root `mc:Ignorable`.
- The source-backed guard refuses them because it compares buffers, and the MCE codec rewrites any part that mentions the namespace URI. The guard also has a hole: a namespace URI written with a character escape slips past both the codec and the guard.
- The eager owner keeps every `mc` construct byte for byte, but it has no MCE or protection check. It accepts a wrong-target edit under body-level `AlternateContent`, accepts `MustUnderstand` documents, and publishes edits to 3 protected fixtures.
- `to_owned_package()` rebuilds the ZIP layout (M0a `SourceMode::Eager` uses it). The execution-context route refuses even reads of all 19 files.
- Reads and edits count paragraphs differently: reads include table cells, edits count only top-level body paragraphs. M1 has to choose for `/body/p[N]`.
- Recommended: a staged admission rule (declared-only, 13/19; then Ignorable-only, 19/19) applied to both owners and the execution-context route, with one protection policy. Run-level `AlternateContent` comes later with a post-commit check. No silent eager fallback.

## 0009 — log

- No leaf guard needs a unique owner. The OOXML publishers take `self` by API choice but only borrow inside, and every share of one opened package carries the same lineage token. The packages cannot be cloned only because their index, catalog and cache are owned inline.
- Probe: the xlsx shared-handle fallback in `Xlsx::into_editor` breaks exact no-ops on all 4 real-producer fixtures (0 of 10–13 ZIP members kept byte-identical) and makes output depend on whether a second handle exists. It must be retired before M2.
- No OOXML source-backed commit can be chained before publishing, so M2 publishes into memory and re-opens at most once per family boundary until overlay chaining (option D) lands.
- Staged leaf plan: S0 litchi-opc cheap `Clone`; S1 xlsx (ops only); S2 docx `&self` publishers; S3 pptx; S4 xlsb non-mutating apply; S5 OLE needs nothing for sharing.

## 0010 — log

- M1 plan:
  - A borrowed node view serialized straight to OfficeCLI's `DocumentNode` JSON, built from static per-element property tables (no per-node map).
  - Leaf-owned evictable indexes.
  - `/body/p[N]` is OfficeCLI's direct-body paragraph, counted on the MCE-visible view. M2 maps it to the edit transaction's numbering under 0011's admission check.
- The default source-backed owners answer little today: DOCX exposes only the main story, PPTX only slide text, XLSX only cells. About 35 core leaf read requests plus 4 correctness fixes are listed in §4.11 (L1 tier 1, L2 tier 2).
- Steps R1–R11. R1–R3 and L1 start before M0a lands, on `uo/m1`; R5 waits for S6, R8 for S9. M1 merges into trunk after M0a. 17 code findings (F-1…F-17) and 21 OfficeCLI deltas.

## M0a — log

- Landed as one merge of `uo/m0a` (`1927763795`). Step detail: `docs/unified-ops/m0a-log.md`.
- `litchi::open`, `open_bytes`, `open_with` and `open_prepared` return the flat, non-exhaustive `File`: 16 bytes, one `Arc` handle per format, plus `Locked` and `Signed`. Detection lives in `litchi_ops::detect`.
- One pointer-sized `Error` with an OfficeCLI-aligned `Code` and a typed `Source`.
- Deleted (S9, −27k lines):
  - the facade's `Document`, `Presentation`, `sheet::Workbook` and sheet openers;
  - the old Markdown adapter;
  - the `__migration` seam.
- The facade, examples, tests, Python bindings and perf harness all use `File` (S7, S8a–d).
- Breaking changes:
  - Changed commits that silently stripped a signature now refuse until `Signed::unsign` (0013, 0014a/b).
  - The eager DOCX owner refuses what the source-backed one refuses (0011); 19/19 Word fixtures are editable source-backed.
- `SourceBackedPackage` clones cheaply (0012).
- S10 gate: passes on the plan's terms.
  - 91/99 facade rows and all 47 kept predecessors pass the raw rule; the other 8 are E1/E2 (above).
  - R5: DOC and the PPT byte route keep about one input size alive after open, and `retain_source(false)` removes it.
  - Follow-ups F3 (drop the whole-child RSS tripwire, which measures the harness's `git` child) and F6 (captures off the shared tmpfs).
- Known failures inherited from the base, unchanged:
  - workspace fmt (iwa-archive, soapberry-zip);
  - 18 facade feature-subset sets;
  - the `keynote_chart_title` fuzz target;
  - rustdoc without features.

## 0015 — log

- M2 plan:
  - Verbs take `&self`: the sugar verbs return a `File`, `edit` returns a `Commit`, and `batch` is atomic or best-effort with per-item results.
  - The registry grows from M1's `PropSpec` and refuses rather than repairs.
  - Batches run sequentially, in runs, with barriers and fusion, and rebase only at family boundaries.
- Critical path: no tier-1 leaf can format existing content today, so the parity engines come first: L-DOCX-W5 (DOCX formatting), L-XLSX-W2 (style allocation) and L-PPTX-W3/W5. They follow 0009 S1–S3 (sharing) and the DOCX run, cell and story addressing.
- Leaf measurements (leaf APIs only):
  - Coalescing is essential: 1k chained XLSX commits took 0.18–12.6 s, against 0.5–7.8 ms in one transaction.
  - The XLSX source-backed cell editor admits 45 of 137 corpus workbooks; the eager route admits 132.
  - PPTX source-backed text edits are 7–18× cheaper than the eager transaction.
  - A DOCX reopen between runs costs 9.4 ms on 10k paragraphs.
  - An XLSX commit rewrites the whole sheet: 53.5 ms for one cell on a 75k-cell sheet.

## Awaiting owner decision

- ~~ADR 0033 acceptance~~ — accepted 2026-10-04 (record 0006).
- **All pending owner decisions are collected in [0331](0331-owner-decision-memo.md).**
- **Adopted defaults, awaiting owner confirmation (2026-10-17):**
  - Record 0319: removing a DOCX `num` keeps the paragraphs' `numPr` references that point at it. OfficeCLI clears them, which edits paragraphs the caller did not name.
  - Record 0319, corpus delta D11: re-adding a built-in DOCX style id is refused (typed) instead of upserted as OfficeCLI does.
  - Record 0316: the XLS CONDFMT rule count > 3 and the out-of-place ShrFmla refusals stay spec-strict.

## 0016 — restart and glossary adoption (2026-10-08)

- Rebased `feat/unified-ops-api` from `dbe66611f1` onto base `32ad64e320`;
  rebased tip before this record is `61630e5635`. Existing M0a/M0b work is
  retained. Historical hashes above are evidence references, not current tips.
- Adopted the owner-supplied glossary/default-policy draft as ADR 0034;
  preserved the existing accepted operations ADR 0033 and its citations.
- Resolved the OPC cheap-clone/fingerprint-memo conflict by sharing the base
  memo through package clones; added regression coverage for clone survival.
- W0 next: rustdoc JSON checker with scoped rules, tests and CI integration.
  Then DOCX → PPTX → XLSX shape waves with simultaneous consumer migrations,
  followed by the remaining format waves and M1/M2 completion. iWork requests
  are recorded for its owners; no iWork implementation changes here.
- No performance, parity, all-format CRUD or milestone-completion claim is
  made by this kickoff. Verification and remaining gates are in record 0016.

## 0017 — W0 glossary enforcement

- Added the rustdoc public-API checker, explicit feature-profile policy, real
  compiler fixtures and CI integration. W0 inspects `litchi-core` with empty
  bans/families; pending API migrations are not claimed conformant.
- Scoped bans cover public names, aliases/methods and public signature type
  names. Registered transaction modules require all six canonical type names.
- 19 tests and the current-tree W0 profile pass. Review findings and rustdoc
  limitations are recorded in 0017; unresolved external globs fail closed.
- Next: DOCX W2 migration with consumer updates and behavior gates. W1 iWork
  ownership and the original all-format CRUD goal remain unchanged.

## 0018 — DOCX canonical entry/error names

- Eager/source `Package::edit()` replaces `edit_document`; main-document
  errors/results now live at `document::{Error, Result}`. Existing edit,
  commit, patch and publication semantics are retained for the ownership step.
- Active consumers, examples, tests and current leaf docs migrated together;
  historical benchmark records keep their original source/API names.
- Minimal/full DOCX glossary profiles enforce removal without public aliases.
- Verification and outstanding W2 ownership work are recorded in 0018.

## 0019 — owned eager DOCX projection

- `Package::document()` returns owned `Document`, retaining a shared physical
  snapshot. It remains usable after package drop and stable across later edits.
- Existing eager mutation paths detach through copy-on-write; clean saves
  preserve the shared graph. Main-part payloads and paragraph indexes remain
  shared, with related-part resolution tied to the retained graph.
- Updated dependent Rust types, scoped style borrows, runnable documentation,
  glossary status and ownership/rollback/no-op regression coverage.
- This completes the eager projection shape, not the immutable package or
  package-owning commit contract. Evidence and remaining work are in 0019.

## 0020 — immutable eager DOCX patch application

- Eager `Package::apply(&patch)` returns a new package and replaces the removed
  mutating `apply_document_patch`. Active consumers and the glossary gate move
  together; the existing validation/publication engine stays private.
- Package clones share the physical graph, metadata and legacy writer state.
  Mutations detach state; failed saves retain edit intent on the affected clone.
- Regression coverage includes immutable apply/inverse/stale rejection, no-op
  sharing, writer/metadata isolation, signatures and encrypted-source policy.
- This is the publication prerequisite for package-bound `Edit`/`Commit`.
  Those shapes, source-backed package publication and writer removal remain
  required; evidence and limits are recorded in 0020.

## 0021 — shared source-backed DOCX publication

- Source-backed `Package` clones share source identity, lazy caches, admission,
  tolerance reports and execution reservations without reopening the archive.
- All 32 package stream publishers borrow their owner. Retained owners support
  repeated publication and retry after sink failure; source-version checks stay
  active. Cache reservations persist until the final retaining owner drops.
- Concurrent clone/no-op and metadata-report tests cover shared state; existing
  managed tests explicitly release the retained source before zero-budget checks.
- This enables retained source publication; package-bound `Edit`/`Commit`, source
  overlays and canonical immutable publication remain required. See record 0021.

## 0022 — eager DOCX package-bound commit

- Canonical `document::Edit` retains its complete eager source. Its commit
  validates and publishes a new `Package` with the patch and diagnostics triad.
- Removed eager `publish_document_edit`/`publish_document_commit`; migrated
  consumers and enforced removal in minimal/full glossary profiles.
- Explicit nominal `document::snapshot` types retain the raw XML engine for
  standalone/source/history callers until their owner migration. Root commits
  cannot carry only a snapshot or an optional owner.
- Tests cover retained related parts, source drop, forked edits, chaining,
  inverse, no-op sharing, later source writer changes and security refusals.
- Source-overlay owners, snapshot composition/history, full package lineage,
  limits and writer removal remain required. Evidence and limits are in 0022.

## 0023 — retained source-backed topology publication

- OPC `PreparedTopology` now cheaply clones its complete prepared candidate
  and reservations. Publication borrows it, enabling retry and repeated output.
- ZIP append plans retain shared generated-member handles; publication does
  not deep-copy prepared relationship XML, member names or appended payloads.
- Added metadata admissions and tests for concurrent clones, retry, source and
  transfer rechecks, candidate readback and final-owner reservation release.
- This is the retained topology foundation, not chained source-package commit.
  Logical overlay revisions and revision-bound XML proofs remain required;
  the original physical source proof is deliberately unchanged. See 0023.

## 0024 — revision-bound source XML edits

- `SourceXmlPart::begin_revision()` establishes a fresh logical edit base over
  the current validated payload, without reopening the ZIP or rescanning XML.
- Range proofs bind to that revision, physical source, part and content type;
  sibling and stale proofs are refused even when their bytes match.
- Physical replacement authority remains the original archive. Managed proofs
  retain payload admissions; sequential revisions do not retain an ancestor chain.
- Chained edits, inverse exact-artifact publication, source fences and managed
  limits are covered. Package-wide effective revisions and canonical source
  commits remain required. See record 0024.

## 0025 — immutable edits of retained effective candidates

- `PreparedTopology::edit_xml` applies a checked edit to effective source XML,
  returning a new flat candidate over the same physical archive.
- Prior additions, omissions, relationships and content types survive; logical
  no-ops preserve the current candidate, and exact inverses restore source bytes.
- Effective parts retain XML provenance. Unchanged metadata and transfer guards
  share ownership; old candidate vectors and intermediate payloads can release.
- Revisions enforce physical/transfer fences, aggregate limits, operation bounds,
  cancellation and managed admission. Swapped callback transactions are refused.
- Added/unproven XML targets, general topology application and canonical source
  DOCX commits remain open. Record 0025 covers evidence and performance limits.

## 0026 — owned effective Part and payload reads

- Retained candidates expose owned lazy Part views and payload handles, including
  unchanged candidates, plus effective iteration and main-document resolution.
- Added/removed Parts, effective relationships and changed payloads route through
  the captured catalog. Views survive owner drop and remain stable across edits.
- New captures recheck base/transfer sources and execution contexts; captured
  metadata and bytes remain immutable readable snapshots. Payloads retain their
  managed admissions and the complete candidate, without a detachable Arc.
- XML proof capture is shared with `edit_xml`. Narrow physical identity, artifact,
  limits and execution accessors avoid requiring a parallel raw package owner.
- Canonical source DOCX integration still needs one effective read/publish adapter
  and logical cache keys. Record 0026 documents those requirements and evidence.

## 0027 — DOCX patch predicate and Debug policy

- Primary main-document `Patch::is_noop()` replaces `changed()` without an alias;
  commit `Diagnostics::changed()` remains. Publishers and harness consumers are
  migrated, and both DOCX glossary profiles enforce the new name.
- Main-document snapshots, patches, operations, transfers, replacement inputs
  and local error/composition wrappers use content-free Debug formatting.
- Regression tests cover compact/pretty output and byte-vector disclosure.
  Display and shared-core diagnostics remain separate open policy work.
- Record 0027 contains verification and the remaining canonical owner scope.

## 0028 — explicit effective graph materialization

- Retained OPC candidates can materialize their effective graph into an owning
  package, preserving edits, additions/removals, relationships and content types.
- Conversion shares retained payloads, checks base/transfer fences and refuses
  managed base or donor escape before payload reads. Signature and ingress
  repair policy survives conversion.
- This supplies DOCX's future effective `to_owned_package` route. Canonical
  source ownership and logical caches remain open; record 0028 also records
  the empty-topology/read-ahead migration hazard.

## 0029 — source DOCX effective-owner boundary

- Same-lineage retained topologies can be adopted into immutable DOCX owners.
  Main-document views, explicit owning conversion and publication use the
  effective graph; logical owners receive fresh semantic caches.
- Pinned effective views retain payload admissions and donor execution guards.
  Queries observe cancellation while captured bytes survive source changes.
- All unmigrated physical operations refuse effective owners, including no-op
  requests. The raw package accessor is fallible; the unified consumer is adapted.
- Canonical source semantic Edit/Commit and specialized effective routes remain
  open. Record 0029 documents scope, verification and current allocation limits.

## 0030 — immutable retained semantic apply

- Source DOCX `Package::apply` returns a new effective owner without ZIP
  serialization/reopening. Snapshots and raw edits read effective main XML.
- Exact source proofs and effective settings policy guard changes; unrelated
  graph edits survive, no-ops preserve signatures, and inverses restore bytes.
- OPC full-document replacement/provenance primitives retain managed admission
  and source fences. Revision direct/replay edits preserve unmanaged identity.
- Canonical package-owning source Edit/Commit, specialized effective routes and
  allocation/performance work remain open. Record 0030 documents verification
  and the exact-byte (rather than strict logical-version) precondition boundary.


## 0031 — canonical source package-owning commit

- Source `Package::edit()` returns `document::Edit<source_backed::Package>`;
  commit returns the new effective owner through the canonical package accessors.
- Generic Edit/Commit share eager/source staging and diagnostics with private
  construction and specialized commit policy. Unified editor consumers migrate.
- Legacy stream/history harnesses explicitly request raw XML transactions.
- Canonical MCE coverage exposed and fixed escaped-namespace compatibility:
  unmanaged full-document apply uses the existing publication audit, while
  managed edits keep typed splice authority and admissions.
- Earlier pending primary-source-triad notes are superseded by this record.
  Specialized families, history, limits, performance and full acceptance remain
  open. See 0031 for scope and verification.


## 0032 — typed primary transaction limits

- `document::LimitKind` and `Error::Limit { kind, observed, maximum }` replace
  free-form resource strings and legacy count field names across primary paths.
- Durable aggregate transfer bytes now fail with a typed limit before payload
  copying; dynamic count/path helpers carry typed dimensions.
- DOCX glossary profiles ban the retired fields at their precise variant scope.
  Public boundary/rollback and durable overflow tests accompany the migration.
- Caller-selected semantic limits, other error families, policy vocabulary and
  full family registration remain open. See 0032 for verification and scope.


## 0033 — metadata on retained source owners

- Source DOCX metadata now follows effective property payloads and relationships,
  including canonical commit owners, additions/redirection and removal.
- A common OOXML prepared-graph adapter reuses existing property validation,
  repairs and decoding. Only the selected payload is decoded.
- Base/donor fences cover success and errors; repair reports stay per-owner and
  deduplicate repeated reads. Managed donor cancellation/release is tested.
- Other specialized reads and metadata writes remain pending. Record 0033
  documents verification, allocation limits and the remaining full-goal scope.


## 0034 — text streaming on retained source owners

- Source DOCX text output follows effective main replacements and relationship
  redirection, including canonical commit owners and managed donor payloads.
- Payload handles stay retained through output and final source checks; declared
  size is checked before capture. Partial-output errors keep exact progress.
- Canonical managed/unmanaged commits, donor cancellation, early limits and MCE
  source-failure precedence have integration coverage.
- Specialized reads/transactions, configurable limits and the wider goal remain
  open. Record 0034 documents verification and allocation boundaries.


## 0035 — shared OPC/OOXML SourceLimits vocabulary

- Renamed the physical profile and builder to SourceLimits/SourceLimitsBuilder
  throughout OPC, OOXML re-exports, unified consumers and active tools, without
  retaining aliases. Defaults, admission and checked builders are unchanged.
- Added minimal/full OPC, PPTX, XLSX and XLSB glossary profiles. Exact-scope bans
  distinguish XLSB's root re-export from unrelated semantic ReadLimits types.
- Broader tests caught an empty-plan publication regression: the one-shot
  path now preserves its exact-copy budget without allocating a retained owner.
- Kind/error names, options envelopes and other limit migrations remain open.
  Record 0035 documents checks and the unrelated existing iWork example error
  encountered by the exploratory workspace all-targets build.


## 0036 — typed physical source-limit errors

- OPC and OOXML expose SourceLimitKind; primary OPC failures use Limit with
  kind/observed/maximum and InvalidLimit with kind/value/maximum.
- Builder failures carry scalar or cross-field bounds. The members/entries
  inconsistency now identifies the excessive member dimension correctly.
- ZIP mapping, cached failures, unified classification and consumer assertions
  retain typed numeric evidence. Glossary profiles ban the retired vocabulary.
- Remaining configuration/dimension and semantic-limit work is documented in
  record 0036; the wider goal remains active.


## 0037 — canonical physical source dimensions

- Shared OPC/OOXML limits use members, entry_bytes, aggregate_bytes, depth and
  events with matching max_* methods and SourceLimitKind variants; no aliases.
- Registered all 24 OPC dimensions and their distinct meanings in the glossary.
  Defaults, admission rules and owner-local semantic profiles are unchanged.
- Migrated consumers, tests, active fuzz/performance tools and added exact-scope
  glossary bans. Record 0037 documents verification and remaining work.
- OpenOptions, semantic profiles and the wider unified goal remain active.


## 0038 — owning DOCX OpenOptions

- Added one crate-root OpenOptions envelope for owning DOCX ingress, canonical
  byte-slice constructors and configured password routes; removed the replaced
  limit-taking and ad-hoc constructors without aliases.
- Migrated unified decrypted adoption and consumers. The same source policy
  reaches clear OPC parsing while encryption mode remains retained.
- Added path/reader/bytes boundary and independent encryption/source limit
  checks; glossary rules ban the retired owning-Package methods.
- Record 0038 records evidence and remaining lazy-owner, semantic and crypto
  policy work. The broader unified goal remains active.


## 0039 — DOCX positional options and resource retention

- Replaced 17 path/positional constructor variants with canonical open and
  from_read_at options routes, using the existing root OpenOptions envelope.
- Cache, read-ahead and execution settings reach OPC admission and deferred
  reads; cloned options share budgets and cancellation. Path cancellation is
  checked before file opening.
- Owning routes explicitly refuse lazy policies before I/O, with typed unified
  classification. Sequential reader accounting and owner consolidation remain
  required; this does not declare the ingress family complete.
- Record 0039 documents migrated consumers, verification and remaining work.


## 0040 — managed sequential DOCX ingress

- Lazy reader and byte-slice ingress use OpenOptions; the old limits-only
  reader constructor is deleted and active consumers are migrated.
- Retained source buffers own their memory reservations through package clones
  and artifacts. Growth admits old/new peaks; reads, retries, copies and input
  counts obey caller execution policies, including the bounded EOF probe.
- Added focused failure/lifetime tests, a managed sequential fuzz target and a
  deterministic 1,024-mode smoke driver. Record 0040 documents results.
- Owner consolidation, semantic/crypto policy migration and the wider unified
  goal remain active; no complete milestone or performance gain is claimed.


## 0041 — lazy DOCX canonical publication and atomic save

- Replaced lazy Package::to_stream with write_to and migrated callers; added
  save (Full durability) and save_with_durability over retained effective graphs.
- Shared OPC atomic staging admits its fixed buffer to the destination owner's
  context and checks caller source state after file sync, before replacement.
  Pre-rename failures preserve the destination; post-rename sync errors retain
  their Committed classification.
- Record 0041 documents durability, preservation, refusal and cleanup tests.
- Owning/specialized output names, budgeted to_bytes, owner consolidation and
  the wider unified goal remain active. No complete family is claimed.


## 0042 — owning DOCX canonical stream output

- Owning Package now exposes write_to; to_stream is removed without an alias.
  Both DOCX owner implementations use the glossary's canonical stream verb.
- Migrated unit/integration tests, operations consumers, active benchmark code
  and the sequential-ingress fuzz smoke fixture producer. Historical evidence
  and low-level OPC names remain unchanged.
- The same serialization path retains encryption refusal, signature policy,
  property staging and retry behavior. Record 0042 lists verification.
- Owner consolidation, immutable owning publication, specialized APIs and the
  broader all-format unified goal remain active.


## 0043 — typed owning DOCX publication failures

- Removed the owning publisher's conversion of OPC failures into formatted I/O
  strings. Error::Opc now preserves typed causes and IncompleteOutput counts.
- Regression tests distinguish zero-byte and partial exact-source failures,
  prove byte-exact retry, and check typed partial-output errors while custom
  property staging rolls back and remains retryable.
- Record 0043 documents verification. Broader error normalization and the
  complete unified goal remain active.


## 0044 — owning DOCX bytes and encryption policy

- Added canonical to_bytes and explicit to_plain_bytes over the existing owning
  publisher. Signed no-op/refusal and pending-edit behavior use the same path.
- Adopted shared OOXML encryption provenance and typed EncryptionPolicy causes
  for ordinary output and missing retained modes; explicit plaintext preserves
  provenance. Unified classification keeps the refusal typed.
- Added feature-isolated, encryption-mode and byte-publication coverage; see
  record 0044. Managed lazy bytes, immutable publication, owner consolidation
  and the full unified-operations goal remain active.


## 0045 — typed DOCX direct property reads

- Added typed run, paragraph, table, row and cell property projections plus the
  positional run set and inline classification needed by the formatting work.
- Adapted the earlier M1 leaf candidate to inherited namespace bindings, strict
  attribute checks, content-free diagnostics and touched-view Debug output.
  Run and cell reads exclude nested owners and revision snapshots.
- Added bounded fuzz coverage and corpus/regression checks. Record 0045 lists
  source provenance, verification and remaining scope. M1/M2 and the complete
  all-format goal remain active; no candidate branch history was merged.


## 0046 — reversible DOCX formatting and body structure

- Adapted the earlier formatting-engine leaf into the current package-bound
  Edit API: typed run/paragraph properties and paragraph/run insert/remove/move.
  Exact patch inverses, durable replay and composition use the same operations.
- Threaded inherited namespaces through existing-fragment readers and splicers;
  added strict replay checks, opaque-comment preservation, namespace-checked
  dependency attributes, bounded generation and content-free Debug output.
- Corpus, owner publication, inverse and policy tests are recorded in 0046.
  This advances the DOCX write kernel; the property registry, managed formatting,
  caller semantic limits, performance acceptance and all-format goal remain open.


## 0047 — node projection and shared value coercion cores

- Added borrowed node views, typed paths, static property specifications and
  streaming JSON projection, with one shared length/color coercion model.
- Adapted earlier source-only candidates to explicit raw identity construction,
  content-free diagnostics, strict depth/non-finite refusal and Word twips.
- Added golden, allocation, minimal-feature and bounded parser fuzz checks;
  record 0047 lists validation and explicit parity/allocation limitations.
- Format adapters remain placeholders. Registry dispatch, File query/write APIs,
  M1/M2 integration and the complete all-format goal remain active.


## 0048 — bounded path resolution core

- Added IntoPath/PathArg, Excel notation, typed grammar errors and a static
  family-vocabulary resolver with relative navigation and get/find semantics.
- Adapted source-only candidate 85300f74e2 to raw-only native identities,
  content-free diagnostics/Debug, typed-input bounds and canonical Unicode
  name matching. Lowercase whole-column notation stays borrowed.
- Record 0048 documents validation, preserved typed causes and allocation/parity
  limitations. Concrete adapters, File navigation, query compilation, registry
  dispatch and the full M1/M2/all-format goal remain open.

### 0049 — DOCX indexed block prerequisite

Record [0049](0049-docx-block-index.md) adds typed direct-block navigation,
namespace-preserving values and generation memos, adapted to immutable eager and
effective source owners. Managed detached views refuse, including donor-backed
projections. The DOCX unified adapter and M1 closure remain pending; the complete
all-format goal remains active.

### 0050 — Query engine prerequisite

Record [0050](0050-query-engine.md) integrates bounded plan compilation and
streaming evaluation, with explicit raw identity admission and content-free
errors/Debug. Test-only Word, presentation and workbook adapters exercise the
engine. Concrete DOCX and File dispatch remain pending; M1 and the full goal
remain active.

### 0051 — DOCX read adapter integration

Record [0051](0051-docx-read-adapter.md) integrates DOCX nodes, navigation,
query binding, projections and internal document views. ADR 0034 adaptations
keep ordinary paths positional and native identities behind raw admission.
Source freshness, invalid-property refusal and feature-specific snapshots are
covered by regressions. File dispatch, M1 acceptance and the full goal remain
active.

### 0052 — File unified read entry points

Record [0052](0052-file-read-dispatch.md) connects DOCX navigation, lazy queries
and text/JSON views to File and the `litchi` facade. Walks now own their adapter,
with retained source checks and borrowed-plan reuse. Security wrappers and raw
identity admission have public dispatch coverage. Other adapters, capability
publication and M1 acceptance remain open. Git metadata is still read-only, so
records 0051–0052 remain uncommitted working-tree changes.

## Capability reports (record 0053)

File/facade capability discovery publishes static DOCX read metadata and
truthful protected/pending-format refusals. Ordinary reports hide native
selection fields and paths. Generated OfficeCLI deltas and CI snapshot-hash
checks expose remaining parity work. Complete gap triage, registry writes,
other adapters and full-goal verification remain pending. Changes in records
0051–0053 remain uncommitted because Git metadata is read-only.

## XLSX leaf read prerequisites (record 0054)

Focused `workbook::source::{Grid, CacheDiagnostics}`, bounded shared retention,
worksheet delegates and early-stop visits are integrated. Cold cell reads
retain streaming; metadata projection enforces caller physical XML limits.
R3/R4 formatting/table prerequisites and the concrete XLSX adapter remain
pending. Source/semantic-limit and owner vocabulary migrations are incomplete.
The full goal remains active; records 0051–0054 are uncommitted due to read-only
Git metadata.


## XLSX formatting and tables (record 0055)

Readable formatting, lineage-bound source style keys, source table access,
checked table bands and ambiguity-refusing Unicode name lookup are integrated.
Caller XML policy and content-free diagnostics passed regression checks. The concrete
XLSX adapter and performance acceptance remain pending. The full goal is active;
records 0051–0055 remain uncommitted because linked-worktree Git metadata is
read-only. See [0055](0055-xlsx-formatting-and-tables.md).


## XLSX unified read adapter (record 0056)

XLSX now participates in File navigation, lazy queries, typed node views and
text/annotated/outline/stats views. Ordinary canonical paths are positional;
native number-format predicates require explicit raw admission. Source fences,
Unicode ambiguity, caller limits and fixed per-walk allocation costs are
covered by regression checks. Operations tests, facade tests, feature builds,
Clippy, rustdoc, workspace compilation, glossary and boundaries pass. M1 acceptance,
remaining adapters and all later milestones remain open. See
[0056](0056-xlsx-unified-read-adapter.md).


## PPTX retained-scene foundation (record 0057)

Eager slides, layouts and masters expose independently retained scenes with
checked positional/name selection, typed ambiguity and content-free diagnostics.
Layout and placeholder token classifiers are integrated. Source-backed scene
caching, related parts and the unified PPTX adapter remain open. See
[0057](0057-owned-pptx-scenes.md); the full goal remains active.


## Source-backed PPTX scenes (record 0058)

Source slide scenes now use the independent owned projection and a bounded
weighted LRU, with managed no-retention, source/execution fences and physical
XML limits before MCE. Scene handles do not retain their presentation owner.
Related parts, richer property/text projection and unified PPTX dispatch remain
open. See [0058](0058-source-pptx-scenes.md); the full goal remains active.


## PPTX lazy shape properties (record 0059)

Non-visual properties and full DrawingML transforms are decoded on demand;
native payload relationships and spans use explicit raw access. Namespace,
ambiguity, retained-source parity and diagnostic redaction checks pass.
Formatting tab stops no longer appear as literal text. Related parts, R7 text
projection and unified PPTX dispatch remain open; the full goal is active.
See [0059](0059-pptx-shape-properties.md).


## Shared DrawingML text bodies (record 0060)

PPTX lazy body reads and SpreadsheetDrawing now share paragraph/run/property
decoding. Strict refusal, raw native identities, retained limits, redacted
diagnostics and atomic authoring guards replace the candidate's lossy legacy
policy. Namespace preparation has a root-inheritance fast path. Related parts,
unified PPTX dispatch and the full goal remain open.
See [0060](0060-shared-drawingml-text-bodies.md).


## Source PPTX related parts (record 0061)

Source slides now acquire detached layout, notes, legacy comments and basic
chart snapshots. Native identities remain under explicit raw access. Reads
apply source/execution fences, bounded preprocessing, strict relationship
validation and rollback of failed scene publication. Full PPTX tests and the
final targeted related-part regressions pass. Presentation metadata, full
chart/modern-comment coverage and unified PPTX dispatch remain pending.
See [0061](0061-source-pptx-related-parts.md); the full goal is active.


## Source PPTX presentation catalog (record 0062)

Source presentations now expose detached masters, layouts and sections in
authored order, plus explicit raw root access. Shared leaf codecs enforce
core namespaces, ID-list structure and schema-valid dimensions. Collection
reads bound accounted retained bytes and roll back failed scene publication.
Full PPTX tests, focused source tests, Clippy, rustdoc, workspace compilation
and both glossary profiles pass. Modern comments, full charts, unified PPTX
dispatch and performance acceptance remain pending; the full goal is active.
See [0062](0062-source-pptx-presentation-catalog.md).


## Source PPTX modern comments (record 0063)

Selected slides now acquire detached modern comments and replies with semantic
text and author names. The source read validates its selected author closure,
preserves inherited body namespaces, and keeps wire identities under explicit
raw access. Shared codecs accept valid XML character references and retain typed
resource refusals. Full PPTX tests, Clippy, rustdoc, workspace compilation and
both glossary profiles pass. Full charts, unified PPTX dispatch and performance
acceptance remain pending; the full goal is active.
See [0063](0063-source-pptx-modern-comments.md).

- Coordinator follow-up [0359d](0359d-graph-package-owner.md): standalone Graph edits and replay now return `Commit<Package, Patch>`; Snapshot publication is removed and every consumer migrates. Edits start in constant time and replay shares validated bytes. All 136 targeted tests, clippy, locked harness, Graph owner/family glossary, 36 tooling tests and matrix checks pass. The initial attachment/heap regressions were repaired: final large attachment ABBA is 0.996–1.003×, replay median 0.081× and retained-history median 0.948×; peak history heap is 437.88 KB versus 439.69 KB. Untouched-CFB-stream sharing, nested raw chart/shared OLE audits and W4/W4b/W8 integration remain required.

- Coordinator follow-up [0359e](0359e-object-capture-budgets.md): shared OLE core capture budgets now use typed kinds/counts; DOC public callbacks and XLS/core/sheet/unified conversions migrate together. Selected-object absolute depth and maximum-depth selector admission are fixed. All 4,141 distinct relevant tests, clippy, locked harness, four glossary profiles, boundaries and tooling pass (five pre-existing ignored tests). Direct budget checks repair the small-open cost; corrected-buffer ABBA and the longer large-replay follow-up are retained, with unchanged 21.30 MB peak heap. Raw shared OLE policies/publication, Graph sharing and W4/W4b/W8 integration remain required.

- Coordinator follow-up [0362c](0362c-odraw-required-limits.md): ODraw image/shape/traversal entry points require limits and every consumer migrates. Delayed resolution enforces both retained byte policies; traversal removes an unnecessary stack allocation. The full initial consumer suite, final focused tests, all-consumer Clippy, harness, glossary and tooling pass. Final traversal ABBA is 0.965×; the focused nine-block delayed/shape-record medians are 0.999×/0.992×, with all initial and final samples retained. Remaining ODraw raw/owner/error work and W4/W4b/W8 integration stay required.


- Coordinator follow-up [0362d](0362d-odraw-picture-package.md): ODraw picture publication now uses immutable `Package`/`View`, core `Commit`/diagnostics, exact content-bound patches, explicit options/raw limits and typed family errors. Every consumer migrates. The initial retained-memory regression was rejected and repaired with a shared base/delta, borrowed change getters and bounded small-table lookup. All 5,156 consumer tests, all-consumer Clippy, the doctest and glossary pass. Final measured medians are at or below baseline; small-history peak is 21,504 versus 21,864 bytes and large-history peak is 21,504 versus 270,551,176 bytes. Both initial and repaired receipts are retained. Remaining ODraw raw policies/builders/errors and W4/W4b/W5/W8 integration remain required.

- Coordinator follow-up [0362e](0362e-staged-output-gate.md): the family gate now rejects artifact output on an `Edit`, and the inventory reports it even for unregistered families. Real-rustdoc coverage passes all 30 tooling tests. Review found OLE single-stream/batch publication bypasses and returned the pending wave for commit-only publication with every consumer migrated and no duplicate-render regression. Changed public OLE debug output and its error exhaustiveness also require correction. These are open engineering work; W4/W4b/W5/W8 integration and the final full gate remain required.

- Coordinator follow-up [0361b](0361b-iwork-error-extensibility.md): 25 shared iWork errors are now non-exhaustive and every affected consumer migrates with them. The inventory detects missing attributes. Workspace all-target/all-feature compilation, eight-crate Clippy, 417 focused tests, 47 tooling tests, fresh common rustdoc and glossary pass. The common inventory's remaining raw `NestedFieldEdit` descriptor is reviewed as a naming-heuristic false positive. Other raw/shared API work, W4/W4b/W5/W8 integration, performance repairs and the final full gate remain required.

- Coordinator follow-up [0361c](0361c-iwork-substrate-error-extensibility.md): the physical, archive, text-appearance and legacy iWork errors now follow the non-exhaustive contract, with every affected consumer migrated. Workspace all-target/all-feature compilation, eight-crate Clippy, 413 focused tests, four fresh rustdoc inventories, the iwa glossary profile, formatting and both locked tool checks pass. The inventory retains 45 required raw/archive `with_limits` migrations and the reviewed `EntryEdit` descriptor false positive. W4/W4b/W5/W8 integration, performance repairs and the final full gate remain required.

- Coordinator follow-up [0361d](0361d-iwork-raw-limits.md) and repair [0361e](0361e-iwork-raw-performance-repair.md): physical IWA records, catalogs and Snappy operations now require limits on their canonical methods; all consumers migrate and 45 raw `with_limits` rows plus the limit aliases are removed. Output admission precedes allocation/writes, encoded entry ceilings propagate into the physical package, and exact-budget compression remains valid. The full seven-crate suite passes 5,088 tests; workspace compilation, seven-crate Clippy, fresh rustdoc/glossary, tooling, fuzz and locked tool checks pass. Final private repairs pass 314 substrate tests and Clippy. Initial framing, small-copy and compression regressions were rejected and repaired; final small/incompressible compression medians are 0.951×/0.682× with fewer allocations. The longer large-copy median is 1.002× with mixed blocks; all attempts remain recorded. Raw constructors, remaining owner-policy propagation, other shared/raw audits, W4/W4b/W5/W8 integration and the final full gate remain required.

- Coordinator review [0360a](0360a-xlsx-resource-typing-review.md): W4's resource typing remains unaccepted. Several helpers infer kinds from strings and discard actual observations; a concrete OLE-object depth failure reports a node limit, and connection namespace-binding exhaustion reports bytes with the wrong ceiling. The wave is returned for explicit kinds/counts, preserved malformed-input classification and focused boundary/host checks. Exact signed no-op cross-copy replay remains legal; an unconditional refusal proposed during review was withdrawn. W4 repair, final native acceptance and integrated gates remain required.

- Coordinator review [0362f](0362f-ppt-replay-performance-review.md): the PPT first-commit candidate is preserved on `uo2/ppt-0362f-candidate` at `be72b3a163`, not merged. Early candidate-length admission was repaired and targeted validation passes. Five guarded blocks improve original first commit to 0.9675×, but matched-Sip forward/inverse replay medians are 1.0077×/1.0466×, slower in four of five blocks. W8 must repair and remeasure replay; no performance exception is accepted. All completed, interrupted and superseded evidence remains recorded. Pending wave integration and the full final gate remain required.

- Coordinator follow-up [0361g](0361g-iwork-text-wire-limits.md): seven raw text-wire APIs now take explicit limits under their canonical names; all 30 affected source/consumer files migrate together, including the separate fuzz target. All 2,767 targeted tests, five-crate compilation and Clippy, fuzz compilation, fresh rustdoc/glossary, inventory and formatting pass. Reproducible normalized body hashes confirm the implementation bodies are unchanged. The seven character-style raw entry points and retained owner-policy propagation remain required; empty heuristic inventory output is not full closure. W4/W4b/W5/W8 integration, performance repairs and the final full gate remain required.

- Coordinator review [0362g](0362g-ppt-replay-buffer-review.md): the replay-buffer candidate `e3aa14d255` passes 2,109 distinct PPT/CFB tests and both Clippy checks, but remains unmerged. Five native blocks improve matched-Sip forward/inverse replay to 0.9685×/0.9141×; original-reference inverse remains 1.0121×, slower in three of five blocks. All observations are retained. Separate heap attribution reduces allocation calls and reported peak heap, so the small RSS delta is not treated as proof of a heap regression or waived as an exception. W8 has a separate in-place-buffer candidate awaiting validation and measurement; pending integrations and the full gate remain required.

- Coordinator review [0362h](0362h-file-performance-review.md): W8's historical Part A timings remain recorded, but final acceptance needs equivalent leaf legs, validated published results, repeated-output measurements, calibrated raw timing totals and failure/host-overlap guards. The broad 1.05 threshold and unsupported code-placement attribution do not close the current contract. No production code is merged by this review; pending integrations, performance repairs and the final workspace gate remain required.

- Coordinator follow-up [0362i](0362i-odraw-limit-kind-names.md): ODraw traversal and image resources now use `LimitKind` and `image::LimitKind`; both old root names are deleted and gated. All consumers migrate together. The 95 unit tests and doctest, all-target Clippy, fresh glossary, five-consumer all-target/all-feature compilation, tooling and formatting pass. Normalized source comparisons cover every changed Rust file and the unchanged relocated enum. Remaining ODraw raw/owner/error work, wave integration, native repairs and the final full gate stay required.

- Coordinator acceptance [0362j](0362j-ppt-replay-in-place.md): the PPT first-commit candidate and both replay repairs integrate together. All eight source files match validated candidate `c0f346607b`; 2,110 distinct PPT/CFB tests and both Clippy checks pass. Five guarded blocks per reference improve every first/steady/forward/inverse comparison: original first commit 0.9550×, forward 0.7377×, inverse 0.9153×; matched-Sip inverse 0.7663×. The 0.1682% matched-Sip process RSS delta remains reported; full-workload heap attribution reduces allocation calls by 2,770,500 and reported peak heap from 117.96 MB to 117.90 MB, with unchanged retained patch bytes. This closes 0357b/0362f/0362g's PPT first-commit/replay performance repair. Remaining wave integration, raw/shared work, glossary/inventory closure and the final workspace gate stay required.

- Coordinator acceptance [0361h](0361h-character-style-required-limits.md) and repair [0361ha](0361ha-character-style-output-repair.md): all seven raw character-style APIs require a wire policy, with every Pages/Numbers/Keynote consumer migrated and typed failures preserved. Aggregate admission precedes nested decoding and output allocation. All 57 raw unit tests, eight compile-fail doctests, 1,015 owner tests, four-crate Clippy and the raw glossary pass. The first memory regression was returned and repaired; every final allocation route is equal to or below the base, and all 11 native routes improve (0.3521–0.9618× medians). All 60 qualified ABBA blocks, the interrupted warmup and exact fixture/output comparisons remain recorded. Retained footnote policies, remaining raw/shared audits, wave integration and the final workspace gate stay required.

- Coordinator follow-up [0361j](0361j-pages-registry-error-fidelity.md): Pages section text-style registry failures preserve their actual resource kind and counts, replacing the fabricated RegistryBytes 0/0 mapping. All 209 Pages library tests, the File classification regression and all-target/all-feature Pages/ops Clippy pass. Successful read/edit/commit/replay implementation bytes are unchanged. Remaining owner policies, shared error conversions, wave integration and the final full gate remain required.


- Coordinator checkpoint (2026-10-22): the unpublished W8 Part A integration
  candidate is committed on `uo2/coordinator-w8-part-a` through `4d464d60b7`
  (`8c4ccb269f` assembly, `17d4cf011d` ODF/OPC isolation and `4d464d60b7` typed
  ODF source errors). Its full ops library run passes 962 tests plus the added
  DOCX policy-isolation regression; 217 doctests, all-feature and ODT-only
  Clippy, and no-format compilation pass. It is not merged or performance
  accepted. ODF semantic-policy propagation, File probe review, W4/W4b
  integration, corpus dispositions and the final gate remain required.
  W4's retained/replay memory repair continues on `uo2/w4-xlsx`; the latest
  implementation removes duplicated deferred endpoint catalogs (`913ed0f3bd`)
  and awaits matched allocation evidence. Protected constructor v4 binaries
  and allocation/fidelity receipts are committed through `7686d7014c` on
  `uo2/coordinator-constructors`; native qualification remains pending.
  Numbers, Pages and Keynote owner Snappy policies remain active engineering
  work. No new performance exception or redaction deferral is accepted, and
  nothing has been pushed.

- Coordinator documentation follow-up (2026-10-22): design §16.4 now names
  the implemented `litchi-sheet-text` leaf and its facade re-export, and the
  retired iWork facade now points to `File` and immutable leaf packages.
  This closes the queued §16.4 update without changing the historical
  measurement records. Remaining implementation, performance and final
  integration work stays open.

- Coordinator checkpoint (2026-10-22, after the ingress review): the W8 Part A
  candidate now includes `8d6abdab69`, applying semantic options through all
  ten prepared ODF owners and four flat families. Validation passes 2,752
  affected library tests (965 File tests), the full 1,768-test ODS suite,
  299 doctests, consumer compilation, Clippy, feature builds, selected
  glossary profiles, boundaries and tracked disposition checks. The totals
  overlap and are not a distinct-test sum. The candidate remains unmerged
  and unqualified for native performance. The repaired File harness is
  being integrated from `dd1052e06d`; its worker-source 109-route untimed
  receipt is retained separately from the coordinator-source validation.
  W4's `5e0672b53c` repair reduces retained history and replay peak memory
  below its original baseline; source-owner admission still allocates more
  and remains under repair. Numbers retained Snappy policy verification is
  committed through `5232e9bcdb`, Pages/Keynote through `eaade0241d`, and
  footnote validation through `660205a29c`. ODraw's final consumer checks
  are recorded in `80ccd68391`; W5's retag optimization is awaiting fresh
  allocation comparison. These branches are not accepted by this checkpoint.
  Constructor evidence through `2f8fa9534b` retains four interrupted v4
  native attempts with zero qualified blocks, including the worker-overlap
  correction and the later unrelated host test. Required flat ODF raw
  limits, remaining wave integration, native qualification, the complete
  inventory and final gate stay open. Nothing has been pushed.

- Coordinator checkpoint (2026-10-22, retained-header and integrated harness):
  the Part A candidate through `fd5f513df6` now records all 109 untimed File
  routes passing on integrated source `d39d179353`. Exact output and semantic
  receipts match the worker run; native acceptance still awaits the final
  integrated wave source. Constructor repair `536d7fdcca` closes the reproduced
  stricter-policy bypass for retained unknown IWA headers; all 115 core tests
  and Clippy pass. Its fresh matched-v5 build preserves all eight output files,
  and all 15 allocation routes remain at or below the original reference.
  V5 native qualification is pending; interrupted v4 attempts retain their
  original source attribution. W4 continues repairing source-owner admission.
  The W5 hybrid representation improves source-backed cells allocation but
  leaves eager byte/peak residuals, so it remains under repair. Flat ODF raw
  limits and the DOCX durable-ink Commit repair are active. These candidates
  are unmerged; remaining content-matching replay, raw audits, complete
  inventory/glossary closure, native qualification and the final gate remain
  required. Nothing has been pushed.

- Coordinator checkpoint (2026-10-22, W4/File assembly and replay review):
  `uo2/coordinator-w4-integration` is committed and clean at `eafeed22a4`,
  combining W4 `b6a5d10935` with Part A `fd5f513df6`. The merge fixes the
  facade sparkline example, six harness detector consumers, the iWork-only
  detector feature build and the combined harness lockfile. Nine affected
  crates pass all-target/all-feature compilation; 119 PPTX metadata tests,
  boundary/matrix checks, formatting and locked all-target harness compilation
  pass. The feature matrix has 27 initial passes and all four repaired iWork
  cases pass on rerun. Record 0360c retains every failed and successful log.
  This is an integration candidate, not trunk or performance acceptance.
  A dry XLSB merge identifies three source conflicts plus generated matrices
  and lockfiles; the actual XLSB merge remains coordinator work.

  DOCX durable Ink source-owner repair `76148455d5` and receipt `f97f0eef7a`
  pass 58 integration tests, 1,489 library tests and 170 doctests, with every
  matched allocation metric at or below baseline. This closes the earlier
  allocation returns; native qualification remains open. Flat ODF receipt
  `965beae5c9` independently verifies both 606-file source sets, protected
  binaries, equal locks and all 36 raw allocation legs. Constructor protocol
  `c72ebf4816` resumes only missing verified blocks without changing the five
  required blocks per route or host/timing guards. Its first matched-v5
  attempt was interrupted by an unrelated build during calibration; receipt
  `631fc7be4a` retains it with zero qualified blocks. All heavy-command holds
  are released.

  W4's `72940100a1` policy compaction improves memory but leaves owner
  admission above the original specialized-family baseline; it remains
  returned. `528e04e6d5` fixes no-op metadata base admission and passes its
  focused common/XLSX tests. W5's captured-package reuse and state compaction
  through `cf3c9a4eb6` pass focused/full affected tests but retain eager XLS
  allocation/peak residuals. File replay `bb745b0d8c` uses a bounded streaming
  content proof. Its signed-state review preserves existing facade admission:
  signed files cannot create an Edit, while leaf exact signed no-op replay
  remains legal. PPTX cross-copy replay and the remaining leaf content scope
  repairs are still under review and validation.

  Review found additional replay-shape debt: DOCX/PPTX metadata and several
  primary/source-backed replay methods still return Package/Self. These need
  Commit migration with all consumers, shared triad review and gate/inventory
  coverage. Raw ODF XML policy work from 0362n, remaining wave integration,
  native qualification, the complete inventory/glossary and final 53-check
  gate remain required. Nothing has been pushed.

- Coordinator checkpoint (2026-10-22, XLSB assembly): the candidate on
  `uo2/coordinator-w4-integration` is committed through `f3fed35c0d`, with
  source merge `933ce43de3`. Record 0360d retains all failed and successful
  validation attempts. The final source passes 1,317 XLSB tests (11 ignored
  doctests), 23 facade tests, 36 unified XLSB tests, eleven-package
  all-target/all-feature Clippy, the XLSB-only ops Clippy profile, the main
  harness and all nine XLSB harness checks, both XLSB glossary profiles,
  formatting and tooling checks. OPC/XLDM tests pass 824 cases. These
  totals overlap earlier runs and do not represent the final workspace gate.
  The assembly remains outside trunk pending performance qualification and
  the remaining source repairs. The standalone XLSB raw API audit is
  committed as `3267b9d627` (0360e); its required policy migrations remain.

  Shared metadata scope candidate `00e6add236` is returned by
  [0360ab](0360ab-w4-metadata-composition-review.md): independent assertions
  reproduce inverse failure after creating a family while others are absent,
  and refusal to compose a cancelled edit with a disjoint family change.
  The worker's subsequent repair is under review, including independently
  captured scopes and actual relationship-ID transitions. DOCX/PPTX metadata
  Commit migration is committed as `2eca82b0f8` from the assembled API, with
  focused suites, consumer checks and Clippy passing; integration with the
  repaired common scope still requires validation. The general replay
  return-shape gate and other primary replay migrations remain required.

  W4 cold-reader allocation measurements also expose three extra calls and
  1,016 extra bytes per first cell/style/cursor read against its original
  reference; this is returned for attribution and repair, alongside the
  source-owner admission residual. ODF raw policy testing found the default
  token ceiling narrowed historical ODS admission. Its worker repair
  preserves the owner default envelope while retaining explicit caller
  policies, with further edited/replayed-owner propagation checks pending.
  W5 deferred XLS projection remains a reviewed proposal, not an accepted
  implementation; commit-time target validation must remain intact. Native
  qualification, remaining integrations, the complete inventory/glossary
  sweep and final 53-check gate remain required. Nothing has been pushed.

- Coordinator checkpoint (2026-10-22, return contracts and metadata scope):
  `0dfb00b590` lands the return-type gate and inventory on trunk. Registered
  local edits must return their family's Commit, and registered owner replay
  must publish that owner. DOCX minimal/full profiles now verify 39/40 replay
  returns. Record 0362q retains 38 focused tests within 1,332 passing tooling
  tests and both fresh DOCX rustdoc profiles. Its full inventory still has
  35 rows, including four replay returns; the remaining raw/model/projection
  rows require migration or documented contract classification.

  The metadata integration candidate is committed as `5532099873` on
  `uo2/coordinator-metadata-integration`. It combines DOCX/PPTX Commit replay
  with the shared-property repairs through `9f9fa3930a`, preserving the W4b
  deferred snapshot helpers. All 17 common transaction tests pass, including
  the two regressions returned by 0360ab and creation-ID/removal-inbound
  guards. The fresh run passes 472 common tests, 230 focused format tests,
  five-package all-target/all-feature Clippy and formatting. Record 0362o
  retains the exact source proof and all command logs by size/hash. This
  closes the correctness return; performance qualification and trunk
  integration remain open.

  Primary PPTX/DOCX/RTF replay Commit migration is committed separately as
  `1c7eb1e84c` (0362p), with migrated crate/tool consumers and locked harness
  checks. Its coordinator integration and result-rule registrations remain.
  ODF policy propagation `747cfd072d` is returned by 0362nb because its raw
  reader exposes a core error and reconstructs the precise ODF dimension
  after conversion. Repair `1042fd7ecc` returns the crate error directly;
  raw tests and affected checks pass, with full post-repair validation still
  underway. No performance acceptance is claimed for either ODF revision.

  Three workers continue W4 cold-reader allocation repair, ODF/XLSB raw-policy
  work and the remaining DOCX API audit. Native qualification, remaining
  integrations, the complete inventory/glossary sweep and final 53-check gate
  remain required. Nothing has been pushed.

- Coordinator checkpoint (2026-10-22, replay integration and raw-policy repairs):
  `fddfbdb0a0` extends the return gate to direct shared Commit types, nested
  type arguments and declared method generic parameters. Record 0362qa retains
  1,334 passing tooling tests and the existing DOCX profile checks. All five
  RTF replay returns also verify against the worker's fresh candidate capture;
  its inventory has no remaining rows in the implemented categories. Those
  RTF source changes and policy registrations are not on trunk yet.

  The combined replay/metadata candidate is committed as `ad1c6cede8` on
  `uo2/coordinator-replay-integration`, in the metadata integration worktree.
  Record 0362pa fixes inverse RTF tail replay diagnostics and migrates the
  remaining ops Commit consumers. Record 0362u describes the assembly and
  bounded validation. Full merged testing additionally found a stale ODF
  replacement-range assertion and W4's accidental DOCX durable-error mapping
  change; both are repaired in the merge commit. The ops library suite now
  passes all 969 tests. The remaining 225-group affected-package run, consumer
  harness checks and final receipts are still in progress. This is not the
  final workspace gate or trunk performance acceptance.

  RTF primary/durable/redaction replay is committed through `cbf4e30276`, and
  logical tail replay span storage through `30ca8a13d8`. Its matched 1 MiB
  probe removes one 1,048,606-byte allocation from forward and inverse replay,
  with unchanged authored allocation totals and exact authored/durable bytes.
  The fresh ops consumer check on that standalone branch still encounters the
  DOCX/PPTX consumers already repaired by the coordinator; the complete merged
  source remains the required validation target. DOCX paragraph-stream API
  cleanup is committed as `b96137acb4`; paragraph-copy raw codec limits remain
  in progress.

  XLSB raw-policy work is committed through `8d59abb26f`: comments/calc,
  styles and host style parsers, filters/page breaks/cell watches, and named
  ranges now take required retained raw limits, with their consumers and
  tests migrated. The last slice passes 741 library tests, integration tests
  and Clippy. Remaining formula, XML and host raw codec slices are required.
  ODF's raw reader repair passes its full affected suites and Clippy. Review
  0362nca found one remaining namespace URI conversion through a coarse core
  error; source repair `3f802aec9d` now creates that typed error at its origin.
  Coordinator integration and native qualification remain open.

  W4's cold-reader prototype restores reference allocation counts, leaving
  +8 allocated bytes and +16 peak bytes per measured operation. Source review
  0360ac returns its defined-name storage change because proof-backed reopens
  retain duplicate catalogues. The repair must preserve shared history and
  include managed/custom-policy admission measurements. No performance
  exception is granted. Three workers continue these bounded repairs; native
  qualification, remaining integrations, the complete inventory/glossary sweep
  and final 53-check gate remain required. Nothing has been pushed.

- Coordinator checkpoint (2026-10-22, resumed handoff and combined repairs):
  Read the updated external `.coord/ctx-resume.txt` and
  `.coord/triage-2026-10-22.md` before continuing. The replay integration
  branch now includes trunk `a18593385c`, RTF through `5909588b99`, DOCX
  paragraph-copy through `9f1b3eb45f`, and durable Ink through `f97f0eef7a`.
  Its current head is `3c1d25de24`. Conflict resolutions preserve inverse
  RTF diagnostics, generic return checking and Ink source-owner sharing.
  Record 0362u retains 63 passing focused RTF tests and 144 combined DOCX
  tests, strict DOCX/ops/XLSB Clippy, all five RTF result rules and fresh
  DOCX minimal/full profiles with 42/43 result contracts. These are
  integration receipts, not final performance or trunk acceptance.

  The resumed full gate on older assembly `ad1c6cede8` completed: its actual
  script ran 52 checks, with 48 passing and four failing. Workspace docs
  passed, including the earlier memory-sensitive IWA documentation. Ops
  failed on seven stale RTF consumers and two XLSB refusal assertions;
  committed repairs `0641d926aa` and `759b09f0d9` now pass focused checks.
  Facade/workspace tests also found a missing doctest import and ZIP refusal
  mismatch; native-resave found an out-of-date lockfile. ZIP repair
  `a4d69035be` passes 104 affected tests and strict Clippy on its branch.
  The doctest and locked-tool repairs are being validated in the assembly.
  All failed logs are preserved; the earlier 225-group run remains partial.

  W4 owner compaction and sequential-reader policy propagation are committed
  as `da0068d02b`: 1,816 library tests, five reader-policy tests, 18 tab-state
  tests and Clippy pass. Its allocation residual and proof-backed history
  costs remain unaccepted. Capability review found that old slicer/timeline
  stores exposed physical relationship IDs and part names; the typed raw
  replacement `47c89d20da` is under review for bounded policy, fidelity and
  full validation. XLSB raw-policy slices reach `1309cfc15e`, but retained
  default-policy leaks are still being repaired; its closure claim is not
  accepted. No capability golden was changed to hide these gaps.

  The updated signed-replay decision permits exact no-op replay on matching
  signed content through both leaf and File, preserving signed bytes.
  Candidate `7e7cfe421c` is under review for removal-authorized, changed-base
  and freshness cases. Changed signed content remains refused. DOCX's fresh
  inventory has 27 rows in implemented categories; variables and SVG
  migrations are active, with remaining document/section/story work required.

  Record 0361fda formally returns iWork constructor and retained-header
  instruction regressions from the resumed 0361fd evidence. New constructor
  repairs require matched measurement; no host-noise waiver applies. ODF
  raw styles, W5 legacy costs, ODP retained-content heap and DOC body-toggle
  work continue in isolated branches. Native qualification for all remaining
  candidates, final inventory/glossary, corpus dispositions and the final
  combined gate remain required. The user authorized additional parallel
  workers; builds remain capped. Nothing has been pushed.

- Coordinator checkpoint (2026-10-22, legacy/ODF assembly and replay review):
  The integration branch reaches `bf529cb322`. It contains the detector and
  resave-lock repairs `557577e512`, ZIP integration `532cd8634b`, and W5/ODF
  through `8d1199f45e`. Record 0362uc preserves all three failed all-features
  harness attempts and the passing repairs: current XLSX adapters take
  precedence over legacy adapters, and the XLSB CRUD driver uses the fixed
  error reason. The locked workspace all-target/all-feature check, locked
  harness and native-resave checks, benchmark matrix and crate boundaries
  pass. The ZIP integration passes 657 tests with two existing ignored tests;
  the detector facade passes all 20 doctests. These remain assembly receipts,
  not final trunk or native-performance acceptance.

  DOCX SVG cleanup `135a7b6e63` is being validated in that assembly. Variables
  reach `5200fb8743`, with a further borrowed/owned input-policy review open.
  Standalone document, section, story/glossary, paragraph-copy and
  redaction/sanitize/hyperlink API migrations have separate workers. Public
  raw codec capabilities must remain available without renamed snapshot
  publication or additional per-operation copies. The fresh full-feature
  SVG capture passes both policy checks; a separate minimal-feature capture
  and the final combined inventory remain required.

  XLSB `89146628f3` re-admits durable content under the recipient's actual
  raw limits rather than rejecting a wider producer profile. Review found
  additional stale-base error paths; 0360el is repairing these and checking
  byte admission before copying replay images. Separate ODF and legacy
  reviews are repairing public replay/freshness classifications. Substrate
  scope record 0362rd is an audit proposal, not an exemption or completed
  migration.

  Signed replay candidate `7e7cfe421c` was returned: an unsigned endpoint
  cannot be promoted to signed merely because it retains old source bytes.
  The repair must construct and replay an actual signed no-op while preserving
  independently reopened unsigned equivalence. ODP candidate `bd7a095a56`
  reduces cold heap but retains a semantic reread/allocation regression;
  0362ya records it without acceptance. The iWork object-size and constructor
  costs, W4 admission residual, DOC body-toggle work and RTF refusal cost
  remain open. Record 0359r rejects the measured packed-Refusal experiment.
  Final semantic integration tests, native qualification, corpus dispositions,
  inventory/glossary closure and the complete assembled gate remain required.
  Nothing has been pushed.

- Coordinator checkpoint (2026-10-22, bounded raw models and semantic assembly):
  Record 0362uc now records all nine combined legacy/ODF semantic groups on
  `50122ae80a`: 750 tests pass. The runner stopped at its target-size guard
  before the final ODT group, then passed that group after obsolete check-only
  metadata was moved outside the target. Both stops and external log hashes
  are retained; the eight passing groups were not repeated.

  The assembly advances through `8d50e23436`. It integrates bounded DOCX
  variables through `7dae046af8`, the paragraph raw-codec repair through
  `e1af87ed91`, DOC body-toggle staging through `2452934481`, signed exact
  no-op replay through `ba8f9610aa`, and DOCX/iWork error classifications
  through `938108c265`. Raw paragraph replay accepts reopened matching XML;
  package replay keeps its full artifact proof and preflights recipient limits
  before copying. The combined checks on this head are running: variables,
  the five related DOCX integration suites, DOC atomicity, and DOC/DOCX facade
  suites have passed. This is an assembly checkpoint, not final acceptance.

  Review record 0360ada returns XLSX raw batch admission, retained-policy and
  owner-options gaps. Record 0359sa returns legacy raw stale-base errors that
  still use corruption categories; actual stale replay must be PatchConflict.
  ODP progress repair now also preserves secondary callback diagnostics; its
  inline error growth was caught and corrected before integration. The ODP
  streaming correctness repairs pass on their branch, while matched original
  0358a lifecycle measurements remain in progress.

  DOC and RTF measurement attribution was tightened. DOC's historical
  single-route whole-process profile included an unselected 1k batch seed;
  route-local ordinary-edit instructions are 0.988216 of base, and the probe
  selection correction is queued as `428bf4d44a`. The RTF cold parser delta
  disappears with matched executable path/arguments/environment, as recorded
  in `eabd1d9f47`; native timing and heap qualification remain required. iWork
  fixes the mutable MessageInfo length admission bypass, but its constructor
  cost and one retained-validation case remain returned.

  Remaining DOCX raw/document/section/story migrations, signed same-value and
  cancelling edits, IWA detection consumers, W4/W4b returns, public ODF error
  boundaries, final inventory/glossary, corpus dispositions, native performance
  qualification and the complete assembled gate remain required. Only the
  comprehensive redaction audit is deferred. Nothing has been pushed.

- Coordinator checkpoint (2026-10-22, completed combined checks and next returns):
  Record 0362ud commits all seven passing checks on `8d50e23436`: 183 semantic
  and facade tests, strict DOCX/DOC/ops all-target/all-feature Clippy, and the
  locked all-feature performance harness check. The integration branch then
  advances to `6571179db7`, merging the corrected DOC probe selection and
  seven typed legacy raw replay conflicts. These later changes still require
  combined validation; record 0362ud describes its exact earlier source head.

  The ODP freshness candidate preserves progress and secondary diagnostics
  with its original error layout, but boundary review found missing facade
  plumbing and lossy save/materialization conversion. The additional return
  is recorded in 0362rha; 0362rhb is repairing it. Record 0362yda returns the
  matched ODP allocation/instruction increases and requires route-level
  attribution and semantic output checks. Native qualification is still open;
  noisy timing attempts are retained without an acceptance claim.

  DOCX story projection sharing passes its focused checks, but original
  consuming-Commit baseline measurements and independent raw replay chains
  remain required. Redaction/sanitize review found admission after allocation
  and caller-controlled relationship IDs in changed public errors; 0362rga
  is repairing them while preserving artifact-bound inverse capability.
  XLSX raw admission reaches `2d44bc4c55`, with independent review, strict
  Clippy and matched insertion costs still pending. DOC's public body-text
  conflict variant also needs the canonical PatchConflict name (0359sc).

  All previous remaining engineering, corpus, inventory/glossary, native
  performance and final assembled-gate requirements still apply. Nothing
  has been pushed.

- Coordinator checkpoint (2026-10-22, replay validation and resource recovery):
  Record 0362ue preserves the combined replay runs and their corrections.
  On `5fbcfc45e7`, 1,487 DOC, OLE and DataSpaces tests passed, with three
  existing ignored tests. The next facade group stopped on denied
  `unused_qualifications` lints introduced by a module-level test import;
  `eb931ca487` corrects that import. The earlier protected-DOC assertion was
  also corrected: matching protected content still requires authorization,
  while changed content returns `PatchConflict`.

  The integration branch reaches `c2122e2678`, including DataSpaces error
  exhaustiveness, legacy alias removal, the qualified substrate scope audit,
  and DOC's consuming final handoff through `76391f1a71`. The DOC change
  removes the redundant Vec clone; conversion into Arc storage still copies
  once. Its matched performance qualification remains open. Forty glossary
  checker tests pass after the new alias bans. A fresh seven-group combined
  run is active on this fixed source; no result is inferred for unfinished
  groups or newer worker branches.

  Source branches now include the ODF public error-boundary migration
  `6ad7dfd8e8`, canonical ODF streaming-limit and reader names `84469fd9e4`,
  DOCX section migration `96c9c4ac5c`, and IWA package-triad migration
  `5a5715688a`. Their remaining consumer checks and combined validation are
  required before acceptance. DOCX document memoization was returned again:
  generation tracking added allocations and invalidated unchanged branches.
  Pinned Rust 1.95 source review confirms that the existing weak key already
  invalidates on `Arc::make_mut`; the suspected in-place mutation hole was
  not present. Memo allocation and branch reuse still need repair. Story raw
  replay is under
  independent review and still needs original-lifecycle qualification.

  Performance returns remain active for XLSX raw admission and reader state,
  XLSB reader reuse, DOC, RTF replay, ODP append and iWork constructors. The
  ODP error repair has passing public facade/save regressions and unchanged
  error layout, with final affected Clippy pending. Disk pressure temporarily
  stopped all new builds; owners reclaimed inactive generated caches without
  deleting protected measurement artifacts. Resumed commands use bounded
  existing targets and a free-space guard. Final native qualification,
  corpus dispositions, inventory/glossary closure and the complete assembled
  gate remain required. Only the comprehensive redaction audit is deferred.
  Nothing has been pushed.

- Coordinator checkpoint (2026-10-22, replay checks complete and API assembly):
  Record 0362ue now records all seven passing combined groups on
  `c2122e2678`: 1,578 tests pass, with three existing ignored tests; strict
  all-target/all-feature DOC/OLE/crypto/ops Clippy and the locked all-feature
  performance harness check also pass. Receipt `eaae26801b` names that exact
  source; it does not validate newer branches.

  The assembly reaches `b588c1a65f`, integrating the canonical ODF names,
  typed ODF replay/progress and public error boundaries, PPTX classifications,
  DOCX document/story/section/redaction raw models, and IWA primary edits.
  DOCX section and redaction facade consumers are migrated in their respective
  local merge commits. These newer changes still require combined validation.
  Further reviewed XLSX raw/reader-state, XLSB admitted-reader and consuming
  DOC handoff repairs are queued for that validation snapshot.

  Signed cancellation candidate `f3c5accd86` is returned: restoring the
  endpoint alone leaves a changed native leaf patch, and unsigned coalescers
  acquire unnecessary original-value tracking. The repair must produce a
  genuine signed no-op patch, cover DOCX and invalid intermediate edits, and
  keep ordinary unsigned costs unchanged. Required OPC raw constructor policy
  and DOCX raw helper migrations remain active engineering work.

  Independent iWork review identifies a constructor benchmark mismatch:
  historical Vec construction was compared with new incremental collection
  admission. The large direct ratios are not same-operation regression
  evidence. Equivalent raw-constructor and full lifecycle qualification are
  still required, alongside safe admission-ledger optimization. XLSX reader
  evidence distinguishes requested bytes from lower live/peak heap and fewer
  allocation calls; native qualification remains open. DOC, RTF and ODP
  performance returns, corpus dispositions, final inventory/glossary closure
  and the complete assembled gate remain required. Only the comprehensive
  redaction audit is deferred. Nothing has been pushed.

- Coordinator checkpoint (2026-10-22, extended assembly and policy returns):
  The integration branch reaches `c48e382321`, adding reviewed XLSX raw
  admission/reader state, XLSB raw policy/admitted-reader reuse, consuming
  DOC body publication, and RTF shared replay storage. The original signed
  no-op return remains unresolved; candidate `2ed47bf56e` still retains a
  changed native patch behind an endpoint marker and is returned again.
  The repair must produce a genuine native no-op bound to the signed source,
  retain admission guards, and cover supported paragraph/run/story routes.

  Record 0362uf preserves the assembled consumer checks. The first stopped
  before compilation on the new ODF dependency's lock state; an offline
  refresh corrects the generated workspace lock and the tracked harness
  dependency edge. The next check reached XLSX and found a raw-budget plan
  type mismatch and ambiguous closure error. The correction keeps original
  owner metrics separate from the proposed relationship edit. These runs do
  not count as completed consumer validation. Fifty-six glossary/iWork
  checker unit tests pass; the earlier 1,578-test receipt remains tied to
  `c2122e2678`.

  Independent review returns two additional policy gaps. DOCX restored
  replay must retain recipient semantic/source/cache/execution policy on
  every path, including deferred endpoints, without losing proven same-owner
  fast paths (0362rn). XLSX externally supplied packages must retain custom
  cache policy and both package/facade execution contexts through deferred
  materialization (0360aea). Both are being repaired and require tests.

  OPC required-limit ingress and DOCX raw codec/helper migrations have source
  commits and are under integration review; all consumers must move together.
  The ODP lazy append-fact repair has no definite static correctness blocker,
  but needs exact-source lifecycle qualification. RTF final-source allocation
  and instruction measurements are running. Fair iWork constructor evidence
  identifies required current-field admission costs; complete consumer
  lifecycle qualification and safe optimization remain required.

  Resource recovery removed 8.1 GB of proven unmeasured debug artifacts after
  retaining hashes/provenance, and relocated the inactive root release cache
  intact outside the target. No measured executable was deleted. Native
  qualification, corpus dispositions, final inventory/glossary and the full
  assembled gate remain open. Only the comprehensive redaction audit is
  deferred. Nothing has been pushed.

- Coordinator checkpoint (2026-10-22, policy review and consumer rerun):
  The reviewed assembly reaches `d797add634`, adding corrected historical
  DOCX error records, XLSB retained-provenance presence assertions, the RTF
  final-source qualification receipt, and the external XLSX cache/context
  repair. Independent review confirms that the XLSX owner retains the supplied
  package's cache policy and both package/facade execution contexts through
  deferred reads and writes. These changes still require combined validation.

  Record 0362uf also preserves the DOCX compilation stop and correction:
  a missing documentation-macro comma and denied redundant qualification
  are fixed in `d43273a13b`. The fourth locked workspace/harness consumer
  attempt is running on the fixed assembly. No unfinished command counts as
  a pass; the last completed 1,578-test receipt remains tied to `c2122e2678`.

  The final-source RTF receipt reports 34 focused tests passing, about
  2.67 million fewer matched Callgrind instructions, and unchanged peak heap,
  with three first-promotion allocations across the measured replay workload.
  Native timing remains unqualified because compiler activity contaminated
  the attempt. Its evidence is retained without a native acceptance claim.

  OPC constructor review returns production calls that use default limits
  instead of the existing package's retained policy, including replacement
  relationship admission. XLSX raw aggregate review returns a lazy-payload
  allocation bypass: `try_iter_parts()` decodes parts before the tighter raw
  aggregate policy is checked. Both repairs remain required. DOCX raw helper
  review likewise requires selected limits through MCE, semantic projection,
  and output planning before allocation, while preserving shared resources.

  DOCX cross-owner restored replay is being repaired with an explicit bounded
  OPC artifact-stream path for deferred/effective endpoints. Producer source
  checks and recipient output accounting must remain separate, including
  nested restores; a full-artifact Vec fallback is not accepted. Genuine
  signed native no-op patches, iWork constructor correctness and performance,
  the other performance returns, corpus dispositions, final inventory/glossary
  and the full assembled gate remain open. Only the comprehensive redaction
  audit is deferred. Nothing has been pushed.

- Coordinator checkpoint (2026-10-23, complete consumer diagnostics and cleanup):
  The assembly reaches `d686be4af0`. Record 0362uf preserves the fourth
  and fifth check failures and their private DOCX accessor/ODT error-boundary
  corrections. The sixth attempt on `a814baa1de` used `--keep-going`, passed
  the format libraries, and exposed remaining ODF facade/native-test error
  migrations and invalid nested PPTX refusal guards. Its runner ended without
  saving a final exit receipt; the retained compiler log establishes failure,
  not an inferred exit code or duration. The harness was not reached. Two
  isolated DOCX/IWA test lints are corrected in `d686be4af0`; the other
  consumer repairs are assigned as 0362ug–0362ui. The last completed
  1,578-test receipt remains tied to `c2122e2678`.

  XLSX raw aggregate candidate `5726f03c7d` now admits the original ledger
  before proposed relationship allocations, retains admitted deferred payload
  lengths, and adds malformed declared-size regressions. Static review and
  combined runtime checks remain required, along with matched owned-open and
  first-read costs for the added metadata. OPC constructor and token policy
  work must preserve captured limits through every existing-owner operation.

  DOCX restored replay is adopting a generic typed artifact-stream bridge
  so producer errors survive the OPC boundary. Its final-output admission
  must match the recipient's actual source policy, preserve nested output
  accounting, and retain both primary freshness and secondary failures.
  Catalog-only checks and whole-file materialization do not close that work.
  Signed native no-op budget preservation remains under repair. The iWork
  retained-header optimization also remains returned: mutations that change
  nested encoded length must fall back to canonical serialization, and a
  prior optimized candidate cannot establish that its failures pre-existed.

  The owner requests periodic consolidation and removal of unused leaf
  worktrees. Cleanup now checks merged ancestry, uncommitted/ignored data,
  active processes and measurement references before removing source-only
  checkouts; protected measurement artifacts and branch references remain.
  All returned performance work, corpus dispositions, final inventory/glossary
  and the complete assembled gate remain open. Only the comprehensive
  redaction audit is deferred. Nothing has been pushed.

- Coordinator checkpoint (2026-10-23, reviewed consumer repairs and consolidation):
  The integration assembly reaches `7abf64f495`, with the independently
  reviewed XLSX raw aggregate admission repair, stable PPTX refusal guards,
  typed ODF facade errors, and native ODF consumer migration merged. ODF
  staging now recodes only invalid-format/parse and generic unsupported
  errors; source changes, resource limits, I/O, and typed refusals retain
  their classifications. Combined consumer and semantic validation remains
  required. The seventh consumer attempt ended during resource preflight:
  the runner and target-size process are absent and no command log or exit
  receipt exists. It establishes no compiler result. A separate runner starts the eighth
  attempt on `7abf64f495` with an explicit process and source receipt.

  Two clean, inactive merged worktrees were removed without forcing, freeing
  341,943,958 bytes. Branch refs and exact recreation commands are retained;
  eleven generated files were quarantined with hashes. Worktrees containing
  measurement evidence, protected artifacts, or tracked changes remain intact.
  The reduced cleanup receipt is recorded in `e726df19c1`.

  The XLSX lifecycle harness is returned for correctness checks before its
  measurements can qualify the raw aggregate change. OPC streamed publication
  is adding exact prepared ZIP output metrics so recipient admission can run
  before the first sink byte while reusing the prepared compressed data.
  Recipient-policy enforcement, raw helper limits, signed native no-op budget
  preservation, performance returns, corpus dispositions, final inventory and
  glossary, and the complete assembled gate remain open. Only the comprehensive
  redaction audit is deferred. Nothing has been pushed.

- Coordinator checkpoint (2026-10-23, required-policy assembly and second cleanup):
  The integration assembly reaches `b170204d17`. The eighth workspace
  diagnostic pass finished with exit 101 after 439.95 seconds; record 0362uf
  preserves its exact source, log hash and failure. Record 0362uj fixes the
  remaining mixed core/ODF consumer boundaries, ODS test imports and XLSX
  test calls without weakening assertions. The ninth locked workspace and
  harness check is running on the assembled source; no pass is inferred.

  Reviewed token-policy and required OPC constructor migrations are merged.
  Conflict resolution retains owner/recipient limit intersections, atomic
  content-type and relationship admission, tightened no-op relationship
  validation, and XLSX source/cache/package/facade execution policies. The
  raw aggregate reader and private fragment helper use the canonical names.
  ZIP admission now reuses final prepared bytes before publication, preserves
  scheduled preparation, and performs metadata scans only on policy paths.
  A metadata iterator covers directory records without enlarging the retained
  ordinary index. These changes still require combined runtime validation.

  The XLSX lifecycle harness now fails on fidelity violations and has an
  explicit source-version adapter for required OPC ingress. Current source
  history probes contain only current APIs; the historical e63 probes have
  pinned external source and overlay reproduction instructions. Measurements
  against the final assembly remain required.

  A second cleanup wave removed four clean, inactive merged leaves without
  forcing, totaling 6,484,524,650 bytes of checkout contents. Branch refs and
  exact recreation commands remain recorded in the cleanup receipt. Dirty
  trees, measurement artifacts and protected executables were retained.

  Independent review continues to return incomplete selected-limit and
  preallocation paths in DOCX glossary helpers, generic streamed-publication
  admission, and signed native no-op budget handling. Those branches remain
  outside this assembly pending repair. The iWork constructor and ODP lifecycle
  performance returns, native qualification, corpus dispositions, final
  inventory/glossary, and full assembled gate remain open. Only the comprehensive
  redaction audit is deferred. Nothing has been pushed.

- Coordinator checkpoint (2026-10-23, native stream assembly and third cleanup):
  The source assembly reaches `44901f50d1`, adding the bounded iWork
  constructor/inline-version repair, ODP bounded text memo and append repair,
  authenticated OPC stream publication and DOCX consumers, and concrete XLS
  cell iteration with deferred target reopening. Independent static reviews
  found no remaining blocker in these source proposals. Combined compilation,
  semantic validation and performance qualification still determine acceptance.

  DOCX merge resolution retains existing redaction patch types, typed
  conflicts, source/effective fingerprint memos, cache/read-ahead policy and
  native/facade contexts separately. It fixes an incoming context fallback
  that would replace a native context or manage an originally unmanaged source.
  New regressions exercise both combinations through deferred edits and both
  restore paths. Freshness checks remain non-materializing.

  Record 0362uf retains consumer attempts 10–13. The focused OPC library/unit
  test check progressed to external consumer imports, which record 0362ul
  repairs. The ignored workspace lock was refreshed offline after the iWork
  dependency change. The first assembled stream check then failed on new
  stream helper/type scopes, retained-field mismatches, generic ownership and
  denied lints. Those compiler repairs are in progress; no workspace, harness
  or semantic pass is inferred from the stopped checks.

  A third cleanup wave removed five clean, inactive merged leaves without
  forcing, totaling 12,743,246,777 bytes of checkout contents. Branch refs,
  exact recreation commands, and measurement evidence remain retained; four
  ignored generated files were quarantined with hashes.

  DOCX raw helper selected-limit and consumer repairs, signed native no-op
  coverage/accounting, and ODraw residual review remain active. The XLSX raw
  baseline qualification build is being prepared. All native qualification,
  corpus dispositions, final inventory/glossary, and the complete assembled
  gate remain open. Only the comprehensive redaction audit is deferred.
  Nothing has been pushed.

- Coordinator checkpoint (2026-10-23, runtime admission and fourth cleanup):
  The integration assembly reaches `ff86f4a1c0`, incorporating the complete
  required-policy consumer migration, ODraw retained-policy/raw-limit closure,
  XLS concrete iteration/deferred reopen, reproducible ODP probe routing, and
  the XLSX historical raw qualification harness. Record 0362uf preserves the
  fourteenth and fifteenth workspace compiler failures and their repairs.
  The sixteenth all-target/all-feature workspace and harness pass is running;
  no unfinished result counts as a pass.

  OPC runtime validation on `8acd229235` passes all 714 library tests. It
  follows repairs for two policy-confounded fixtures and a real final-output
  admission bypass: appended replacement relationships must be counted even
  when the original relationship entry is omitted. ZIP runtime validation
  reaches 505 passing tests and one failing ZIP64/comment geometry regression;
  that failure remains assigned for diagnosis and repair.

  DOCX whole-owner replay now checks source freshness, complete base identity,
  recipient policy and the redaction inverse capability before publication.
  Same-owner and equal-content ordinary restores retain their supported paths;
  unrelated content changes cannot fall back to scoped replay. The raw-helper
  branch remains unmerged pending generated-root shape admission before capture
  and elimination of duplicate catalog planning, despite its passing leaf
  library receipt. Signed native no-op accounting also remains under repair.

  The tooling pass records 1,334 tests with one matrix coverage failure, plus
  archive-header boundary coverage and dependency-classification failures.
  The matrix repair now preserves allocator/instruction diagnostics and
  explicitly unqualified native results, without inventing comparison ratios
  for unpaired values. Crate-boundary fixes and the assembled rerun remain open.

  A fourth cleanup wave removed six clean, inactive merged leaves without
  forcing, totaling 37,259,489,031 bytes of checkout contents. Seven generated
  files (457,223 bytes) were quarantined with hashes. Branch references,
  recreation commands, measurement evidence and active worktrees are retained.

  Native qualification, full corpus replay/dispositions, final inventory and
  glossary, and the complete assembled gate remain required. The integration
  branch is staging these changes; they are not yet accepted into trunk. Only
  the comprehensive redaction audit is deferred. Nothing has been pushed.

- Coordinator checkpoint (2026-10-23, prepared DOCX admission and tooling pass):
  The integration assembly reaches `c9e8d8dfbf`. The DOCX raw-helper chain
  is merged after final review of prepared graph admission: known incoming
  catalog/graph bounds precede current-source capture, generated XML is
  admitted before retained allocation, and canonical reads reuse same-policy
  admission while preserving MCE and semantic validation. Merge resolution
  retains the current required OPC constructors and recipient-policy calls.
  The final leaf library receipt records 1,491 tests passing on `3b3a7c66fb`;
  it is separate from combined assembly validation.

  ZIP passes all 506 library tests; OPC's prior 714-test pass remains tied to
  `8acd229235`. All eight Python tooling unittest modules pass (1,340 tests)
  on `828d9473da`. Record 0362uf preserves the seventeenth workspace failure
  and its complete consumer repair. The eighteenth check confirms four
  glossary default-policy helpers are now used only by unit tests; production
  compilation rejects them as dead code. Their test-only scoping and another
  assembled rerun remain required. Neither stopped run reaches the harness.

  Corrected glossary/OLE records distinguish implemented source from pending
  qualification. Both W4/W4b tips are already ancestors; content-matching
  replay is implemented across the assigned formats. Public ODF archive,
  manifest and raw-writer limit paths remain required and are assigned to a
  dedicated leaf. Standalone generic CFB/ZIP primitives are not silently
  treated as format-facing API debt; public format leaks remain in scope.

  The frozen XLSX candidate at `ff86f4a1c0` passes 16/16 fidelity operations
  in each allocator/native/Callgrind run. Its instruction total is 0.2965%
  above the pinned historical baseline; owned-open and first-read costs are
  reported separately. Concurrent compiler activity means native timing is
  unqualified. Signed snapshot accounting is split into XLSX/shared, XLSB,
  and PPTX repair lanes; complete pre-capture budget enforcement remains
  under review. The XLS deferred-reopen qualification harness is also being
  corrected before measurement.

  A fifth cleanup wave removed three clean, inactive merged source leaves
  without forcing, totaling 18,782,352,655 bytes of checkout contents. Two
  generated bytecode files (323,671 bytes) were quarantined with hashes. All
  protected targets, measurement artifacts, active leaves and branch refs
  remain retained. The reduced receipt is merged through `a7e4393bbf`.

  Native qualification, full corpus replay/dispositions, final current
  inventory/glossary validation and the complete assembled gate remain open.
  Source integration is not yet accepted into trunk. Only the comprehensive
  redaction audit is deferred. Nothing has been pushed.

- Coordinator checkpoint (2026-10-23, combined DOCX pass and ODF raw assembly):
  The integration branch reaches `16ba04ff9f`. Required ODF archive, manifest
  and raw-writer policies are assembled, including retained facade policies
  and flat-projection XML limits. Internal ZIP scratch projections retain
  their explicit source-ceiling distinction without widening XML policy.
  Static review found no extra publication, source copy or scan. Combined
  runtime and performance qualification remain required.

  The assembled DOCX library passes all 1,515 tests on `d1a49d80a1`. The
  execution-context regression now exercises the deferred publication
  boundary directly and both restore paths; its earlier fixture and type
  failures remain recorded. OPC/ZIP all-target/all-feature compilation passes.
  Workspace checking stops on ODF policy getter/import and remaining consumer
  qualifications/imports across ODF, DOCX and XLSB. Those are assigned as one
  repair batch. Both harness compiler groups remain unreached. The prior
  workspace formatting pass is recorded separately from these later merges.

  Signed-edit review returned missing pre-capture admission across XLSX
  feature/cell paths, XLSB lazy patch payloads, PPTX source edits and DOCX
  facade XML captures. Shared dispatcher feature configurations and remaining
  budgets also need repair. These leaves remain outside the assembly. The
  semantic retention budget continues to charge native package/patch data,
  not derived parser/cache memory; signed replacement phases retain their
  staged charge and release discarded proofs before canonical capture.

  XLS deferred-cell qualification exposed original-source freshness gaps at
  commit, first read, output and same-owner replay. The source repair and
  independent replay/lineage harness are under review before measurement.
  Other compiler activity still prevents a quiet-host timing claim.

  A sixth cleanup wave removed six merged, inactive source leaves without
  force, reclaiming 31,566,664,232 checkout bytes. Eight generated files
  totaling 6,444,308 bytes were quarantined with hashes. Branch references,
  recreation commands, external DOCX build output, protected measurement
  executables and evidence remain retained. Receipt `dbd343ebd1` is merged.

  Native qualification, full corpus replay/dispositions, final current
  inventory/glossary and the complete assembled gate remain open. Source
  integration has not yet been accepted into trunk. Only the comprehensive
  redaction audit is deferred. Nothing has been pushed.

- Coordinator checkpoint (2026-10-23, workspace and harness compilation):
  The source integration reaches `3107f05f23`. Workspace checking with all
  targets/features passes on `7de3789ede`; the complete performance harness
  and isolated XLSX raw harness both compile with all targets/features on
  `808f9883a0`. Their consumer migrations preserve existing workload and
  publication timing boundaries. Record 0362uf retains exact source and log
  hashes for both successes and the preceding failures.

  The third broad Clippy pass reaches XLSX, DOCX and XLSB production/test
  lints after the complete OPC/ODF cleanup. Its terminal failure is recorded
  and the whole diagnosed batch is assigned for repair. These compilation
  results do not constitute the complete 53-check gate or native qualification.

  A separate diagnostic branch, `uo2/budget-validation`, composes signed
  native budget proposals with the current immutable owner and required-policy
  APIs. Its first compilation exposed a stale OPC retained-field reference;
  that failure and repair are recorded in 0362up. A new combined native/facade
  check is running on `3ae4cb2420`. Review still requires exact XLSX capture
  accounting without extra endpoint scans, PPTX sizing without a second graph
  proof, and final verification of DOCX admission before snapshot capture.
  This diagnostic branch is not accepted into the source integration or trunk.

  A seventh cleanup wave removed two clean, inactive merged leaf worktrees
  without force, reclaiming 12,522,187,690 checkout bytes. Branch references
  and recreation commands remain recorded; dirty leaves, build targets and
  measurement evidence were preserved. Receipt `e16821b336` is merged.

  XLS deferred-reopen fidelity/provenance repair, native qualification, full
  corpus replay/dispositions, final inventory/glossary and the complete
  assembled gate remain required. Source integration is not yet accepted into
  trunk. Only the comprehensive redaction audit is deferred. Nothing is pushed.

- Coordinator checkpoint (2026-10-23, isolated documentation pass):
  The exact workspace documentation gate passes on `0f8ce50050`, with
  warnings denied, one Cargo job and a 20 GB memory cap. The 316.00-second
  passing run has no diagnostic output and a 3,440,007,569-byte target.
  Preceding isolated runs exposed stale ODB, ODC and OPC links and a public
  OLE link to private state; all are corrected and their failures retained
  in 0362uf. No memory kill occurred in these reruns. This closes the
  documentation failure on that source, not the full 53-check gate.

  Integration now reaches `045d3721ae`, including the independently reviewed
  Clippy v3 batch and the Numbers native limit payload repair. Numbers now
  preserves exact kind/observed/maximum values through popup/reply helpers;
  reply-specific field/nesting mapping retains its existing public axes.
  The fourth broad Clippy run and a combined Numbers library run are queued
  together on the fresh documentation target. Results remain pending.

  Diagnostic budget composition reaches `152eb8a53b`. Scoped DOCX, XLSB and
  XLSX feature reviews are clear, but shared XLSX cell transfer/cancellation
  accounting, rich-text fidelity and native metadata preparation still need
  repair. Review returned late target admission, duplicate metadata planning
  and noncanonical limit entry points; no such behavior is accepted. A new
  combined native/facade compiler check is running. XLS deferred freshness
  review additionally covers cold materialization and package output handoff.

  The eighth cleanup removed three clean inactive leaves, retaining their
  branch references and exact recreation commands, and reclaimed
  18,784,234,793 checkout bytes. Generated cache was moved reversibly to an
  external hashed quarantine, reducing the main target by 440,535,840
  apparent bytes. Hard-link accounting and restoration evidence are retained;
  measured/release binaries and dirty worktrees were preserved.

  Source integration is still staging, not accepted into trunk. Final
  combined runtime validation, native qualification, corpus dispositions,
  inventory/glossary and the complete gate remain required. Only the
  comprehensive redaction audit is deferred. Nothing has been pushed.

- Coordinator checkpoint (2026-10-23, workspace Clippy pass and budget review):
  Workspace Clippy with all targets/features and warnings denied passes on
  `c3a9a0cef4` in 213.64 seconds. Its preceding run found an internal raw
  container size lint and a nested Numbers test import. The container keeps
  the complete ODF policy on the stack without an extra allocation; its
  narrow expectation applies only when both variants exist. The Numbers
  library run is still active on this same frozen source.

  Diagnostic budget composition reaches `62c11a3dde`. The reviewed combined
  XLSX sheet-property/filter repair `750d918386` charges the peak of sequential
  discarded native proofs and final endpoints, including shrinking and
  expanding edits. Common metadata `1929682421` passes scoped API/fidelity
  review for prepared layout reuse, escaping, relationship-ID release and
  no-op identity. Default-path overhead still needs performance qualification.
  Shared XLSX cell fidelity/accounting and metadata endpoint admission remain
  returned engineering work, assigned to separate workers. No diagnostic
  source has been accepted into the main integration or trunk.

  XLS source-freshness repair `c2ca0efbe8` passed static review and is composed
  with current immutable OLE policies. Combined compilation exposed regression
  test coordinate types and stale helper/API references; narrow fixes are
  committed and the seventh native check is active. Runtime source-fence and
  fresh F3 qualification remain required. The complete gate, native measures,
  corpus replay/dispositions and final inventory/glossary remain open.
  Only the comprehensive redaction audit is deferred. Nothing has been pushed.

- Coordinator checkpoint (2026-10-23, Numbers/XLS runtime and ninth cleanup):
  Main integration reaches `eed44807af`. Workspace Clippy passes on
  `c3a9a0cef4`, followed by 488 passing Numbers library tests and four ignored
  cases. Record 0362uf retains exact commands and log hashes. The glossary
  now closes integrated DOCX helper/model and ODF raw-limit source rows;
  fresh rustdoc inventory, native/performance and final-gate statuses remain
  explicitly open. These documentation changes do not claim acceptance.

  Diagnostic composition reaches `374fb0e5d9`. Common OOXML passes 482 tests
  on `af09b4fcf9`. A generated numeric XLS fixture repairs seven setup failures
  without weakening the new freshness assertions. The full XLS unit and
  integration run then passes 1,863 tests across 85 suites, with one ignored,
  on `374fb0e5d9`. XLSX library tests are still running on that frozen source.
  Native compiler checking has reached the facade, where the complete batch
  of consumer type/import/borrow diagnostics is assigned for repair.

  The second XLSX cell proposal fixes plain/rich transfers, source styles,
  ordinary transfer atomicity, existing merges and hyperlink aggregation,
  but review still returns effective-style composition, bounded no-op joins,
  exact no-op admission and repeated rich-text source scans. Separate workers
  own cell accounting, bulk rich reads and metadata endpoint admission.
  None is accepted as an exception. DOC final qualification is preparing a
  controlled, provenance-recorded baseline that retains current error codes.

  The ninth cleanup removes two clean merged leaves and reclaims
  12,518,628,337 apparent checkout bytes. Two ignored files totaling 151,969
  bytes are held in an external hashed quarantine with a restore script;
  branch references, active/dirty leaves, build targets and measurements are
  retained. Receipt `95904ba027` is merged into the main integration.

  Full combined runtime/gates, native and quiet-host qualification, corpus
  replay/dispositions and fresh inventory/glossary validation remain required.
  Source integration has not been accepted into trunk. Only the comprehensive
  redaction audit is deferred. Nothing has been pushed.

- Coordinator checkpoint (2026-10-23, strict budget compilation and scoped proofs):
  Main integration reaches `66110f0fb9`. DOC final qualification now includes
  the structural CHPX audit: removed history blocks are unreachable through
  the active FIB/BTE graph, and all 28 active formatting-page payloads and
  ranges are unchanged. The ordinary route is byte-exact. Record 0359qe
  preserves the controlled baseline, instruction/heap results and the pending
  quiet-host timing boundary. Reviewed DOCX Ink and ODP qualification harnesses
  are merged; fresh builds and correctness preflights are starting. Harness
  infrastructure does not itself establish performance acceptance.

  Diagnostic budget composition reaches `4470ecce16`. Strict combined
  all-target/all-feature native/facade compilation passes on `2cadd934b7`,
  followed by the facade without default features. Metadata runtime initially
  passes ten tests and exposes one empty-fixture setup error; the corrected
  populated fixture then passes all eleven tests, preserving exact-limit and
  one-below assertions. Record 0362up retains source/command/log receipts.
  Cache-independent endpoint accounting and bulk rich-transfer fidelity
  passed scoped review. The two PPTX runtime repairs and obsolete helper
  removal are composed, with combined PPTX runtime validation now running.

  The new cell-normalization proposal `06afe43fa0` remains under review for
  atomic joins, signed and source-equivalent no-ops, typed error preservation
  and default-path costs. It has not been accepted. IWork qualification is
  returned for its original footnote workload, semantic expectations and
  complete validation proofs; W4b/W5 orchestration is correcting analyzer
  output and stale-result isolation. These remain engineering requirements.

  The tenth cleanup removes one clean merged metadata leaf without force,
  reclaiming 6,262,042,222 apparent bytes. Its ignored 139,592-byte lockfile
  is retained in a hashed external quarantine with restoration instructions.
  Two other leaves remain because active processes hold their directories
  or files. Branch references, targets and measurement artifacts are retained;
  receipt `3baa225cf8` is merged.

  Final combined runtime/gates, native and quiet-host qualification, corpus
  replay/dispositions and fresh inventory/glossary validation remain required.
  Source integration is still staging, not accepted into trunk. Only the
  comprehensive redaction audit is deferred. Nothing has been pushed.

- Coordinator checkpoint (2026-10-23, complete facade return and qualification readiness):
  Main integration reaches `8300a93345`. Reviewed W4b/W5 execution tooling
  now prepares explicitly selected historical/current source pairs, guards
  the recorded baseline consumer adaptations, preserves real analyzer output
  and rejects stale result files. Its upstream-checkout rejection guard no
  longer matches the lexical test itself; all three reference tests pass on
  the assembled content. The final inventory note records decimal target
  limits, valid cache reuse and completed design/replay source work without
  claiming a fresh inventory or final gate pass.

  Diagnostic composition reaches `d4d1820dbe`. All eleven metadata-focused
  XLSX tests and all 27 PPTX module tests pass on `4470ecce16`. The broader
  all-feature facade suite on `55051a8e4f` passes 993 tests and fails nine.
  Record 0362upa returns the complete batch: DOCX classification, retained
  hyperlinks and signed no-op behavior; ODF typed-cause assertions; signed
  DOC/XLS no-op/cancellation policy; and signed PPTX cancellation. The two
  ODF assertions now retain their contract-correct core cause, and reviewed
  PPTX private unsigned staging permits cancellation proof before the signed
  publication fence. Their runtime rerun remains required. DOCX and legacy
  repairs are active. The runner stopped before XLSB, which remains queued.

  Shared XLSX normalization remains returned on the separate current-source
  repair leaf `b2a1e6e735`: explicit empty-cell fidelity, native string
  identity, failed-setter atomicity, compact join planning, bounded lookup
  cost and typed normalization errors must all be corrected. Temporary
  compiler overlays are no longer needed on that composed source.

  DOCX Ink qualification retains the original baseline and explicitly
  records shared-source changes for the integrated-route comparison. The
  streamed digest correction and all producer/runner/collector provenance
  consumers passed review; a fresh matched build is active. Earlier partial
  runs remain preserved. ODP's candidate release build succeeded, but its
  probe needs separate same-owner whole-artifact inverse proof and reopened
  content-scope proof; untimed correctness must run without a quiet-host
  requirement. The old binary cannot qualify the corrected probe. IWork
  baseline mapping preserves the original constructor and owner sources;
  its integrated source-delta successor is under review before real builds.

  W8 readiness preserves `f2bbbc8d07` as the original baseline, with a
  workload-equivalent compatibility overlay still required. The 0359qe
  matrix rows are reviewed as diagnostics, with busy-host wall results
  explicitly unqualified. Final provenance tooling still needs exact ops
  package binding. No source archive, binary or measurement result is
  silently reused under a different baseline label.

  Additional cleanup assessment retains leaves with active evidence handles
  and all protected build/measurement artifacts. Verified stale read-only
  Git commands from completed reviews were stopped; source and build work
  were not interrupted. The earlier ten cleanup receipts remain authoritative.

  Full combined runtime/gates, quiet-host native qualification, corpus
  replay/dispositions and fresh inventory/glossary validation remain open.
  Source integration is staging, not accepted into trunk. Only the
  comprehensive redaction audit is deferred. Nothing has been pushed.

- Coordinator checkpoint (2026-10-23, complete facade/XLSX passes and original baselines):
  Main integration reaches `a658b88ee3`. The complete all-feature facade
  library passes all 1,002 tests on diagnostic source `bf09ddfd9c`, closing
  the nine v12 failures. The full all-feature XLSX library passes 1,871 tests
  on `a58dbaceb3`; representation-aware normalization, atomic joins and typed
  failure sources passed independent review. Unbounded rich setters avoid
  eager source reads, and join refusal restores both edits without cloning
  whole edits or retained action payloads. Record 0362upe retains source,
  commands, hashes and durable external log mirrors; native performance
  qualification remains open.

  Diagnostic composition reaches `47caed4627`. The complete DOCX library
  passes all 1,519 tests on this source. A preceding full run exposed bookmark
  loss risk in whole-cell replacement. The corrected shared scanner keeps
  whole-cell operations strict, while paragraph writers, readback and replay
  preconditions consistently preserve markers. New direct, nested-cell and
  content-control proofs cover changed text, durable replay and inverse
  restoration. The workspace all-target/all-feature Clippy run is active.
  Signed DOC fallback now preflights the original CFB under caller limits
  before private signature removal; exact DOC/XLS no-ops and cancellation
  pass the complete facade suite. These source repairs are reviewed and
  composed, not accepted as performance exceptions.

  XLSB now passes 771 library tests with one remaining transitive replay
  fixture assertion: both retained original content-types images can differ
  from the effective publication after different edit generations. Its owner
  is running the focused cases while preserving exact archive, effective
  graph and untouched payload/relationship checks. Records 0362up/0362upa
  retain all complete runtime failures and successes. The full assembled
  gate remains required.

  DOCX Ink matched correctness and allocation preflight passes all six routes;
  candidate allocations are lower or equal, with native timing unclaimed.
  ODP's original e4a baseline now builds and passes its six untimed legacy
  proofs on the frozen fixture. Its adapter preserves the actual legacy
  materialize/write and plan/publication lifecycle. Shared timing routes are
  open, read, write and append; candidate-only patch/inverse/scoped correctness
  remains strict. A final candidate build and paired qualification are pending.

  IWork qualification now preserves the original constructor/owner baselines
  and the exact seed registry resolutions. Initial metadata preparation
  preserves all common identities and admits only the reviewed SmallVec
  addition; release builds remain locked. Producer, validator, runner and
  collector bindings passed review. Fresh matched builds and untimed
  preflights are active. W8's baseline tool now verifies the complete pristine
  f2 source tree and uses only four probe consumer adaptations; all 109 route
  bindings retain native/facade workloads and publication boundaries. The
  original-baseline build is active; no new timing result is claimed.

  The final inventory runner is merged. It captures all 68 direct library
  packages with all features and hidden items, validates source/lock/toolchain
  stability around captures, records capped execution and peak RSS, uses
  external temporary storage, and leaves disposition decisions for review.
  Fresh final-source inventory and glossary validation have not run yet.

  The eleventh cleanup removes one clean inactive merged worktree without
  force, retaining its branch and quarantining its 139,553-byte lockfile.
  The removed checkout represented 6,257,821,238 apparent bytes. An abandoned
  1,060,162,316-byte ODF partial checkout was relocated with a complete external
  manifest and restore helper; that relocation is not counted as disk space
  freed. Protected targets, measurements and active leaves remain untouched.

  Final runtime/gates, quiet-host native qualification, corpus dispositions
  and fresh inventory/glossary validation remain required. Source integration
  is staging, not accepted into trunk. Only the comprehensive redaction audit
  is deferred. Nothing has been pushed.

- Coordinator checkpoint (2026-10-23, shared types and twelfth cleanup):
  Main integration and diagnostic composition reach `5dae0abf78`. The earlier
  XLSB return is closed: all 772 all-feature library tests pass on
  `73a38420b8`, with logical transitive member comparison scoped by the existing
  OPC contract and same-owner/unreferenced exact checks retained. The full
  facade, DOCX and XLSX passes recorded above remain source-pinned evidence;
  the complete final gate is still required.

  The final-sweep audit found changed-only Diagnostics and plain Commit copies
  still required to migrate under 0351 section 2.2. PPT/PPTX, XLS/XLSB, Pages
  and ODF now use the shared core shapes; genuinely richer diagnostics and
  PPTX's extra opened snapshot state remain custom. Pages metadata replay now
  returns the Commit triad, and its consumers and replay/no-op tests migrate
  with it. ODI flat planning/source helpers move to Patch. Records 0362us and
  0362usa–0362usd retain the audit and review boundaries. ODM's changed-byte
  comparison now occurs during commit rather than lazily in diagnostics; its
  final native qualification remains required, with no performance waiver.

  The glossary now enforces 30 nominated shared identities through rustdoc
  nominal IDs, including alias/re-export resolution and fail-closed missing,
  ambiguous, cyclic and unrelated cases. The Pages replay rule and ODI helper
  bans survive conflict resolution. Glossary/inventory tool tests pass 54/54.
  The shared Commit Debug now prints only content-free diagnostics, avoiding
  a newly exposed payload route through derived format Debug implementations;
  six core edit tests pass, including payload-formatting rejection. This is
  a narrow obligation of the changed API, not completion of redaction audit 22.

  The complete composed Clippy v23 run returns three spelling-only lints in
  XLSX facade admission and a PPTX test seam. These are fixed. The v24 complete
  all-target/all-feature workspace pass is active on `5dae0abf78`. A separate
  owner-replay sweep found ODG's three apply_joined plan-publication methods
  still returning Package; that work is returned for migration into Edit
  staging, preserving bounded composition, policies, source checks, aggregate
  patch/diagnostics and one publication.

  The pristine f2 W8 baseline build now passes all 109 untimed route checks
  and its fail-closed endpoint verification; receipt 43c9fc24fd is reviewed
  and merged. IWork's harness now admits only source-declared inactive lock
  edges and rejects undeclared edges even when both endpoint locks and seed
  contain them. Its 20 tests pass. The reusable Pages fixture example builds
  without the previous inclusion errors. Earlier build/preflight packets
  remain external; fresh coherent final-candidate qualification waits for
  source freeze. No new native timing result is claimed.

  Twelfth cleanup removes four clean, inactive, merged source worktrees
  without force, retaining their branch refs. Their apparent checkout size
  totals 25,037,023,460 bytes. Twenty-five ignored generated files totaling
  8,173,100 bytes are hash-verified in external quarantine with restore
  instructions. Protected targets, measurements and incomplete checkouts
  remain untouched; the coordinator verified every retained file identity.

  Source integration remains staging, not accepted into trunk. ODG completion,
  the full 53-check gate, all 68 final inventory captures and dispositions,
  corpus replay, glossary closure and native qualification remain open.
  Only comprehensive redaction audit 22 is deferred. Nothing has been pushed.

- Coordinator checkpoint (2026-10-23, composed Clippy and thirteenth cleanup):
  Main integration reaches `5ee8387c2c`. Complete all-target/all-feature
  workspace Clippy with warnings denied passes on clean source
  `8225b82dde7bac272e9e77789a78751ae07aaa2a`, after the six v24 consumer
  lint sites were migrated in `0bed2f0710`. The run completed in 199.602
  seconds with a 4.40 GB external target. Record 0362us and the reduced
  v25 receipt retain the source, command, lock and external log identity.
  This is a composed workspace pass, not the complete final 53-check gate.

  The reviewed proposal-only replay registration helper is merged at
  `8225b82dde`. It checks concrete/default aliases before resolving declared
  generic bindings, including reordered parameters; every proposed row
  still needs nominal result and owner proofs. It never changes or accepts
  policy. The glossary and final-inventory tool suites pass 60 tests. Fresh
  all-68 rustdoc capture, policy registration and row dispositions remain
  required after the remaining source correction.

  ODG's replacement of the three owner apply_joined plan-publication methods
  remains returned for correction. Source `15f677ed52` and receipt
  `c5e14c8d7f` pass 48 focused tests and strict checks, but independent review
  still finds replacement-segment name validation, intermediate replacement
  policy, and projected/inherited style-dependency gaps. The worker is
  preserving ordered operations within one sub-edit while fixing conflicts
  between independent work; a broad mixed-operation capability ban is not
  accepted. No ODG joined migration has been merged into main integration.

  Thirteenth cleanup removes five more verified clean, inactive, merged
  source leaves without force. Their apparent checkout size totals
  19,028,805,424 bytes. All branch refs remain; ten generated files totaling
  739,055 bytes are hash-quarantined with restore instructions. The
  coordinator independently verified file identities, ancestry and removed
  path absence. Receipt `db36955293` records the operation. Targets, native
  binaries, measurements and incomplete checkouts remain protected.

  Source integration remains staging. Final inventory/glossary closure,
  the full gate, corpus replay and native qualification remain required.
  Comprehensive redaction audit 22 is the only deferral. Nothing is pushed.

- Coordinator checkpoint (2026-10-23, lazy edit entry and fourteenth cleanup):
  Main integration reaches `fb95c29e64`. The ODG joined migration is now
  reviewed and merged at `c3f56b7a9d`, retaining all correction records.
  Canonical effect checks, ordered operations within a sub-edit, projected
  style dependencies and active replacement policies preserve composition
  capability. Lazy content/styles state removes entry-time XML copies,
  including unused copies for pure package replacements. Final source
  `05c15847c1` passes 23 library tests, 155 integration tests and one doctest,
  with strict checks. This is source validation, not native qualification.

  ODB/ODM lazy entry is reviewed and merged at `8779d059ad`. ODB borrows
  source XML until mutation; ODM borrows projected titles and reads raw
  meta.xml only for a changed-title commit. Full suites pass 106/58 regular
  tests respectively plus one doctest each; strict Clippy passes. The
  successor receipt identifies its evidence as reduced terminal-output
  summaries. It explicitly records the deferred ODM error boundary and the
  ODB helper paths that can materialize before discovering a semantic no-op.

  Record 0362ug tracks required cost corrections under 0351 sections 2.1.8
  and 2.6.3. Signature conformance does not prove O(1) edit construction.
  Confirmed eager source copies/scans remain under parallel correction in
  ODP, ODS, DOC, XLS, XLSX, Pages, Keynote, Numbers, PPT and in-memory PPTX.
  ODS library tests pass 700 cases; its full-suite fixture availability is
  being checked against the sparse checkout before validation continues.
  XLSX changes and fallible-accessor consumer migrations are committed for
  review, with Cargo validation pending a build slot. The accepted PPT digest
  work remains intact; its new constructor finding is a separate obligation.
  No constructor cost is waived and no source freeze is claimed.

  Safe-gate readiness now covers inherited cap-marker clearing, registration
  of inventory tests and preservation of child logs after resource failures.
  Separate qualification tooling passes 63 tests and the iWork tooling passes
  20. These are not the final 53-check gate or the all-68 rustdoc inventory.
  The ODF shared-triad probe now checks real outputs and triad behavior, but
  review returned its timing boundary, manifest fidelity proof and paired
  qualification protocol for further correction. The ODG performance lane
  also still needs executable paired qualification. Final native endpoints
  remain held until the required source corrections are assembled.

  Fourteenth cleanup removes the clean, inactive, merged ODG source leaf
  without force, retaining its branch. Apparent checkout size was
  6,263,368,185 bytes; three generated files totaling 205,668 bytes are
  hash-quarantined with a restore helper. The coordinator verified their
  identities, ancestry and removed-path absence. All external targets and
  measurement evidence remain protected; new correction leaves use sparse
  checkouts to avoid copying unrelated evidence.

  Source integration remains staging. Full gates, inventory dispositions,
  corpus replay, glossary closure and native qualification remain required.
  Comprehensive redaction audit 22 is the only deferral. Nothing is pushed.

- Coordinator checkpoint (2026-10-23, pinned ODS closure and consumer review):
  Main integration reaches `f741d67f46`. ODS lazy entry and cheap
  limits/context admission are independently reviewed and merged at
  `f144cd6029`. Source `2e2f902179` passes 1,767 tests across 119 successful
  result blocks, including 702 library tests, and strict all-target/all-feature
  Clippy. The coordinator verified the raw logs and actual root configuration.
  Sparse fixture expansion had excluded the pinned toolchain and Clippy
  configuration; both were restored and final checks rerun under Rust 1.95.
  The earlier Rust 1.98 run is historical evidence, not acceptance. All 37
  fixture paths were already tracked; no fixture or large artifact was added.

  The completed constructor audit also confirms DOCX and XLSB source-sized
  entry work; records 0362ugl/0362ugm are assigned alongside the existing
  correction lanes. ODC, ODI, OTH, formula and ODT entries were checked against
  their actual separate owners and Arc-backed state. The earlier mistaken
  audit assumption that those formats all route through ODT was corrected
  before the clear disposition was recorded. Bounded raw codec drafts remain
  distinguished from archive Package entry under the existing contract.

  Review continues to return concrete capability and cost defects. Numbers'
  deferred chart refusal now preserves its host fallback without cloning the
  full grid. Keynote's standalone profiling consumer is migrated. ODP removes
  its remaining XML entry scan and preserves typed errors on read projections;
  final-source validation is active. PPTX now defers OPC graph/catalog cloning,
  with a remaining refusal-test consumer correction. DOC's scoped tests pass,
  but final review still requires omitted facade/profiling consumer migrations
  and preservation of ordinary Edit auto traits with profiled callbacks.
  These source leaves are not yet accepted into integration.

  ODF paired qualification now has executable preparation and orchestration,
  but review returned build-receipt trust, actual source/config verification,
  complete fidelity proof and allocator-observer boundary gaps. ODG's separate
  executable lane is being implemented. Neither source-only tooling nor debug
  correctness is credited as native performance qualification. Four capped
  Cargo slots are rotated among workers; final candidate freeze remains held.

  Source integration remains staging. The final full gate, all-68 inventory,
  corpus replay, glossary closure and native qualification remain required.
  Comprehensive redaction audit 22 is the only deferral. Nothing is pushed.

- Coordinator checkpoint (2026-10-24, DOC/ODS integration and constructor returns):
  Main integration reaches `0c2357755f`. Reviewed DOC lazy entry and its complete
  consumer migration merge at `cf7def80df`. Ordinary edits preserve Send/Sync;
  profiled edits retain borrowed observers across failed opens and honor the
  selected protection policy. Facade budget propagation and standalone
  profiling consumers are migrated. The coordinator verified 1,445 default
  tests across 50 result blocks, 1,223 profiled tests, strict default/profiled
  Clippy, the DOC facade check and both profiling-tool checks. Record 0362ugd-a
  binds clean tested sources and the exact committed two-file consumer overlay.

  ODS resource admission follow-up 0362ugca merges at `3676e21cfa`. The resource
  census stays cold through open/edit/drop/shared handoff; first staging,
  commit and replay retain checks before candidate retention/publication.
  Final source `1fdfcdcdba` passes 1,768 tests and strict all-target/all-feature
  Clippy under pinned Rust 1.95. The coordinator verified all raw-log hashes,
  result blocks and actual root configuration. These receipts establish source
  correctness, not native performance acceptance.

  The ordinary XLSX audit found additional eager pivot, graph, object and
  catalog constructors. Records 0362ugn–0362ugq divide those corrections among
  separate workers. The source-backed packet at `02684ae108` passes 2,542 tests
  and strict Clippy after the slicer failures were reproduced on the unchanged
  baseline and repaired. Later ordinary catalog changes have separate review
  and validation pending. DOCX's full lazy-entry source slice is committed for
  review. No source freeze or constructor-cost exception is claimed.

  Review continues to return concrete defects. Numbers still needs retained
  selector/traversal budget accounting and standalone fuzz consumers. Keynote
  needs selector allocation admission and a minimal standalone lock update;
  broad unrelated dependency upgrades are rejected. Its IWA fixture failure
  was reproduced on the clean baseline and narrowly repaired. PPT's native
  source fence now has a deterministic in-flight mutation proof, including a
  positive control; runtime validation is queued. Pages, PPTX, XLSB and ODP
  have source corrections under review or capped validation. Four owned Cargo
  slots remain the limit; foreign jobs are not disturbed.

  Fifteenth cleanup removes the clean, merged DOC source leaf without force,
  retaining its branch. Apparent checkout size was 314,138,078 bytes; its
  139,591-byte ignored lockfile is hash-quarantined with restore instructions.
  The coordinator verified the receipt and file identities. Targets and logs
  remain external and unchanged. After stale owned searches exited, a fresh
  audit also permitted ODB/ODM removal as cleanup sixteen: 6,263,385,091
  apparent bytes, with three ignored files totaling 205,680 bytes quarantined.
  Its branch, target and logs remain; all receipt identities were verified.

  Source integration remains staging. Final full gates, all-68 inventory
  dispositions, corpus replay, glossary closure and native qualification remain
  required. Comprehensive redaction audit 22 is the only deferral. Nothing is
  pushed.

- Coordinator addendum (2026-10-24, ODP admission closure):
  Reviewed ODP source `363a2497f9` and receipt `88c77dfe11` merge at
  `d316d04291`; main integration reaches `1b68709e0b`. Source construction
  stays cold, security admission precedes root/PI checks, and an empty splice
  still performs first-use admission. All 448 tests across 38 suites, strict
  all-target/all-feature Clippy and the all-target/all-feature facade check
  pass. The coordinator verified the five raw logs and pinned configuration.
  The receipt preserves the unrelated XLSX formatting overlay explicitly;
  only committed ODP changes were merged. Native acceptance remains open.

  Final DOCX SVG correction, Pages selector-allocation correction, Keynote
  selector retention/minimal lock correction and XLSX pivot changes clear
  scoped source review; their final runtime validation remains queued. Numbers
  was returned for preserving its budget ledger across failed first-use
  selection, and XLSX catalog work for early aggregate admission, effective
  edit limits and cold initial-batch staging. Source integration remains
  staging until the complete gates and required qualification pass. No push.

- Coordinator checkpoint (2026-10-24, XLS integration and composed validation):
  Main integration reaches `ab366e3b10`. Reviewed XLS lazy-entry/source-fence
  correction `b8ed2d46bb` and receipt `8c2b79f7c5` merge at `50d4f4b88d`.
  Content-matching cross-owner replay retains the recipient's deferred fence;
  admission and later guards preserve typed SourceChanged precedence for an
  in-flight original mutation. Chart/pivot failures retain retry state.
  The coordinator verified 1,870 passing tests across 100 suites, strict
  Clippy, all-feature facade checks and four metadata tests. The existing
  ignored test writes an evidence workbook. No native timing is claimed.

  The coordinator composed Pages/Numbers/Keynote with current integration in
  the existing Pages leaf at `3ae5fdf0ff`, then validation caught and corrected
  a Footnote Arc type error, two test fixtures/names and a Keynote selector-kind
  mapping. Source review confirms the Arc fix preserves sharing and error
  behavior. Pages and Numbers checks pass at their recorded source boundaries;
  Keynote's all-target suite also passes. The remaining iWork checks, explicit
  doctests and standalone consumers continue under record 0362ugr. These
  assemblies are not yet accepted into main integration.

  XLSX source/catalog/batch work, pivot, graph and object corrections are
  composed with current XLS integration at `7ed6fb0ca7`. Effective nested pivot
  limits now retain their typed error; the regression distinguishes the exact
  oversized nested part from earlier candidate serialization. The omitted
  tracked Cargo alias configuration was materialized and all root configuration
  verified before the combined build. Validation is repairing concrete compile
  omissions in the composed families, then rerunning the complete relevant
  suites, strict checks, consumers and doctests under record 0362ugs.

  ODF shared-triad tooling `7f85a0dd5d` clears scoped semantic/provenance review;
  18 ODF-specific Python tests pass independently. Its Rust probe is included
  in composed facade compilation. Native measurements remain unrun and final
  endpoints unresolved. ODG tooling is still being corrected and reviewed for
  actual source-byte closure, effective build environment and final target
  size enforcement; passing Python tests alone do not qualify native results.

  Four build slots rotate across composed iWork, composed XLSX, XLSB and PPTX.
  PPT and DOCX validation remains queued. An attempted XLS worktree cleanup
  stopped because a foreign process holds a fixture; no files or targets were
  removed. Final full gates, all-68 inventory, corpus, glossary and native
  qualification remain required. Only redaction audit 22 is deferred. No push.


- Coordinator checkpoint (2026-10-24, PPTX integration and seventeenth cleanup):
  Main integration reaches `0223f55f42`. Reviewed PPTX source `c10d6c6539`,
  comment correction `fa63bdd22e` and receipt `e59c72184d` merge at
  `a3eb017da2`. The coordinator verified 1,837 regular tests, 48 doctests,
  strict Clippy and 11 facade tests, including exact raw-log identities and
  pinned configuration. Candidate reopen preserves typed errors; bounded
  fallback is limited to recompressing read-limit failures. Output policy and
  memory admission remain distinct. The earlier uncapped diagnostic is
  documented separately and receives no acceptance credit.

  ODG paired qualification tooling `7c7d7a54fb` clears its returned source,
  compiler-environment and final target-limit checks and merges at
  `a41b437ca8`. Its 26 Python tests pass independently. Rust adapter builds
  and native qualification remain required; no timing result is claimed.
  The final gate runner's post-exit process/target checks and attempt-specific
  retry logs are assigned for correction before final gate execution.

  The composed iWork suites, 96 doctests and strict four-crate Clippy pass
  through their recorded sources. Facade and standalone checks continue.
  Scoped review returned the modified Pages table-name Debug field for
  content redaction and its lazy resolution error path for preserving the
  charged budget across retries. These are required corrections, with no
  expansion of deferred audit 22. Composed XLSX, XLSB and DOCX validation
  continues; PPT receives the next available capped build slot.

  Seventeenth cleanup removes the clean, merged ODS leaf without force,
  retaining its branch and quarantining its sole generated lockfile with
  verified hashes and restore instructions. Apparent checkout size was
  440,249,548 bytes. The external target and logs remain untouched. Foreign
  searches still hold the completed XLS/PPTX leaves, which remain intact.

  Source integration remains staging. Final full gates, all-68 inventory,
  corpus replay, glossary closure and native qualification remain required.
  Only comprehensive redaction audit 22 is deferred. Nothing is pushed.


- Coordinator checkpoint (2026-10-24, iWork/XLSB merges and composed strict check):
  Main integration reaches `1705371eb5`. Composed iWork source `a0bf134246`
  and receipt `4441c86bbe` merge at `9ea602e0e1`. The coordinator verified
  4,695 package tests, 96 doctests, 50 selected facade tests, strict Clippy,
  operations targets, standalone fuzz and unified-batch checks, with exact
  source boundaries, logs, locks and root configuration. Pages name,
  appearance and merge edits preserve their charged budgets across failed
  resolution and reuse them on retry; the changed Debug surface is redacted.
  A missed fallible Keynote fuzz getter was migrated before the final check.

  XLSB source `51c8628446` and final receipt `16ba8da04f` merge at
  `0d35a83a31`. The coordinator verified 1,355 package tests including 42
  doctests, strict checks through the facade/Python bindings, 126 operations
  tests, 23 facade tests and nine standalone CRUD harness tests. Record
  0362ugm binds all closed logs and distinguishes the standalone run's
  documentation-only overlay from exact-clean runs. No native claim is made.

  The composed strict consumer check found a missed PPTX find/replace benchmark
  getter, fixed in `1f7c188087` under 0362ugka. The gate runner's detached-child,
  retry-log and vanished-PID corrections are reviewed and merged, ending at
  `f9cf4aa6d5`. The rerun passes at `0f007c4d74`: all-target/all-feature strict
  Clippy for CLI, operations, facade and Python, with no owned build processes
  remaining and stable target size 4,706,526,684 bytes. Its reduced receipt
  binds the actual source, command, external logs, plan/state and policy.
  This is one selected step, not the final 53-step gate.

  The actual XLS metadata call path exposed three omitted cold-entry families:
  metadata, property sets and sheet features. Required correction 0362ugea
  covers their deferred state and every consumer; other public XLS Package
  edit entries checked are already lazy. The follow-up also retains the
  signed no-op/changed-write contract. Its fresh leaf now owns the released
  coordinator build slot; PPT, XLSX and DOCX occupy the other three slots.
  This constructor debt is not accepted as an exception.

  Cleanup eighteen removes the merged PPTX leaf (362,854,381 apparent bytes).
  Cleanup nineteen removes the merged Keynote, Numbers and Pages leaves
  (1,237,053,440 apparent bytes). Branches, hash-quarantined locks, restoration
  instructions and all external targets/evidence remain. Foreign searches
  still hold XLS/XLSB leaves, which remain intact. Receipts were independently
  verified and committed.

  Source integration remains staging. Final source freeze, full gates,
  all-68 inventory, corpus replay, glossary closure and native qualification
  remain required. Only comprehensive redaction audit 22 is deferred.
  Nothing is pushed.


- Coordinator checkpoint (2026-10-24, XLSX/DOCX/PPT composition):
  Main integration reaches `5b06b2136a`. Reviewed composed XLSX source
  `784d0ea83e` and receipt `7201866b32` merge at `3d129677ce`, closing the
  source-backed, catalog, pivot, graph and object constructor work. The
  coordinator verified 2,558 package tests, 36 doctests, strict Clippy,
  standalone/fuzz consumers and three planning-guard correctness routes.
  First-use errors and retries remain outside the constructor measurement;
  the unknown-attribute fixture was corrected to an unsupported child element
  after confirming that attributes are intentionally preserved. All evidence
  retains exact source/overlay boundaries; no native claim is made.

  DOCX final source `0cf1a17174` and receipt `3a6c56958c` merge at
  `5ce9829784`. Review returned and then verified the signed empty-property
  correction: valid targets preserve exact signed bytes; invalid targets keep
  their target error. The coordinator verified 2,355 package tests, 170
  doctests, 739 operations tests, 23 facade tests, strict checks and the
  migrated fuzz/performance consumers. The recorded 31 ignored doctests
  remain explicit.

  PPT source `ca125d2ba1` and final consumer/binding receipt `70a9b57622`
  merge at `127c26d59f`. The coordinator verified 1,568 package tests,
  22 doctests, 71 selected operations tests, strict checks and standalone
  consumer compilation. The ODG candidate probe compiles; refreshed probe
  identities pass all 26 build-free protocol tests. No native qualification
  is inferred from compilation.

  A composed strict check returned the ODF example's fourteen-argument receipt
  helper. Private identity grouping fixes the lint without changing proofs,
  fields or measurement boundaries (0362ugsa). The rerun passes at
  `ecf598b637`; after the DOCX/PPT merges, the same all-target/all-feature
  CLI/operations/facade/Python strict check passes at `99c875ca41` in
  145.777 seconds. Post-exit checks find no owned builders and a stable
  4,706,726,912-byte target. Neither selected run is the full 53-step gate.

  The finite W8 error audit returned XLS feature mapping and the PPTX
  slide-removal override. PPTX correction 0362ugt merges at `c1eda2fdf6`: the
  existing FinalSlide InvalidValue kind now reports invalid_value, with both
  capability and registry goldens and the generated digest synchronized.
  Focused tests, strict PPTX-only Clippy and both static OfficeCLI checks pass.
  The remaining XLS correction 0362ugea has cold property/feature entry,
  signed exact no-ops and corrected facade classification. Its final render
  and candidate-admission source fence remains under review: an early error
  return must not bypass SourceChanged precedence. This is required work.

  Cleanups twenty through twenty-three remove nine completed leaves after
  fresh ancestry/cleanliness/process checks. Branches and external evidence
  remain; ignored lockfiles and the local XLSX target are hash-quarantined
  with restore instructions. The coordinator independently verified every
  preserved file, including all 2,937 entries from cleanup twenty-two.
  Reviewers stopped only their own stale Git jobs before their leaf removal.

  Source integration remains staging. Final source freeze, full gates,
  all-68 inventory, actual corpus replay, glossary closure and native
  qualification remain required. Only comprehensive redaction audit 22
  is deferred. Nothing is pushed.


- Coordinator checkpoint (2026-10-24, XLS closure and acceptance restart):
  Reviewed XLS source `3ae7b83140` and receipt `9d38e2c4f2` merge at
  `2dfc4f134d`. Metadata, property sets and worksheet features retain cold
  edit entry, signed exact no-ops and typed refusal mapping. Both candidate
  admission and worksheet rendering now preserve SourceChanged precedence;
  each has stable-error and post-leading-source-change controls. The
  coordinator verified all source-bound test receipts and retained all DOC,
  XLS and PPT fallible metadata benchmark consumers.

  The first full gate at `2dfc4f134d` stopped at formatting. Ten files were
  formatted, and the composed all-target/all-feature strict CLI, operations,
  facade and Python check passes at `97817cb75f` in 167.272 seconds with a
  stable 4,706,734,564-byte target and no remaining owned builders (0362ugu).

  Native protocol checks returned stale ODF/W8 probe and IWA fixture source
  bindings. Reviewed corrections preserve the actual fixture generator and
  workload. W8 also required a pristine-f2 adapter for three deferred XLSB
  selector calls; the corrected adapter is byte-identical to the previously
  reviewed baseline consumer. Its test no longer requires historical Git
  objects. All 65 general native protocol checks and 20 iWork checks pass at
  frozen candidate `b6732d02b241705ba153c1a2cb9ff03bf0a13f93`.

  The full 53-step gate is running on that fixed candidate in the diagnostic
  worktree. Formatting and combined strict consumers pass; remaining steps
  are pending. ODG, iWork and W4b/W5 native endpoint preparation/builds have
  three capped slots, with timing held until the host is quiet. The fixed ODG
  baseline accepts the candidate lock; all 228 external dependency rows
  match, independently verified. No native acceptance is claimed.

  Cleanups twenty-four through twenty-six remove seven more merged leaves,
  retaining branches, external targets/evidence, and hash-quarantined ignored
  artifacts. The coordinator verified all 5,266 cleanup24 payloads and all
  three XLS payloads, plus seven final binding-leaf files. A restore-layout
  mismatch was corrected before cleanup25
  acceptance. Story-text and ODraw leaves remain because of live foreign
  handles and missing integration ancestry, respectively.

  Source integration remains staging. Completion still requires the full
  gate, all-68 inventory, actual corpus replay, glossary closure and quiet
  native qualification. Only comprehensive redaction audit 22 is deferred.
  Nothing is pushed.


- Coordinator checkpoint (2026-10-24, feature correction and frozen acceptance):
  The full gate on b673 passed its first 17 steps and returned an unused ODF
  helper in the ODG-only profile. Correction `e1576fa6dc` matches the helper
  to its production callers and ODB-only tests; all eight remaining strict
  feature profiles pass. Record 0362ugx retains both failed attempts and the
  passing selection. The ownership guard also correctly rejected an iWork
  fixture generator that used the gate checkout; its owner validation remains
  content-only diagnostic evidence, with an explicit operational rejection.

  Reviewed historical W4b/W5 probe adapters `ccc536d524` merge at
  `f6fe430a4d19d94aa360c0e9ed03ebd757265225`. They preserve all workloads,
  including both DOC publication controls, and translate only the historical
  selector/options/result API differences. All 21 protocol tests and the
  historical consumer compilation checks pass. Record 0362ugy binds exact
  source, overlays, locks and commands; no diagnostic compilation is treated
  as native qualification.

  Production/protocol source is frozen at that f6fe merge. The fresh full
  53-step gate has passed formatting, combined strict consumers and every
  per-format strict profile (25 steps); remaining tests and documentation are
  running. Receipt-only commits retain that exact acceptance source. Required
  production or protocol fixes will trigger a new freeze.

  Independent iWork constructor/owner and DOCX Ink diagnostic verification
  finds no artifact mismatch. Constructor variants validate 15 routes each;
  owner variants validate four each, subject to the operational rejection
  above. Ink native/allocation untimed outputs match. ODP diagnostics exposed
  differing dependency resolutions, so a fresh e4-backed projection now
  matches f6fe across all 204 external lock records and 157 ODP external
  dependency rows. The coordinator verified the matched sets and corrected
  provenance identities. Quiet native execution remains pending.

  Cleanup twenty-seven removes the merged W4b/W5 leaf after its reviewer
  stops two owned stale searches. The branch, all native endpoints, targets
  and logs remain. Four ignored files are hash-quarantined with a restore
  helper; the coordinator independently verified every payload and retained
  source ref.

  Source integration remains staging. Full gate completion, all-68 inventory
  and dispositions, actual corpus replay, glossary closure and quiet native
  qualification remain required. Only comprehensive redaction audit 22 is
  deferred. Nothing is pushed.




- Coordinator checkpoint (2026-10-24, acceptance repairs and source-bound evidence):
  The f6fe full gate passed 25 steps before its operations tests exposed the
  XLS duplicate source check, stale fidelity/CRUD reports and encrypted DOCX
  closure admission failure. Its session then ended without a finalized
  operations result. Record 0362uma retains the interrupted attempt and
  subsequent selected checks; none is relabeled as a full pass.

  Reviewed inventory coverage source `668249cebf` merges at `e7f19874c1`.
  Policy-declared immutable owners now participate even when their names do
  not match the old heuristic; missing or ambiguous owners fail closed.
  The new disposition validator checks exact row coverage and structure,
  while semantic evidence still requires review. CI runs both new test
  modules. The all-68 final capture and dispositions remain outstanding.

  Encrypted DOCX source `6b05b8063c` and locked receipt `696352d066` merge at
  `dd78c02a23`. The O(1) publication-policy check refuses an unlocked
  encrypted document before invoking its edit closure, preserving signed
  no-ops and source fences. Both locked feature selections pass (22 and 20
  tests). The original record's lock explanation and size transcription are
  explicitly corrected. Unix socket fixture source `35f47e34dc` merges at
  `250769e1dc`; all three tests pass with a 167-character TMPDIR.

  Corpus source `2510bf5b40` merges at `c13737083e`. Every one of the seven
  fidelity and five CRUD changes was reviewed. Both flat ODP fixtures prove
  a real rename followed by exact inverse restoration; XLSX's second run
  supplies the intended sibling-pair coverage. Five refusal-detail rows only
  change the fixed display wrapper. The unblessed report rerun opens 1,020
  files and reports zero unreached claimed CRUD cells. The fidelity TSV's
  empty final fields are retained as data.

  Cleanup thirty subsequently found an undeclared local corpus target above
  20 billion apparent bytes. The unblessed/focused results retain their
  content evidence, but their claimed bounded-build provenance is returned
  for correction. Reviewed receipt correction `36cf4e3dad` closes that
  return with fresh, explicitly unblessed commands: all 1,020 corpus files,
  zero unreached cells, and all 7 ODP plus 6 XLSX focused tests pass. Both
  commands prove the memory shim, exact source/lock/environment and external
  target; its 3,483,689,639 bytes remain stable before and after. Earlier
  receipts remain preserved; the oversized cache was quarantined and later
  pruned only after the specific regeneration/preservation audit below.

  Selected f6fe diagnostics pass 332 core tests, the iWork leaf API check,
  the non-iWork gate and static corpus dispositions. Facade testing returns
  only the repaired socket fixture failure (208 passed, one failed). The
  1,360-test tooling run returns stale boundary rules, ODG guard-literal
  scanner matches and unclassified diagnostic evidence. Reviewed scanner
  and evidence corrections merge at `ec0d6bb0db` and `2b88f4ad75`;
  boundary correction `5a8cfa93c1` now passes independent review. Its
  guarded deferred selectors preserve O(1) entry and require typed error
  propagation before staging; the shared Pages Diagnostics rule rejects
  duplicates and aliases. All 1,144 boundary tests pass, and the CLI checks
  68 packages and 290 declarations. Its 11 legacy-host edges remain explicit;
  this packet alone does not establish final inventory closure. Reviewed XLS source
  `dfe8281fb3` and receipt `3ddd72e97e` are now merged: a canonical shared
  catalog preserves O(1) adapter entry, direct native source checks and
  trailing-fence precedence without cloning eager sheet names. All 1,907
  native tests/doctests pass; the two existing ignored evidence/example
  cases remain explicit. Catalog conversion and Arc costs still require
  native qualification.

  W8 protocol source `e12b57d624` merges at `cf8e83e716` after repeated
  review returns closed artifact and mutable-validation gaps. All 39 unique
  W8 tests pass, including the complete pair-runner path. The explicit
  policy preserves pristine f2, pins the reviewed local graph changes and
  rechecks source/build/validation artifacts before timing; ordinary equal-
  projection checks remain strict. Its f6fe candidate is diagnostic until
  the new final freeze. The composed inventory, ODG, glossary and matrix
  tooling selection also passes all 113 tests at that merge.

  Reviewed gate packet `8a795fb769` prepares a separate next runner under
  0362uva. It retains all 53 steps and registers both inventory modules and
  all three W8 modules, preserving the original nine tooling modules once
  each. Independent verification confirms the one-line command change,
  unchanged wrappers, process/cap checks and active runner hash. Adoption
  waits until the currently running diagnostic gate exits.

  Cleanups twenty-eight through thirty-four remove thirteen merged leaves after
  independent payload/ref/ancestry/restore-layout verification. Branches,
  ignored locks, targets and validation evidence remain, including the
  rejected corpus leaf's local cache. The external corpus target's historical
  size snapshot is explicitly superseded by the approved rerun; all 14,437
  cleanup-thirty payloads and 94 hardlink groups verify. Removing 126 exact
  duplicate build executables restores 3,712,220,600 bytes of gate headroom;
  benchmark/profile outputs remain intact. The immutable ODG baseline
  capture completes with all 460,379 entries, recorded externally under
  0362ugy; no build or native timing result is inferred from it.

  A subsequent audited prune releases 20,181,409,792 allocated bytes from
  the rejected corpus cache. All 544 retained wave-thirty files have a complete
  filtered restore route with independently verified modes and dry-runs;
  raw receipts and native/protocol outputs remain. Another 144 exact stale
  example twins release 1,554,094,416 bytes of budget-target headroom while
  preserving all benchmark/profile and failure artifacts. The two large
  deletion receipts stay external with hashes and reproduction instructions.

  A documentation audit reconciles the settled legacy source-retention and
  OpenOptions decisions in design §16.3/§16.6 and removes the stale claim in
  §16.5 that rustdoc gates cover only iWork. Final qualification and inventory
  requirements remain explicit.

  All 14 final tooling modules pass together at `885cdc82af`: 1,413 tests.
  The boundary's 11 legacy-host edges are not item-24 naming exceptions:
  0352 explicitly requires renaming rather than deleting the host, and
  ADR 0028 separately controls its eventual retirement. No supported facade
  or semantic leaf depends on that unpublished host.

  Source integration remains staging. Strict all-target/all-feature checks
  for the 19 native gate packages and ODS pass on f6fe; ongoing native tests
  found a real OPC read-ahead retention failure plus four OPC and eight ODP
  stale doctest consumers. Reviewed ODP repair `21c47647af` now migrates all
  eight examples without changing production logic or assertions; all 19
  doctests and the all-target compiler check pass. Reviewed OPC source
  `cdbee64a36` and corrected receipt `42e3f2314c` release the retained window
  through a weak archive link, preserve live clone ownership and source-change
  precedence, and migrate all four doctests. All 1,207 native tests and 11
  doctests pass; the existing independent ZIP64 fixture test remains explicitly
  ignored. Strict all-target Clippy passes. Initial sparse-checkout failures
  were resolved by restoring tracked fixtures. The configured read-ahead path
  needs its own matched native ABBA comparison under 0362umc; the exact-policy
  path performs no weak upgrade. The completed f6fe native selection records
  201 chunks, with only those 13 repaired failures; its 83 ignored cases stay
  explicit. No later step is claimed as run.

  The clean diagnostic checkout now holds frozen source
  `49c25ef67f62a8bb943e7c6a0bf464e49ae87d07`, with the dependency lock
  unchanged. The reviewed 0362uva runner is adopted from its separate path.
  Isolated workspace documentation passes with warnings denied in 1,057.472
  seconds, with no competing builds and a 4,504,969,096-byte final target.
  All 68 fresh library rustdoc captures subsequently pass on the same source:
  1,094,716,714 JSON bytes, maximum per-command RSS 2,997,526,528 bytes, and
  maximum target size 5,644,910,073 bytes. The coordinator verifies every
  artifact and receipt; raw JSON stays external under the reduced 0362uma
  capture receipt. The final inventory generator then fails on the facade's
  foreign module glob `litchi_core::detection`, before emitting findings.
  The process exits 2 and the run remains explicitly incomplete. Header-only
  output is not zero debt; resolver repair and independent coverage review
  are required before the full gate launch. Every row still needs semantic
  disposition review. Full gate
  completion on this source, all-68 closure, actual OfficeCLI corpus replay,
  glossary closure and quiet native qualification remain required. Only
  comprehensive redaction audit 22 is deferred. Nothing is pushed.

- Coordinator checkpoint (2026-10-24, complete capture and inventory findings):
  All 68 frozen-49c rustdoc artifacts pass independent identity and hash
  verification. The failed aggregate run remains unchanged. Reviewed repair
  `2099c66520` merges at `fae96321e3`: the inventory follows explicitly
  delegated facade roots, requires their JSON inputs, retains local facade
  findings, and removes only exact duplicate rows. The diagnostic rerun
  produces 440 unique rows from the 442 pre-dedup findings. Record 0362ul1
  distinguishes the captured source from the repaired tooling; this is not
  final-source inventory closure or a disposition waiver.

  Reviewed OPC protocol `80983c725d` merges at `945eb1d6b9`. Its exact and
  4 KiB forward-read oracles, bounded builds, source identities and quiet-host
  validation are explicit; no measurement has run. Cleanup thirty-five
  removes its merged leaf with zero payloads and retains the branch.
  Reviewed W8 packet `e814e4512f` merges at `54398757b8`: the candidate
  binding uses 49c and preserves the pristine f2 baseline and historical
  lock. Its analyzer now validates the original-to-final pair mode, with a
  complete 109-route regression in the existing registered test module.
  The coordinator independently hashes the external 6,630,123,520-byte
  candidate archive. All 119 composed inventory, glossary and W8 tests pass
  at `fae96321e3`; the W8 policy remains preparation only, pending the next
  production source and native qualification.

  Cleanup thirty-six removes the two reviewed inventory/W8 leaves, retaining
  both branches and all 15 ignored files (598,662 bytes) in verified
  quarantine. The three leaves removed in this checkpoint are no longer
  active worktrees; current API repair worktrees and targets remain intact.

  Inventory-driven engineering is now split into the six remaining format
  limits entry points, DrawingML and Markdown publication families, nineteen
  public error enums, four iWork wire-request compatibility aliases, and
  replay/type-resolution coverage. The alias repair keeps generated protobuf
  schema names and wire layouts intact. Diagnostic review identifies 262
  replay rows requiring policy registration, nine ODF external generic-commit
  resolution cases and two File owner-registration cases. The File triad's
  `file`/`into_file` names are explicitly required by 0351 §2.6; they are not
  new naming exceptions. Raw-value and transport findings require exact
  semantic evidence before final disposition.

  The full 53-step gate has not been launched on 49c after this capture.
  Required API repairs will change production source, so final capture,
  corpus replay, gate, glossary closure and native qualification must bind
  the resulting assembled source. The documentation pass on 49c remains
  valid evidence for that revision. Only redaction audit 22 is deferred.
  Nothing is pushed.

- Coordinator checkpoint (2026-10-24, reviewed API repairs and semantic returns):
  Staging merges the Markdown shared triad at `f442730bb8`, canonical IWA
  wire-request names at `6ed8035e4a`, font/cache required limits at
  `0be121e437`, and signature policy/raw limits at `e35cede32a`. Every source
  change includes its consumers. Markdown's 70 tests, IWA's 972 library tests
  plus integration binaries, and font/cache/archive's 216 tests pass on the
  pinned 52ed dependency graph. The signature slice passes its focused tests
  on the separately recorded 6d diagnostic graph; final composed validation
  still requires 52ed. No sibling-lock result is relabeled as pinned evidence.

  Reviewed error extensibility merges at `83a0ace78b`: eighteen remaining
  public enums gain `non_exhaustive`, with every known downstream mapping
  preserved and fixed future fallbacks. Together with the font error repair,
  all nineteen identified enum sites are addressed in source. Six affected
  library suites pass 7,007 tests with seven existing ignores on the recorded
  diagnostic graph. Standalone fuzz consumers are source-audited separately.

  The result-contract tool is returned for foreign-coverage and alias-identity
  gaps before merging `435e1547bd` at `f620a8dd26`. It now requires exact
  auditable defining types, rejects duplicate defining graphs and inconsistent
  crate identities, preserves type/macro namespaces, and registers 273 replay
  contracts. Exact AST and wire-value disposition categories merge at
  `1fb0744f82`; they are structural classifications, not broad waivers. All
  fourteen registered tooling modules pass together on `83a0ace78b`: 1,428
  tests. The frozen-49c diagnostic remains separate from final-source closure.

  Source review returns DrawingML before merge: alternate publication methods,
  action-owner entry/replay, compatibility methods and builder placement still
  require correction. The spreadsheet Chart also needs canonical builder
  placement. Dedicated lanes address legacy raw limits/error boundaries and
  the CFB/ZIP canonical limits work scheduled by 0362rd. Record 0362ul5 makes
  these engineering requirements explicit; it does not classify them away.
  Canonical shared value types remain allowed, and raw Reader mutability is
  reviewed against the specific codec contract rather than builder status.

  Cleanup thirty-seven removes three merged leaf worktrees non-forced. All
  eight ignored files (6,677,946 bytes), their modes, branches and validation
  evidence remain; restoration was exercised successfully. The coordinator
  independently verifies payload hashes and merge ancestry. Active repair
  leaves and external targets are preserved.

  Source integration remains staging pending the resulting source freeze,
  fresh all-68 inventory/dispositions, full 53-step gate, real OfficeCLI corpus
  replay and native performance qualification. Only redaction audit 22 is
  deferred. Nothing is pushed.

- Coordinator checkpoint (2026-10-24, chart/raw/ZIP composition):
  Reviewed chart builders merge at `1f2d0509b5`, bounded legacy raw APIs at
  `db433b2474`, and canonical ZIP limits at `0dcc767c68`. All consumers move
  with their API changes. Review returns close typed-error/redaction gaps,
  XLS revision-log allocation before admission, overlooked lazy-cache aliases
  and ODF retained-budget propagation before merge. Record 0362ul6 separates
  each leaf's validation from the still-pending final assembled gate.

  The root dependency pin is now 4100b1fbb272: the sole graph change from
  52ed is the required `litchi-ops -> litchi-vba` edge. All 518 registry
  identities remain unchanged; both inputs are preserved. Historical 49c
  evidence retains its original source and lock. The chart slice reports
  2,675 passing tests, ZIP reports 631 with one existing ignore, and the
  legacy repair's focused native/facade regressions pass. Seventy-seven
  composed tooling tests and the corrected workspace formatting check pass
  before the later raw/ZIP merges; neither substitutes for final acceptance.

  Cleanup thirty-eight removes five merged leaf worktrees non-forced and
  retains all 4,127 ignored files (2,946,217,884 bytes), branch refs and
  restoration evidence. Payload hashes, modes and ancestry are independently
  verified. The quarantined build target is retained, not claimed as freed
  disk space. CFB and DrawingML engineering remain active. The next source
  freeze still requires fresh inventory, the full gate, real corpus replay
  and native qualification. Only audit 22 is deferred. Nothing is pushed.

- Coordinator checkpoint (2026-10-24, exact remaining source roles):
  Independent review proves 38 current raw-reader methods and 11 composition,
  wire, archive and OPC splice descriptors as exact non-lifecycle values.
  Six removed raw-reader aliases remain stale diagnostic rows and are excluded
  from every allowance. Record 0362ul7 binds all 55 inspected old-capture
  identities to source `0dcc767c68`; it does not claim a fresh final capture.

  The validator registration merges at `14161bfd63`. Coordinator conflict
  resolution retains all 34 prior AST/wire identities and the 49 new exact
  identities, with no module, crate, Reader or Edit suffix exemption. All 91
  combined glossary/inventory/disposition tooling tests pass, and independent
  review confirms the existing evidence, one-to-one and deferral checks.

  CFB retains custom policies through facade ingress, layout adoption,
  sequential validation and overlay/inverse composition in its active repair.
  DrawingML owner/replay review passes, while typed producer observations and
  nested DOCX/PPTX classifications are being corrected. Neither lane is merged
  yet. Preparation also checks the missing PPT lazy-entry and XLSX raw native
  execution protocols; no native qualification has been claimed. Final
  source capture, the 53-step gate, real corpus replay and all native lanes
  remain required. Only redaction audit 22 is deferred. Nothing is pushed.

- Coordinator checkpoint (2026-10-24, CFB composition and raw review):
  Reviewed CFB `bf7f48741b` merges at `f1d63257e6`, with four retained-policy
  returns closed. Composed checks expose and repair the independently added
  VBA error matches and a chart-test qualification; the complete workspace
  all-features/all-targets check then passes on the current 4100 lock. Record
  0362zada retains both failures and the passing receipt rather than replacing
  them. Reviewed XLSX raw native protocol `5e6555d840` merges at `95ce396d83`;
  all 37 composed Python tests pass, with no native measurement claimed.

  DrawingML raw engineering is split into chart, diagram, text/media,
  theme/ink and geometry/color/transform/formatcode2 lanes. Fixed static caps
  do not exempt raw codecs from required policies. Text/media is returned
  after review finds allocation-admission and SVG byte-fidelity defects
  despite focused test passes. PPT and legacy native protocols also retain
  concrete fixture/provenance and historical API-adapter returns. Record
  0362ul9 describes the exact scope; no new exception is accepted.

  Cleanup thirty-nine removes three merged leaves non-forced, retaining all
  branches and eleven payloads (566,586 bytes), with restoration and independent
  hash/mode/ancestry checks. Final assembled capture, the full 53-step gate,
  actual corpus replay, glossary closure and all required native qualification
  remain pending the source repairs. Only redaction audit 22 is deferred.
  Nothing is pushed.

- Coordinator continuation (2026-10-24, native protocol composition):
  The PPT source-only qualification packet `4fbfe1205a` merges at
  `e563cc486b`. Independent review verifies nonempty object, movie and
  picture-bullet fixtures, route/digest bindings outside timing and mutation
  rejection for both baseline/candidate normal and allocator binaries.
  All 13 composed Python tests pass. Actual standalone baseline/candidate
  compile checks are now scheduled; no native measurement is claimed.

  The legacy packet `384740be13` is returned before committing its merge.
  Although 64 combined Python tests pass, review of the executable handoff
  finds a stale OLE feature contract, mismatched preparation/runner identity
  digests and missing preparation-path separation. The resolved patch and
  draft record are preserved externally, and the uncommitted merge is
  aborted for a corrected leaf. Required integration tests now cover actual
  manifest loading and the preparation-to-runner identity handoff.

  Diagram and geometry/color/transform/formatcode2 raw packets also remain
  under review returns for admission, XML grammar and retained-policy gaps.
  Text/media's typed admission repair exposes a successful-output namespace
  regression, which must be fixed before merge. Theme/ink remains under
  review with output-preflight and nested-policy corrections underway.
  New paired DrawingML and full OLE owner/history protocols are active work;
  the earlier 17-lane catalog does not close those obligations. Only audit 22
  remains deferred. Nothing is pushed.

- Coordinator continuation (2026-10-24, executable legacy handoff):
  Corrected legacy qualification packet `e8d51bb07a` merges at `4624f4bea6`.
  The shared runner retains both XLSX and legacy receipt validation, with
  mutually exclusive receipt selection and explicit source/target separation.
  All 67 combined tooling tests pass. Independent review also runs the real
  XLS preparation CLI and proves the runner consumes that receipt before
  refusing the absent build receipt. Record 0362ul8a distinguishes this
  preparation evidence from the still-required native builds and execution.

  Text/media `e1d9ab1e29` passes independent source review after collision-safe
  namespace emission and typed admission repairs. Its focused leaf tests pass;
  integration waits for the common DrawingML owner/harness prerequisite and
  the other raw lanes. Geometry/color/transform/formatcode2 receives a further
  return for mapping a scene's output budget onto input parsing; `00fb949e54`
  corrects both transform and text-body readers and adds asymmetric-policy
  regressions. Those new tests still need coordinated Cargo execution.

  Actual PPT baseline and candidate release-build attempts both fail: the
  baseline preparation omits the fixed probe fixture, and the older candidate
  hits the VBA match already repaired in staging. Neither produces qualifying
  native evidence. The preparer repair and properly pinned current-candidate
  build remain required. Full OLE owner/history and DrawingML native protocols
  also remain under engineering review; source-only tests do not qualify them.

  Cleanup waves forty and forty-one remove the merged XLSX protocol and CFB
  leaves non-forced, retaining both branches and all fourteen payloads
  (604,978 bytes). Restoration and independent hash/mode/ancestry checks pass.
  Wave forty's immutable engine receipt records its final-check abort after
  the coordinator advanced staging; separate verification confirms the
  completed removal and preserved payloads. Final assembled capture, gate,
  corpus and native qualification remain pending. Only audit 22 is deferred.
  Nothing is pushed.

- Coordinator continuation (2026-10-24, remaining row readiness):
  Record 0362ul11 checks all 156 historical result-contract rows against clean
  staging `778d7a5e6c`: 83 exact registered semantic identities, nine removed
  aliases, 42 repaired non-DrawingML rows and 22 active DrawingML rows. No
  surviving non-DrawingML gap is found among those historical rows. This is
  source readiness, not a replacement for fresh all-68 capture/dispositions.

  Fragment source review passes; its test receipt is corrected to the actual
  diagnostic 6d lock and unlocked offline commands. Pinned composed validation
  remains required. Diagram is returned for remaining selected-budget and
  nested-graph fidelity defects. Text/media's six standalone SVG consumers
  are queued for actual locked checks. The native DrawingML protocol is
  returned for simulated owner operations and preflight evidence; real API
  adapters and distinct Ink baselines remain mandatory. The repaired OLE
  qualification packet is under renewed independent review.

  Cleanup forty-two removes the merged legacy protocol leaf non-forced,
  retaining its branch and all nine payloads (356,146 bytes), with successful
  restoration and independent hash/mode/ancestry verification. Existing
  native endpoints, targets and the protected trunk probe remain intact.
  Final capture, full gate, real corpus and native qualification still await
  the assembled source. Only audit 22 is deferred. Nothing is pushed.

- Coordinator continuation (2026-10-24, OLE protocol and PPT preparation):
  Reviewed OLE owner qualification packet `89142c8518` merges at
  `15263c83d2`. It binds the true pre-owner baseline, eleven routes, complete
  output oracles, retained histories, actual toolchain selection and ordered
  artifact evidence. All 89 combined tooling tests pass; record 0362ul10f
  accepts the protocol implementation. Actual paired builds, correctness and
  performance qualification remain required.

  PPT fixture-preparation packet `95257e5103` merges at `7fc786eedc` after
  independent validation of the real sparse-checkout CLI receipt and all 94
  composed tooling tests. Five additional compile fixtures were already in
  the historical tree but absent from the sparse worktree. Record 0362ugjac
  preserves the failed attempts and distinguishes preparation acceptance from
  native qualification. The baseline release build subsequently passes;
  correctness reporting exposes an optional-endpoint null/false bug, and its
  consumer repair still needs paired rebuilt validation.

  The actual current RTF crate-example check fails at the byte/string parse
  and Commit accessor boundaries. Consumer-only baseline/current adapters
  subsequently compile and pass eight correctness routes each, but review
  returns the packet because the original live crate example remains broken.
  Both consumers must migrate before acceptance; the historical failed
  receipt stays intact.

  Chart, diagram, fragment and text/media source reviews pass. All six SVG
  standalone harnesses also pass locked checks, with registry identities
  unchanged. Theme/Ink remains under a final return for replacement font
  counts and mixed modeled namespaces. The true common DrawingML owner
  baseline is `945eb1d6b9`; `2d330c3305` is already post-owner and cannot
  qualify that migration. Raw lanes retain their separate baselines. Final
  source capture, the full gate, real corpus and all native qualification
  remain required. Only redaction audit 22 is deferred. Nothing is pushed.

- Coordinator continuation (2026-10-24, DrawingML composition and six-leaf cleanup):
  Shared DrawingML owners and all five raw codec lanes merge at `04058bc7e1`.
  Composed 4100 locked consumer checks and 501 Rust tests pass; record 0362dvj
  preserves stale-cache and missing-argument failures before the repairs.
  The migration retains input budgets and opaque namespace fidelity, and all
  affected live consumers follow the required API.

  PPT probe correctness merges at `571f31a8bd`: both normal release endpoints
  verify 28 routes across four fixtures with identical output reports. RTF
  native adapters and the live example merge at `2a50681aec` after 42 tooling
  tests; OLE oracle and per-endpoint lock repair merges at `07cd359eb4` after
  63. These are source/protocol and correctness receipts, not final paired
  performance acceptance.

  DrawingML owner adapters, raw fixture/fidelity checks and independent child
  memory observation merge through `cef9dd63cb`, with all 80 composed tooling
  tests passing. Actual current compilation exposes fixture-generator and
  SVG oracle mistakes; repair `5c67e58938` compiles and passes ten untimed
  fixture preflights. Historical 945/11728 compile, while 2f7/996 expose
  raw-family feature gates corrected at `418cff01db`. The broad performance
  harness still needs current defaults, isolation from unrelated historical
  lanes and actual paired build validation. Record 0362dvid distinguishes
  those obligations from the successful minimal-driver diagnostics.

  Cleanup waves 43–45 remove six merged leaves non-forced, retaining every
  branch and all eighteen ignored payloads (1,255,989 bytes). Restore exercises
  and independent hash/mode/ancestry checks pass. OLE cleanup waits for foreign
  read-only holders; external targets, native endpoints and the protected trunk
  probe remain intact. Record 0362ul12 captures the integration and cleanup.

  Fresh all-68 capture/dispositions, the full 53-step gate, actual corpus
  replay, glossary closure and every required native qualification remain
  active work. The prepared gate now registers 25 tooling modules without
  changing its 53 steps. Only redaction audit 22 is deferred. Nothing is pushed.

- Coordinator continuation (2026-10-24, native correctness and DrawingML return):
  Fresh capture at `29f072f19c` stops after nine packages on an ambiguous
  DrawingML rustdoc link. Repair `642f2e45f6` selects the error enum explicitly;
  all-feature hidden JSON rustdoc passes with the pinned 4100 lock. The failed
  capture is retained unchanged, and a new all-68 capture is running from
  frozen `082fbeafe8`. Record 0362dvk preserves both failure and focused pass.

  Both OLE normal release endpoints now execute the complete correctness
  oracle successfully: eighteen artifacts each, all byte-identical. Independent
  source, command and artifact verification passes 146 checks. Record 0362ul10j
  accepts that correctness evidence; allocator and paired timing qualification
  remain required.

  DrawingML native qualification is returned for timed setup/copies, incomplete
  source-backed mutation checks and weak semantic oracles. Record 0362dvie binds
  a reviewed plan that retains every row, sample requirement and baseline fence.
  Separate owner and raw workers are repairing it. Harness isolation merges at
  `47ea67dc44` after source review, and all forty composed tooling tests pass;
  actual paired builds and native timing/fidelity acceptance remain pending.

  Cleanup wave 46 removes the merged DrawingML observer and PPT protocol leaves,
  preserving both branches and seventeen ignored files (918,418 bytes). Restore
  and independent hash/mode/ancestry checks pass; record 0362ul13 contains the
  receipts. The protected trunk probe and all native endpoints remain intact.
  Full gate, actual corpus replay, final inventory/glossary closure and native
  qualification remain active. Only redaction audit 22 is deferred. Nothing is
  pushed.

- Coordinator continuation (2026-10-24, OLE allocator metadata):
  Both locked allocator builds and native oracle executions pass. Actual output
  exposes two protocol-fixture mistakes: replacement size 19 instead of 18 and
  a hyphenated stream name instead of the native underscore. Record 0362ul10k
  repairs both consumers and the fixture, adds stream-size agreement checks,
  and retains the original rejected diagnostic. Both normal and allocator
  output pairs pass the repaired validators; 66 tooling tests and independent
  artifact/conservation/history review pass. Peak allocation-count increases
  remain visible in the record for final qualification review. No native
  performance acceptance is claimed; no branch is pushed.

- Coordinator continuation (2026-10-24, complete capture and wire canonicalization):
  Fresh all-68 capture at `082fbeafe8` passes with no reused artifacts. All 84
  inventory rows receive independent source-role review and exact dispositions;
  no DrawingML or historical result/limits-error rows remain. Record 0362ul14
  binds every capture and review artifact. Review also finds six unused IWA
  wire aliases outside the inventory. Repair `69194102b6` removes them, adds
  exact glossary bans, and registers the canonical wire-request descriptor.
  All 68 tooling and 64 codec tests, consumer checks, focused rustdoc and the
  glossary pass. The full capture remains bound to 082; the alias repair has
  a separate fresh IWA-protos capture.

  DrawingML raw packet `b2da8f8ba6` is returned before merge. Record 0362dvifa
  requires real inherited SVG context, independent mutation fidelity checks,
  removal of timed error normalization/setup, and matching Diagram cap routes.
  It also corrects the false claim that the Ink writer canonicalizes source:
  both historical and current writers retain exact bytes, and qualification
  must keep that proof. Owner and raw repairs continue in separate leaves.

  The current 691 source is assigned a fresh full 53-step gate and actual
  corpus verification. Final source binding, native qualification and glossary
  performance closure remain required. No new deferral or performance waiver
  is accepted; only redaction audit 22 remains deferred. Nothing is pushed.

- Coordinator continuation (2026-10-24, gate restart and two-leaf cleanup):
  The current 691 full gate stops at formatting in DrawingML's diagram model.
  Repair `f1d166b9c4` changes layout only, and the full workspace format check
  passes. Record 0362dvl retains the failed gate and focused repair. A fresh
  full 53-step gate starts from frozen `b4c18f2ede`, with a separately rebound
  OfficeCLI postchecker and the unchanged pinned locks. Isolated documentation
  steps still require a window without other builders.

  RTF qualification pins were stale after the harness feature migration.
  Reviewed repair `d9d7abab4e` merges at `b4c18f2ede`; independent review and
  all 44 build-free protocol tests pass. Record 0362ul8d preserves the exact
  baseline, routes and fidelity requirements. Fresh RTF and OLE packets bind
  b4 and await actual build and quiet measurement slots; neither preparation
  is performance acceptance.

  Cleanup wave 47 removes two more merged leaves, ODG native protocol and the
  legacy safe runner. Both branches and all fourteen ignored files (706,737
  bytes) remain available; restore exercises and independent hash, mode,
  ancestry and worktree-absence checks pass. Record 0362ul15 binds the receipts.
  Native endpoints and the protected trunk probe remain unchanged. DrawingML
  qualification returns, full gate/corpus completion, native measurements and
  glossary performance closure remain active. Only audit 22 is deferred.
  Nothing is pushed.

- Coordinator gate update (2026-10-24, CFB error layout):
  The b4 gate passes formatting, then stops at all-feature Clippy with 45
  `result_large_err` diagnostics for CFB's `SequentialWriteError` candidate
  validation variant. The failed source, locks and logs remain unchanged.
  A separate leaf is repairing the error layout while preserving its typed
  cause, complete publication report and progress. Slot 2 is reassigned from
  the stopped gate to that repair. A new source binding and full gate remain
  required; prepared b4 native packets are source-only diagnostics.

- Coordinator continuation (2026-10-24, composed proofs and consumer closure):
  CFB error layout and canonical output accessors merge with the final ordinary
  harness consumer repair at `c796d0acc2`. The error keeps its typed source and
  complete report; the success path gains no allocation. All eight missed
  benchmark consumers now use `output_bytes()`. Independent source, lock, log,
  target and adapter-pin review passes. Records 0362ul16, 0362ul16a and 0362ul16b
  preserve the initial failures and successful 555-test CFB suite, 23 tooling
  tests and ordinary performance harness all-target check.

  DrawingML owner and raw proof repairs compose at `a3ec979993`. Native owner
  helpers are measured directly, with mutation, remainder and stale-patch
  proofs outside timing. All six raw diagnostic lanes pass their boundary,
  malformed, output-cap and fidelity checks; exact Ink source preservation is
  retained. Independent reviews pass, and all 41 combined tooling tests pass.
  Records 0362dvig, 0362dvifb and 0362dvih distinguish those results from paired
  performance qualification, which remains required.

  The Clippy return also exposes a quadratic DrawingML attribute iterator.
  Repair `eb94d5a6b2` uses the existing bounded checked-attribute helper and adds
  a late-attribute/duplicate regression test. All 486 DrawingML tests and the
  assembled facade, ops, CLI and Python all-feature/all-target Clippy check
  pass. Record 0362dvla retains each intermediate failure. A fresh full 53-step
  gate and actual corpus replay start on immutable c796 source. Fresh all-68
  API capture continues on a3; its production API equivalence to c796 is
  verified separately before final inventory acceptance.

  Combined native correctness runs and the remaining qualification builds are
  active. No performance regression is waived, and the prepared packets alone
  do not establish acceptance. Only redaction audit 22 is deferred. Nothing is
  pushed.

- Coordinator cleanup update (2026-10-24, two merged gate leaves):
  Cleanup wave 48 removes the merged zip gate repair and Numbers Snappy leaves
  without force. Both branches and their two ignored lockfiles (279,058 bytes)
  remain preserved. Restore exercises and independent artifact, mode, ancestry
  and absence checks pass. Record 0362ul17 binds the receipts. The removed
  checkouts contained 12.51 GB of regular files; native endpoints, targets and
  the protected trunk probe remain untouched. The staging hold is released.

- Coordinator capture update (2026-10-24, fresh API evidence and glossary return):
  The a3 all-68 capture completes with no reused artifacts. All 84 inventory
  rows retain valid reviewed dispositions after independent source-role
  revalidation, including three IWA wire-request rows affected by alias
  removal. The coordinator verifies every artifact, manifest and log hash;
  production API equivalence binds this capture to c796. Record 0362ul18
  preserves the fresh evidence and distinguishes the blank scaffold from
  reviewed decisions.

  Supplemental glossary checks pass 55 of 59 profiles. Two DOCX failures
  identify the same public ArtifactStreamPublication exposure; return 0362ul19
  renames that raw receipt and migrates every consumer without losing streaming
  capability. Six stale soapberry policy scopes also need their actual office
  namespace. The Keynote-minimal finding is an all-feature mismatch, with no
  source or policy change required. The complete feature-specific gate remains
  active on a separate c796 checkout; formatting passes. Native builds and
  correctness checks continue, with performance acceptance still pending.

  Final native source preparations for durable Ink and W4b/W5 are reviewed and
  merged in records 0362ld and 0362ul16c. Their seven and forty-one tooling
  tests pass respectively; they make no native-performance claim. OLE owner,
  RTF, sheet and property preparations are also ready for their assigned
  build slots. The iWork allocator protocol and DOC body-toggle fidelity
  preparation have separate pending reviews. No previous binary or result is
  relabeled as current evidence.

- Coordinator native-protocol update (2026-10-24):
  The iWork allocation protocol now measures all fifteen constructor and four
  owner routes using separate instrumented binaries, operation-scoped counters,
  typed metric validation and guarded fixed-work ABBA collection. Independent
  review and 23 Python tests pass; the protocol merges at `f729439957` with
  native Rust compilation and execution still required. Record 0361fdc binds
  the source and exact commands.

  DOC body-toggle preparation merges at `27c7ffe0b5` after independent review.
  Record 0359qf binds the exact two-file baseline reversal, current probe and
  CHPX fidelity checks. Its candidate and baseline test commands are corrected
  to separate filters and the actual leg-specific names. Output dumping is
  opt-in and outside the measured interval. No native run is claimed.

  The ordinary composed c796 DrawingML release build now passes, followed by
  all six raw and four owner native correctness routes. The receipt is
  independently under review; generated-lock historical diagnostics remain
  excluded from final baseline qualification. The released build slot goes to
  actual OPC/DOCX report-rename compilation and glossary checks. Full gate,
  actual corpus replay and paired performance acceptance remain active.

- Coordinator correctness handoff (2026-10-24):
  Independent review accepts the exact c796 DrawingML release build and all
  ten current native correctness reports. All artifact identities and native
  semantic checks pass, as do thirteen protocol tests. Record 0362dvii merges
  at `5affb487cf`; its four true historical baseline preparations validate all
  ten profiles with the accepted 52ed/4100/eac862 locks. Historical builds and
  paired performance runs remain unexecuted, and generated-lock diagnostics
  are excluded.

  The OLE paired runner refuses its build phase before Cargo because unrelated
  builders are active. The failed receipt is retained; fresh retry roots are
  ready for a globally exclusive build window. An available slot alone does
  not satisfy that runner's guard. The OPC/DOCX rename is undergoing actual
  validation, the full gate continues, and PPT builds remain active. Nothing
  is pushed.

- Coordinator glossary repair (2026-10-24):
  OPC ArtifactStreamReport and every DOCX consumer merge at `472b670547`, with
  no alias, layout or streaming-behavior change. The exact old-name bans and
  six corrected soapberry office scopes remain enforced. Independent review
  verifies the 4100 lock, all source and log hashes, nine OPC stream tests,
  one DOCX consumer test, all-target/all-feature compilation, and fresh OPC,
  DOCX and soapberry glossary checks. Record 0362ul19 binds the actual passes.

  A new all-68 capture is assigned to immutable 472b source. The c796 full gate
  has passed formatting and all twenty-four Clippy profiles; its remaining
  tests, isolated documentation and actual corpus verification continue.
  Native preparations must bind the final source before performance acceptance.

  Independent property-packet review finds an executable manifest-identity
  mismatch: preparation accepts an external final manifest while the legacy
  build validator requires its canonical in-tree path. Return 0362ul8e is
  correcting the shared preparer/validator contract, including RTF, without
  weakening source or workload checks. Source-only preparation is not promoted
  to native acceptance. Only redaction audit 22 is deferred. Nothing is pushed.

- Coordinator executable-protocol and corpus update (2026-10-24):
  Legacy manifest identity repair 0362ul8e merges at `2b01fc91d7` after
  independent review and 46 tooling tests. RTF, OLE-properties and
  XLS-properties now bind the same exact external manifest through
  preparation, build and validation. The fresh final-source packet 0362ul8f
  independently validates all three manifests and six source legs; no native
  qualification is inferred from preparation.

  The c796 gate passes formatting, all 24 Clippy profiles and the full facade
  selection. Its corpus failure is a stale golden caused by the newly tracked
  PPT fixture. Every previous report row is unchanged. Record 0362ul20 adds
  that fixture's byte-exact no-op row and correct summary counts; the fresh
  diagnostic and non-blessing comparison both pass. The correction merges at
  `93bdf16bf4`, with the failed original gate receipt preserved.

  W8 original-f2 provenance closes its 459,676-file source fence. Review then
  returns three executable-script defects: missing release/target flags,
  missing build receipts, and an unbound candidate checkout. The repaired
  script requires the exact clean 472b source, validates pinned inputs and
  archive, creates build receipts and rejects stale targets. Independent
  review and 45 tooling tests pass; the repair merges at `28c2ffcd90`.

- Coordinator final-source evidence update (2026-10-24):
  Record 0362ul21 binds 68 fresh API captures to 472b, with no reused artifact.
  Independent review verifies every receipt and all 84 reviewed inventory
  roles. The blank disposition scaffold remains explicitly unaccepted.
  Supplemental glossary checks pass 58 of 59 profiles; only the known
  Keynote-minimal/all-feature-JSON mismatch remains. Actual feature-specific
  checks remain required. The capture record merges at `9781cf6d87`.

  That assembled revision is frozen for the next full 53-step gate. Its Rust
  production source is identical to 472b, and it includes the report and
  qualification-tool repairs. The earlier c796 diagnostic endpoint stays
  immutable. Isolated documentation and actual corpus verification remain
  required before final acceptance.

  PPT's c796 normal and allocator binaries pass all 28 correctness routes
  across four fixtures under independent review. Final 472b builds are now
  running. The final iWork constructor's four endpoint variants pass all
  fifteen correctness routes with matching reports; owner builds continue.
  Neither result is a timing or allocation-performance acceptance. Reviewed
  DOC, Ink, W4b/W5, legacy and sheet preparation packets bind the final source;
  actual remaining builds and quiet-host measurements stay required.

  Cleanup wave 49 removes the merged DOCX validation leaf without force,
  preserving its branch, six ignored files and validation evidence. Restore
  exercises and independent hash/mode/ancestry checks pass (0362ul22). The
  protected trunk probe remains intact. Nothing is pushed; only audit 22 is
  deferred.

- Coordinator correctness and cleanup update (2026-10-24):
  The final 472b iWork constructor and owner preflights complete. All four
  constructor variants pass fifteen routes, and all four owner variants pass
  four routes, with byte-identical reports. Independent review verifies 95
  artifact identities and 28 command receipts. Record 0361fde preserves the
  actual owner baseline input and its canonical Git resolution; the corrected
  record merges at `16d8e593b4`. Collection has no timing or allocation packet,
  so native performance acceptance remains required.

  Glossary reconciliation 0362ul23 merges at `74c032cc09` after independent
  review. It corrects W1's source-backed commit shape, binds fresh API evidence,
  closes stale ZIP/OPC source-debt wording and accurately records the earlier
  PPT first-commit improvement. Final native qualification and the workspace
  gate remain open. The signed-native audit identifies missing timed signed
  routes in the otherwise unsigned W8 matrix; a dedicated qualification
  protocol is in progress without changing the frozen production source.

  The fresh 9781 gate has passed its first sixteen checks at this checkpoint.
  Final PPT and W8 builds continue. Both DOC body-toggle release builds finish,
  followed by targeted fidelity tests; their actual test selection is being
  checked before acceptance. Isolated documentation, remaining builds and
  quiet-host performance measurements stay required.

  Cleanup wave 50 removes three merged qualification/report leaves without
  force. Branches and five ignored files remain preserved, restoration checks
  pass, and independent verification confirms retained source/evidence paths
  and the protected trunk probe (0362ul25). No performance exception is added
  and nothing is pushed.

- Coordinator native preflight update (2026-10-24):
  Final DOC body-toggle build and correctness evidence merges at `c204999a44`
  after independent review (0359qg). Both release builds pass. Three corrected
  module-qualified tests each execute once; the earlier zero-test invocations
  remain diagnostic. Single-edit outputs are byte-identical, and the CHPX
  oracle verifies the intentionally smaller batch output. Allocation and
  quiet-host timing remain open.

  Final PPT normal and allocator builds pass, followed by all 28 correctness
  routes across four fixtures. Independent review verifies the reused
  baseline and matching normalized proofs. A truncated probe hash is corrected
  from the actual source; the original receipt remains preserved. The corrected
  record merges at `479234623d`, with no timing or ABBA claim.

  The signed-native protocol is returned for concrete adapter, provenance,
  process-guard, comparison and memory-evidence repairs. Two workers are
  implementing the executable build adapter and qualification protocol; neither
  is accepted yet. A separate compile-only scheduling change is under review
  to permit bounded concurrent builds while preserving exclusive measured,
  documentation and profiler phases. Frozen packets still use their existing
  guards.

  The final gate has passed formatting and all 24 Clippy profiles and is
  running the ops tests. W8's original baseline build finishes and its final
  candidate builds; sheet builds are active. PPT's released slot is assigned
  to OPC build/correctness only. The original OPC script would continue into
  ABBA, so only its reviewed prefix is authorized; its timing guard is returned
  for repair before any measured run.

  Cleanup wave 51 removes the merged DOC and iWork author leaves while
  preserving branches, five ignored files and all native evidence (0362ul27).
  A held glossary leaf is retained. Three obsolete searches are stopped by
  their owners without changing files. Only audit 22 is deferred; nothing is
  pushed.

- Coordinator W8 correctness and gate-resource update (2026-10-24):
  Both W8 release endpoints now pass all 109 correctness routes. Independent
  review verifies the original f2 and final 472b archives, role-specific locks,
  consumer adapter, 518 shared registry identities and six reviewed local
  graph changes. The endpoint receipt-path repair merges at `553dab0517`;
  record 0362ul31 merges at `bf603625fa`. It preserves the initial endpoint
  assembly failure and records the later assembly from retained validation
  outputs. The pair is `verified_protocol_only`; timing, allocation and final
  performance acceptance remain open.

  The reviewed compile-only lease protocol merges at `b5964465a2`. It admits
  only explicitly permitted generic-runner build commands, with exact manifest,
  process, source and target identities. Ordinary builders and the full gate
  remain outside those permits; measurements and controls require a quiet
  host. No existing preparation packet is silently rebound. The composed
  integration passes 114 focused tooling tests, and its production Rust,
  Cargo manifests and toolchain remain identical to 472b. A subsequent targeted
  lifecycle review finds that escaped descendants can outlive the generic
  runner's leader; that repair must pass review before any lease cohort runs.

  The 9781 gate passes checks 0 through 24, then stops during ops examples
  after its target grows to 20,763,352,771 bytes. Completed test chunks pass;
  the row remains failed and its target, logs and receipts are preserved.
  The first bounded-retention repair is returned for failed-artifact retention,
  escaped-descendant handling, live guards on every gate command, exact private
  artifact selection and command-count reconciliation. No replacement gate
  run is accepted yet. Isolated documentation and actual corpus replay remain
  required.

  ODG and ODF source preparations pass independent review, including both ODG
  source closures and a fresh 18-test ODF suite. ODP's plan is returned because
  its expected binary has no matching build handoff and its source, lock and
  probe bindings are incomplete. Signed-native qualification also remains
  under repair for process lifetime and receipt provenance. These are active
  engineering returns, not additional deferrals.

  Both XLSX release endpoints and the XLS F3 baseline build successfully.
  The F3 candidate fails on seven outdated consumer-overlay calls and a
  tracked fixture missing from its sparse checkout. Its preparation is
  returned for repair, with the failed log and target preserved. No sheet
  timing or allocation result is accepted from these builds.

  Cleanup wave 52 removes five merged leaves with their branch references
  retained (0362ul28). Independent verification checks all 13,912 preserved
  payloads, 25 protected identities and the current sheet command dependencies.
  A helper-path ambiguity is resolved against both exact engine hashes without
  altering the raw receipts. The protected trunk probe remains intact; only
  audit 22 is deferred and nothing is pushed.

- Coordinator guarded-execution update (2026-10-24):
  The gate retention and descendant-lifetime repair passes independent review
  and merges at `a658fee291`. Its first actual retry exposes a target-sampling
  race: Cargo removes a temporary file while `du` scans it, terminating the
  guard thread. That run is stopped, its processes are verified gone, and its
  source, target and failed evidence remain preserved without acceptance.

  The follow-up retries transient sampling races and makes unexpected guard
  failures stop and reap owned processes. Independent review verifies all
  seventeen tooling tests, escaped-writer and hung-child failures, retained
  artifacts, unchanged 53-step coverage and the exact 122-command ops plan.
  It merges at `5e1881da1a`. A fresh outer-runner invocation is executing only
  checks 25 through 33; both live guards are healthy at the checkpoint.
  Isolated documentation and the rest of the gate remain required.

  Signed-native protocol 0362ul24 merges at `75e5dd4b0f` after independent
  review and 26 tests. It binds real route receipts, source and fixture
  identities, bounded process lifetime and retained memory evidence. Its
  ownership contract explicitly requires one serial controller and protects
  pre-existing foreign children. The build adapter, actual qualification and
  final collection remain pending; no signed performance result is claimed.

  Further F3 review rejects an overlay that restores obsolete limits calls
  in other active benchmark consumers. The next candidate restores those
  files to exact 472b source; full paired preparation still requires baseline
  rebinding and review before a build grant. The original failure is seven
  arity errors plus one missing-fixture error, rather than eight arity errors.
  Generic build-leasing cleanup passes independent review and merges at
  `53c8056ece`. It checks process identities before admitting descendants,
  captures children through final cleanup and fails closed on monitoring
  errors. Restricted process metadata is allowed only when readable identity
  fields prove the process is outside the protected scope. The composed
  integration passes 76 native-runner/paired tests and 26 signed-protocol
  tests. Fresh build permits and manifests are still required before any
  concurrent qualification build.

  The gate retry's first ops library chunk passes all 1,012 tests and advances
  to integration tests; the full ops row and remaining selections are still
  running. The W8 ABBA execution wrapper remains under review. These returns
  preserve the every-consumer and performance requirements; only audit 22 is
  deferred, and nothing is pushed.

- Coordinator paired-build and execution-preparation update (2026-10-24):
  The F3 driver now binds the original baseline harness and clean 472b
  candidate separately (`f5961b28db`). Independent review verifies unchanged
  workload semantics, all active consumers and the reusable baseline binary.
  The fresh candidate release build passes; its actual binary and later
  qualification command bind the new target path. Record integration
  `e064a22fc4` preserves the earlier failed build and accurately records a
  target-monitor observation gap. No F3 ABBA or memory qualification is
  inferred from the successful build.

  OPC's two release builds and four correctness reports pass independent
  review. Record 0362umfa is accepted at `e9aa4cada9`, separately from the
  returned timing guard. The original missing-fixture failure and corrected
  sparse-checkout materialization remain documented. Timing-guard repairs
  still require review before any measured run.

  ODG/ODF source preparations and the corrected ODP build/memory handoff
  merge at `53a99025ac` after independent review. The final ODP packet checks
  source, locks, binaries and fixtures and rejects incomplete memory metrics.
  Its five negative cases pass; actual builds, measurements and the historical
  memory residual remain open.

  W8's guarded execution wrapper merges at `4dbf31e4b1`. Independent review
  verifies 91 focused tests and actual harmless systemd-scope lifecycle
  checks, including the host's post-exit metadata representation. The frozen
  probe and policy remain unchanged. Its real ABBA run still needs a quiet
  window. The five-lane compile-only cohort is being corrected to bind fresh
  temporary directories consistently; it has not run. The 9781 gate continues
  its ops integration chunks with bounded target retention and healthy guards.
  All remaining gate, corpus and performance requirements stay active.

- Coordinator execution and cleanup update (2026-10-24):
  The corrected five-lane compile packet passes independent review and merges
  at `589d5d0adc`. All lanes bind distinct temporary directories; fourteen
  build phases, twenty-seven referenced artifacts and the unchanged workloads
  validate. Seventy-eight tooling tests pass. Execution still waits for the
  active, unleased gate; preparation is not a build result.

  The composed F3/W8 supplement passes 71 tests. Its 23 quiet-wrapper tests
  are also rerun directly; those are overlapping coverage, not additional
  unique cases. A source-equivalence audit reconciles the later tooling with
  the frozen 9781 gate. The glossary definitions and production sources are
  unchanged, but the actual remaining glossary/gate selections still need to
  run. The separate signed and gate-helper tests retain their accepted,
  source-identical receipts.

  OPC timing guard `2732b45b61` passes independent lifecycle review and five
  tests on the composed tree. A combined monitor-failure/escaped-child leak
  found in its predecessor is repaired; owned children are reaped and
  pre-existing foreign children remain untouched. The accepted build record
  retains its earlier return history. Final command-packet review and actual
  quiet-host timing remain separate requirements.

  Cleanup wave 53 removes one unused, merged DOC author worktree without
  force (0362ul33). Its branch and ignored lockfile are retained; a sparse
  restore exercise and seventeen independent postchecks pass. Measurement
  sources, binaries, failed evidence and the protected trunk probe remain
  intact.

  The DOC and Ink handoff reviews identify execution controls and receipt
  bindings that still need repair; their accepted correctness/build evidence
  is unchanged. The sheet and W8 final command packets are being reviewed.
  The gate continues its ops chunks with bounded retention and healthy
  guards. Isolated documentation, corpus replay and native performance
  acceptance remain open; only audit 22 is deferred and nothing is pushed.

- Coordinator gate return (2026-10-24):
  The 9781 ordinary-gate retry stops at ops chunk 36 after thirty-six completed
  chunks pass. `pptx_batch_equivalence` finds a real signed-package mismatch:
  the batch returns `DocumentProtected` at item zero, while applying that
  item's missing-notes path individually returns `PathNotFound`. The failure
  is preserved and assigned for a correctness repair; equivalence must not be
  weakened to accept it.

  A separate outer-runner cleanup error incorrectly identifies the runner
  itself as a remaining owned child. That failure is also retained and sent
  back for repair with a regression. The endpoint has no remaining owned
  processes; its target and quarantine remain intact, and no retry is started.
  A fresh isolated documentation selection is prepared pending the runner
  assessment. Candidate performance execution remains held while the affected
  source and required receipt rebinding are determined.

- Coordinator admission-contract and documentation update (2026-10-24):
  Further review traces the signed PPTX mismatch to the explicit batch
  preflight, rather than a coalescer failure. The existing DOCX admission
  test requires refusal before the iterator is polled; the older PPTX
  equivalence test requires the same path error as an individual command.
  Those requirements cannot both hold for an unread item. The owner has been
  asked which contract should govern. No production change or proposed test
  narrowing is accepted while that decision is pending; signed exact no-ops,
  cancellation and source-freshness checks remain protected.

  The runner-ownership fix merges at `21fcafa7ef`. It excludes the controller
  and its pre-existing ancestry by PID and start time. Independent review and
  the composed tree both pass nineteen tests, including a real failing child
  and escaped-child cleanup. The isolated `wsdoc` selection is now running
  against unchanged 9781 source with that repaired runner. Other compiler and
  measurement jobs remain held; the separate core/ops/facade documentation
  selection has not started.

  Original-f2 signed-consumer preparation completes its source-only closure
  verification. Adapter composition is under review; candidate preparation
  and actual builds remain held. DOC, Ink and PPT execution-adapter repairs
  continue against concrete review findings. None of these preparations is
  performance acceptance, and nothing is pushed.

- Coordinator reviewed execution-tooling update (2026-10-24):
  The Ink boundary and concrete final-source handoff merge at `6e323fcbc5`
  from reviewed leaf `0dbc09268b`. The actual systemd 259 checks verify the
  named scope, runtime limit and cleanup of escaped children; the composed
  tree passes all eight boundary tests. The accepted paired release builds
  and correctness artifacts remain unchanged. Allocation and native timing
  are still required.

  The signed-consumer source-preparation adapter merges at `1c2ce1f7d5`
  from `efad4fabbc`. Independent review verifies its composition with the
  accepted signed protocol, the original-f2 source closure, role-specific
  locks and fixed dependency policy. All seven adapter tests pass on the
  composed tree. Original-f2 preparation has completed; it has no build or
  binary receipt. Candidate preparation remains held for the signed-batch
  contract decision.

  Isolated workspace documentation continues on unchanged 9781 source.
  Other build and measurement jobs remain held. The separate core/ops/facade
  documentation row may run alone after this row passes and its processes
  terminate. The earlier interrupted documentation invocation is retained
  as an aborted partial run: Cargo and rustc did start, so it is not a
  setup-only attempt. Neither attempt is counted as passed before its
  terminal receipt. Cleanup review continues for inactive merged leaves;
  no additional leaf has yet been removed. Nothing is pushed.

- Coordinator cleanup wave 54 (2026-10-24):
  Removed the clean, merged W4 integration checkout without force and kept
  its branch at `f3fed35c0d`. All thirteen ignored files (6,960,154 bytes)
  are quarantined with exact identities; sparse restoration passed. Root
  review passes 44 applicable postchecks, including the protected trunk
  probe and retained active source roots. Record 0362ul34 preserves the
  corrected preaudit count and directory-mode audit distinction. The
  checkout-allocation reduction is approximately 7.54 GB after quarantine;
  no isolated filesystem free-space delta is claimed. No target, native
  evidence or foreign process was removed. The integration hold is released.

  Supplemental tooling reconciliation distinguishes the historical
  nine-module failure from the still-pending 9781 tooling row. A separate
  leaf is registering the newly accepted suites for the current source;
  the runner used by active documentation remains unchanged. The owner
  signed-batch contract choice and actual native qualification remain open.

- Coordinator DOC execution-protocol review (2026-10-24):
  Reviewed leaf `29c7879b3f` merges at `d91dd56c08`. Its final schedule
  records 31 actual samples per invocation and two ABBA blocks. The 250 ms
  aggregate floor now uses the certified lower bound
  `16 * median_us + 15 * min_us` for each 31-sample invocation, with finite,
  nonnegative and ordered timing validation. It does not claim that a
  median multiplied by the sample count is measured accumulated work.
  Independent fake-command review and the composed protocol checks pass,
  including short-work, source-identity and missing-profile failures. The
  accepted release binaries remain unchanged; native timing and profiles
  have not run.

  The signed original-f2 handoff now records nine exact prepared files and
  the retained archive path. Root independently verifies those files and
  the pinned Cargo shim; the corrected build preflight succeeds on the
  integration ancestry without invoking Cargo. Candidate preparation and
  all dependent policy changes still await the owner's signed-batch
  contract choice. Isolated workspace documentation remains active; the
  core/ops/facade documentation row and actual CLI corpus are authorized
  only sequentially after their preceding row passes and all its owned
  processes terminate. No full gate or performance acceptance is claimed.

- Coordinator isolated documentation result (2026-10-24):
  Workspace `wsdoc` row 34 passes on clean frozen 9781, including generated
  `litchi-iwa-protos` documentation. The isolated, offline/locked,
  all-features run with warnings denied completes in 2,138.861 seconds with
  no diagnostics. Target storage remains below 3.39 GB; all owned processes
  terminate and all six locks remain unchanged. Record 0362ul36 binds the
  actual command, tool identities and independently verified artifacts.
  This closes the isolated workspace-documentation run, not the full gate.
  Row 36 is now running alone; actual CLI corpus follows only if it passes.

  Tooling registration merges at `3431ba566c` from reviewed `80caba940d`:
  53 gate steps now include 31 tooling modules. The six added suites pass
  together (68 tests); the OPC test's file-created-before-PID-written race
  is fixed with a bounded handshake that preserves ownership and cleanup
  assertions. Record 0362ul35 retains the failed run and corrected receipts.
  Frozen 9781 tooling row 48 remains unrun; active documentation keeps its
  original f5 runner. The signed-batch contract choice and final native
  qualification remain open, and nothing is pushed.

- Coordinator documentation and consumer-sweep update (2026-10-24):
  The separate core/ops/facade documentation row 36 also passes on unchanged
  9781: 612.126 seconds, no diagnostics, target below 1.92 GB, unchanged
  locks and no remaining owned processes. Record 0362ul36 now binds both
  isolated documentation results. CLI execution is briefly held for the
  targeted validation of a newly discovered consumer repair.

  The final source audit finds one compile-stale RTF fuzz consumer calling
  removed decompression names, plus stale CLI/fuzz/DOC/PPT/XLSB examples.
  Record 0362ul37 is assigned to migrate those consumers and register the
  missing fuzz check; no alias will be restored. This is required engineering
  work, not a new deferral. The audit also confirms that historical 0362j
  changed-commit/replay measurements cannot be attributed to current 472b
  because measured CFB/PPT paths have changed. A final run of that same
  workload is being prepared separately from the diagnostic primary-noop
  cold probe. Actual qualification and the signed-batch decision remain open.

- Coordinator consumer and PPT tooling integration (2026-10-24):
  RTF/CLI/DOC/XLSB consumer and guide repairs merge at `1ff6aae60a` from
  `2cb3cdd55d`. The actual locked RTF fuzz check passes in 72.066 seconds,
  with a maximum sampled target of 118,970,506 bytes. Root independently
  verifies the receipt and all 83 registry identities against the approved
  root lock's 518 identities. The current fuzz gate also covers RTF and
  observes its standalone lock. Record 0362ul37 retains the initial sparse
  metadata failure and distinguishes the later successful check.

  Reviewed PPT tooling merges at `30aaafbf98` from `8d8dfce1f`. Public
  execution validates its build projection and child-event provenance;
  26 composed tests pass with the tools import path configured. The initial
  plain module invocation exposes an import failure, so 0362ul41 is assigned
  to fix the current tooling row and register both new suites. The frozen
  28-route workload is unchanged. Cold primary-noop observations remain
  supplemental; neither tooling readiness nor old 472b builds are final
  performance acceptance.

  A further source review finds public PPT record/codec decoders that still
  choose limits internally. Required engineering record 0362ul39 migrates
  their canonical entry points to explicit limits and updates current
  consumers, preserving package-selected profiles and public capabilities.
  An independent sweep is checking other formats for the same gap. Final
  candidate API and performance bindings must be refreshed after this repair;
  the CLI run remains held during source stabilization. The signed-batch
  contract choice is still pending.

  Review returns the changed-commit qualification handoff for executable
  setup, changed-output oracle, quiet-process and target-size corrections.
  A newly found fast-exit output/TMP cap race is repaired in 0362ul40 and is
  under independent review. No native measurements are claimed. Cleanup
  wave 55 remains in preflight; the owner has retired an obsolete scanner
  that pinned inactive leaves. No additional worktree has been removed, and
  nothing is pushed.


- Coordinator raw-codec integration and cleanup update (2026-10-25):
  Required raw-profile signature enforcement merges at `c51a539757` from
  reviewed `939d401609`. The checker resolves local and external nominal
  limits types, rejects ambiguous identities, and verifies exact required
  arguments. All 55 composed checker tests pass. XLSX drawing limits and
  every current consumer merge atomically at `b21bb1ff82` from `34e6b0c6c`;
  its 20 unit tests, 12 integration tests and all-target/all-feature clippy
  pass. Root verifies the final source tree matches the tested tree and the
  composed glossary policy parses successfully.

  Ink's fast-exit output/TMP cap repair merges at `267ea71526` from reviewed
  `18da76e3fb`. All eleven composed boundary tests pass, including exact-cap
  acceptance and final snapshots after child exit. The earlier systemd
  preflight remains attributed to its original wrapper; this repair does
  not claim a new native measurement.

  Cleanup wave 55 removes the inactive, merged constructors checkout without
  force and retains its branch. All 21,662 ignored/untracked files, including
  its target, are preserved with exact identities. Sparse restoration passes.
  Record 0362ul53 binds root review, the external manifest and a payload-based
  regeneration command. The allocation difference after quarantine and its
  supplements is 7,535,771,648 bytes; no isolated free-space delta is claimed.
  No active source root, target or foreign process is removed.

  Raw-codec repairs continue for DOCX/OOXML, PPT, DOC/OLE, XLS, ODF, image
  codecs, fonts, SpreadsheetDrawing, crypto and RTF; iWork result names also
  require consumer validation. DOC/OLE's scoped suites and clippy pass and
  await independent source review. DOCX and ODF were returned for concrete
  nested-budget and owner-policy defects; corrected commits are under review.
  Four bounded build slots are allocated to DOCX, XLS, RTF and image codecs.
  Tooling-row and PPT qualification repairs remain under review. Final source
  capture, full gate, signed-batch contract choice and actual performance
  qualification remain open. Nothing is pushed.

- Coordinator tooling, review and cleanup update (2026-10-25):
  Reviewed tooling environment and process-lifecycle repairs merge at
  `f86d2e90e3`. The lifecycle regression now observes child exit without
  reaping it, preserving the actual `Popen` return code; 43 composed tests
  pass. PPT qualification guards and the strengthened forward/inverse replay
  oracle merge at `a10687e3b2`; 44 focused composition tests pass. The
  unchanged timed workload remains separate from the correctness companion.

  The full current tooling row then passes on clean `a10687e3b2`: 1,687 tests
  in 110.974 seconds, with zero target bytes, empty Cargo trace, unchanged
  source/locks and no remaining owned processes. Record 0362ul55 binds the
  independently verified receipt and preserves the separate setup failure.
  Both runner and external `.coord/gate.sh` now register the same 35 modules.
  This is one row of the 53-step gate, not full acceptance.

  Independent source review returns the raw-codec lanes for concrete repairs:
  DOC/OLE output admission and opaque Custom XML fidelity (0362ul45a), XLS
  tooltip and source-backed reader policy retention (0362ul46b), image
  compatibility-converter policies (0362ul47a), SpreadsheetDrawing raw
  append/rewrite budgets (0362ul51a), crypto output/typed-limit paths
  (0362ul52a), and PPT nested aggregate/editor limits (0362ul39b). RTF chained
  edits must retain their originating open options. ODF 0362ul44b repairs
  are source-ready for review. DOCX 0362ul42b now passes independent source
  review at `25ed716942`; actual corrected Cargo validation remains required.
  Fonts' automatic-loader fix is preserved, while bounded subset planning
  remains engineering work; unproven allocation multipliers are not accepted.

  The earlier XLS clippy failures were caused by absent tracked
  `clippy.toml` in the sparse checkout. Configured current and original-base
  checks pass, matching the frozen gate's policy; no lint exemption is added.
  New source-review findings still prevent accepting that lane. Four bounded
  validation slots are assigned to DOC/OLE, image codecs, iWork and RTF;
  PPT and DOCX source work does not launch competing builds.

  Cleanup wave 56 removes the clean, merged XLSX rich-transfer checkout,
  retains its branch and quarantines its sole ignored lockfile. Sparse
  restoration and root postchecks pass. Record 0362ul54 reports approximately
  7.55 GB of checkout-allocation reduction after the complete quarantine.
  No target, source capture, foreign process or protected probe is removed.
  Final capture, full gate, the signed-batch contract decision and actual
  native/performance qualification remain open. Nothing is pushed.

- Coordinator image-codec integration (2026-10-25):
  Reviewed atomic imgconv leaf `aaeaa79a0b` merges at `566451176d`. EMF and
  WMF compatibility playback retains the caller-selected parser profile;
  two explicit output-limit regressions and final all-target/all-feature
  clippy pass. Root verifies the final tree equals tested e392, all terminal
  receipt hashes, unchanged lock, target below 1.54 GB and completed process
  cleanup. Record 0362ul47 distinguishes predecessor facade checks from the
  final source validation and preserves the earlier configuration/lint
  failures. The composed glossary policy keeps both XLSX and imgconv rows.

  The freed slot is assigned to source-reviewed DOCX `25ed716942`. Crypto
  `6c3277b18e` also passes its focused independent source review and awaits
  bounded Cargo validation. ODF 0362ul44c is returned for redundant full
  scans with large token reservations and remaining authored-fragment/output
  budget gaps; this performance regression is not accepted. All remaining
  source fixes, full-gate and performance obligations stay active. Nothing
  is pushed.

- Coordinator iWork naming and validation update (2026-10-25):
  Canonical iWork removal-result names merge at `21fafcc599` from final
  `0524a814ae`. Root verifies the seventeen identifier migrations and exact
  source equality with pinned `375f55e8df` apart from documentation. The
  pinned Rust 1.95 iWork and ops/facade all-target Clippy runs pass; retained
  log/trace digests, source identity, lock identities and target guards match.
  Record 0362ul48 preserves the missing-fixture failure and explicitly states
  that only selected terminal receipt fields, not a full preflight, survived.
  All 82 composed glossary, iWork API and benchmark-matrix tests pass.

  The attempted matrix classification exposes a separate guard defect:
  numbered follow-ups such as `0362ul48` are not recognized by its record-ID
  pattern. A dedicated tooling leaf is fixing the pattern and auditing every
  newly recognized evidence directory; no measurement is relabeled as source
  validation. The unsuccessful classification check remains in the record.

  Four bounded validation slots now belong to source-reviewed DOC/OLE
  `1678388376`, DOCX `4046be273f`, SpreadsheetDrawing `e71f7b95e2`, and
  crypto `6c3277b18e`. XLS `378a311c74` passes independent source review
  after separating physical CFB limits from logical worksheet scan limits.
  Corrected ODF `6861ee48f9` and PPT `f0bf93fea1` are undergoing source
  review. RTF's scoped suites pass, but its standalone performance consumer
  still needs the new required open options and a locked build before merge.
  The font subset adapter remains under repair for complete allocation
  accounting and separate output/working limits; no dependency lock change
  or native execution is approved by this checkpoint.

  Cleanup wave 57 is preflight-only. The previous two completed waves retain
  all quarantined evidence and reclaim about 15.1 GB of checkout allocation.
  Final capture, the full gate, signed-batch contract choice and actual native
  qualification remain open. Nothing is pushed.

- Coordinator DOC/OLE, RTF and evidence-coverage integration (2026-10-25):
  DOC/OLE raw-codec limits and fidelity repairs merge at `fa833c7d1b` from
  atomic `d6cc88337c`. The final source differs from tested `983ab2cfb2`
  only in its record; root reviews the small compile/Clippy correction delta
  and verifies all seven receipts. The full suites pass 1,246 DOC tests
  (two ignored) and 238 OLE tests, alongside seven focused limit/fidelity
  cases and all-target check/Clippy. Target storage stays below 541 MB.
  Record 0362ul45 retains every diagnostic attempt and the final identities.

  RTF merges at `7a369a1580` from atomic `1d273acbf2`, including the missed
  standalone performance-harness consumer. Its required open-options value
  is prepared outside the timed loop and its labels describe that call.
  Ten RTF/core/ops/facade validation rows pass; the additional locked
  standalone all-target check passes on corrected `a41a0dc163`. Root verifies
  source equality, all receipt artifacts, unchanged locks and completed
  process cleanup. Record 0362ul49 binds the actual scoped checks without
  inferring a native measurement from compiler runtime.

  The numbered evidence-record guard repair merges at `479a3278db` from
  reviewed `4cc7a78fbb`. It recognizes numbered follow-ups, rejects bare
  five-digit record names and gives all 29 newly recognized directories
  explicit evidence classifications. The composed matrix/glossary/iWork
  checks pass all 83 tests. Native allocator observations remain diagnostic;
  no timing or performance acceptance is manufactured by classification.

  Cleanup wave 57 retires the merged iWork naming checkout, retains its
  branch and preserves all eight ignored files. Root verifies payload
  identities, actual sparse restoration, protected roots and final removal.
  Record 0362ul56 reports 221,061,120 bytes of net checkout allocation
  reduction after the complete quarantine. External validation artifacts
  remain intact, and the integration hold is released.

  Slot 1 now belongs to corrected XLS validation; DOCX, SpreadsheetDrawing
  and crypto keep slots 2–4. ODF's strict XML fidelity is repaired; its latest
  script-wrapper namespace/depth correction is under review. PPT still
  needs aggregate admission to count nested payloads exactly. DOCX's actual
  library run found two modern-comment output-length mismatches, which are
  being corrected with failed evidence retained. Fresh full-gate capture,
  signed-batch contract choice, font budget adaptation and final native
  qualification remain required. Nothing is pushed.

- Coordinator composed tooling and remaining validation update (2026-10-25):
  The fresh 35-module tooling row runs all 1,688 tests on clean `319690ffc6`
  and reports one failure in 108.883 seconds: DrawingML's reviewed adapter
  hash still names the earlier performance-harness library. RTF's canonical
  streaming rename changes exactly two source paths in that library. The
  historical DrawingML adapter copies the library into older source trees,
  so record 0362ul57 must preserve both historical and current API roles
  while deliberately rebinding the reviewed bytes. Updating a digest alone
  would not establish historical-build compatibility. All failed artifacts
  remain retained; target and Cargo trace stay empty and owned processes
  terminate. This current tooling row is failed, not accepted.

  Its external state is
  `.coord/tooling-row-319690ffc6-20261025T-state/state.json` (20,236 bytes,
  SHA-256 `3eaa75e98885aa3c249fcd49846f6e6c1721ecfd9dac6fab838c68bc34b8569d`).
  The 636,346-byte child log has SHA-256
  `1f52db64e95036ff40de1ef51bc5793fd5b085f0b9a8c86991f2b79323d1bf59`.
  Root verifies those artifacts; the integration hold is released.

  ODF `6672bde714` passes independent review of the remaining script-wrapper
  limits and has a prepared validation plan, awaiting a build slot. Crypto's
  compile/Clippy corrections at `3c843b9e9f` also pass source review; its
  final guarded run remains active. Its eventual DOC/OLE composition must
  retain the selected custom-XML policy and migrate the new writer arguments.
  PPT is correcting the final shape-click aggregate projection after repairing
  existing-list accounting. The new Allsorts adapter `1bb2e207ec` is under
  independent review, including a separate audit of pinned decompressor
  allocations; no dependency-budget exception or lock update is accepted.

  SpreadsheetDrawing atomic `9f43932dd0` has passing bounded diagnostic runs,
  but those runs did not use the required ownership/target guard. Root verifies
  that the approved f5 runner is available and returns the packet for actual
  guarded validation in its existing slot. Diagnostic results and temporary
  directory failures remain preserved. Its final source delta also receives
  independent review. Slots 1–4 remain XLS, DOCX, SpreadsheetDrawing and
  crypto. All final capture, gate and performance obligations remain active;
  nothing is pushed.

- Coordinator corrected tooling result (2026-10-25):
  DrawingML's explicit historical/current adapter repair merges at
  `264316098e` from reviewed `1c210ac24e`. Fourteen focused protocol tests
  pass. The full current tooling row then passes on clean `1bc3769f43`:
  1,689 tests in 107.073 seconds, zero target bytes, empty Cargo trace,
  unchanged locks and no remaining owned processes. Root verifies all
  artifact identities; record 0362ul58 binds the result and retains the
  previous failed row through 0362ul57. This remains one row of the full gate.

  PPT `5a3d8ee950` now passes independent source review of exact recursive
  aggregate admission and valid-source removal at exact record/payload
  ceilings. It joins ODF in the validation queue. SpreadsheetDrawing's
  final source delta is independently ready, while guarded validation remains
  required. Crypto is validating the remaining encrypted DOCX/PPTX/ops
  consumers. The font adapter now includes a pinned exact zlib workspace
  query and is under source review; no lock refresh or final performance
  acceptance is claimed. Four bounded validation slots remain active, and
  nothing is pushed.

- Coordinator XLS and SpreadsheetDrawing integration (2026-10-25):
  XLS atomic `2ffacb16e9` merges at `517029ec66`. Root verifies all scoped
  checks and the complete-corpus library run: 1,334 tests pass, with unchanged
  locks and completed process cleanup. The prior sparse-fixture failures
  remain recorded. Its canonical validation type and the explicit historical
  DrawingML benchmark adapter migrate together; reviewed adapter ancestry
  merges at `6510c215ba`. The composed standalone harness check remains queued.

  SpreadsheetDrawing atomic `9f43932dd0` merges at `d14f3613e5` after
  independent source review and eleven passing actual f5 rows. Root verifies
  all command logs and Cargo traces, source and lock identities, final target
  size of 826,790,824 bytes and empty terminal owned-process lists. Record
  0362ul51 corrects the earlier runner-availability report and keeps the
  diagnostic TMP disappearance unexplained. The composed glossary/matrix/
  historical-adapter checks pass all 81 tests; 62 glossary profiles validate.

  Cleanup wave 58 retires the merged DOC/OLE and RTF worktrees while retaining
  branches and all six ignored payload files. Actual sparse restoration and
  root identity checks pass. Record 0362ul60 reports 7,592,120,320 bytes of
  net checkout-allocation reduction after the complete quarantines. External
  validation evidence and the protected constructor probe remain intact.

  The font leaf now assembles cumulative staging charges, typed refusal
  snapshots, DOCX error preservation and the reviewed distinction between
  structural failure and policy refusal. A direct dependency path keeps the
  Allsorts allocation query and decoder on the same vendored zlib-rs instance
  for standalone consumers. New root and standalone lock identities still
  require guarded validation; root records the reviewed source transitions
  and installs the new locks only in the isolated font leaf. Source review
  returns the remaining public glyph-mapping/subsetter APIs for required
  limits, a selected loader limit for typed propagation, and real preparation
  boundary tests. Public capabilities must remain available; an ambient-only
  budget or privatization is not accepted.

  DOCX's corrected modern-comment sizing passes independent review and the
  full library suite. Broader checks still require complete tracked fixture
  hydration. PPT's copied-payload test now has an independently derived exact
  288-byte boundary and a 287-byte refusal; the earlier failed guesses remain
  preserved. ODF is correcting missing required-limit consumers exposed by
  actual compilation. Crypto's final encrypted-consumer checks pass, but
  review finds that ops unlock drops selected encryption profiles for DOCX,
  PPTX and XLSX. That lane is returned for retention fixes and focused tests
  before atomic composition. The four bounded validation slots remain active.
  Fresh integrated capture, full gate, the pending signed-batch contract choice
  and actual native/performance qualification remain required. Nothing is pushed.

- Coordinator DOCX and PPT integration (2026-10-25):
  DOCX atomic `f9e3bff3b3` merges at `192197f975`, with the missed
  standalone PPTX validation call and historical adapter in the same commit.
  Adapter ancestry merges at `805c08facf`; record 0362ul57b preserves the
  historical full harness library through six exact spelling reversals.
  Root verifies 1,525 passing DOCX library tests, affected all-target check
  and final Clippy, including logs, traces, locks and process cleanup. The
  final source differs from the library-tested source only by an equivalent
  test initializer. The composed standalone harness check is assigned
  slot 3 on a fresh capture of `805c08facf`, with its original tool lock.

  PPT atomic `6a1317ab04` merges at `2771017a58` and contains the full leaf
  on its actual common ancestor `939d401609`. Its final tree equals preserved `ec014b1592`; independent
  review accepts the compile/error conversions and Clippy corrections. Root
  verifies the exact 288/287 boundary, focused profile tests, 1,267 passing
  library tests with one ignored, affected check and final Clippy. Record
  0362ul39 retains the earlier failed compile and boundary attempts, and
  distinguishes its scoped validation from actual native qualification.

  Current-source audit confirms remaining PPTX default/no-policy raw APIs.
  Three source workers own required returns 0362ul61 (laser, transition and
  shape), 0362ul62 (content codecs), and 0362ul63 (view/presentation properties
  and model3d). These are engineering work, not accepted exceptions; all
  consumers, owner-selected profiles and glossary signatures migrate together.

  Font source `1511dbc9ea` passes independent review of the required raw
  limits, private cumulative session and exact staging rules. Its actual f5
  run passes all eight standalone budget-counter tests, then fails with
  thirteen vendor compile errors. The vendor owner takes over the combined
  font leaf and slot 2 for finite repairs with failed evidence preserved.
  ODF retains slot 1; crypto retains slot 4 for selected encryption-profile
  retention. The generic f5 runner needs no font-specific relaxation. Final
  qualification needs new versioned policies bound to the new path sources
  and separate historical/current locks; historical evidence stays unchanged.
  The integrated gate, actual native/performance qualification and pending
  signed-batch contract choice remain open. Nothing is pushed.

- Coordinator standalone consumer verification (2026-10-25):
  The actual guarded standalone `perf-baseline` all-target check passes on
  clean composed `805c08facf`. Root verifies source, locks, log/trace hashes,
  target size of 368,735,177 bytes and completed process cleanup. Record
  0362ul64 closes the compiler check queued for XLS and DOCX; it does not
  claim coverage of the later PPT, crypto, font or new PPTX source. Slot 3
  is released. Native timings and the final integrated gate remain open.

  PPTX 0362ul61 is source-ready for independent review. The other two PPTX
  raw returns remain in source work. Font vendor compile repairs through
  `06aa4f8897` and crypto retention/classification repairs through
  `106e48ea98` pass independent source review; their actual guarded runs
  continue. ODF's latest retained-profile and opaque-vendor XML changes are
  under review. Nothing is pushed.

- Coordinator font and encryption integration (2026-10-25):
  Crypto atomic `7d765ad6e0` merges into the integration branch at
  `f1816486ec`. Independent review confirms all selected encryption profiles
  survive unlock and later output for DOCX, PPTX and XLSX. Root verifies the
  actual unlock regression, affected all-target check and Clippy receipts,
  including logs, traces, source/lock identities and empty owned-process lists.
  The merge retains the selected DOC/OLE custom-XML profile and required
  writer arguments. Record 0362ul52 preserves the earlier failed attempts.

  Font atomic `07680db969` merges at `b8897649d6`, preserving the complete
  reviewed first-party and vendor changes on their true common ancestor.
  Root verifies all eleven actual guarded rows on `e69d108830`: vendor budget
  checks, font tests/Clippy, affected format checks/tests and both glossary
  profiles. The largest sampled target is 5,610,859,304 bytes; all final
  ownership samples are empty. The atomic tree differs from tested source
  only by the final receipt added to record 0362ul50. The old root lock is
  preserved externally; integration now uses reviewed root `efff87af6c`,
  performance-tool `de1f8ae53b` and native-tool `989da70f10` locks. Records
  0362ul59 and 0362ul59a document the three registry-to-path transitions.
  Both merges pass the 75 selected glossary/matrix/adapter tooling tests.
  A fresh composed check is assigned slot 4, including standalone consumers.

  The current-source census finds additional required policy migrations:
  RTF validation/compression (0362ul65), Numbers wire codecs (0362ul66), VBA
  (0362ul67), XLDM (0362ul68), formula codecs (0362ul69), and Markdown owner
  options (0362ul70). These remain engineering work, not exceptions. RTF's
  production changes pass source review and receive slot 3; its historical
  benchmark adapter must migrate atomically. Markdown is returned to preserve
  recipient policy through cross-copy replay/history, and VBA is returned
  for legacy PPT and complete signature coverage. PPTX returns 0362ul61–63
  and ODF validation remain active; ODF also needs the remaining tracked-change
  and formula serializer policies. Source-only census tooling 0362ul71 is
  returned for fail-closed anchor/input validation. New DrawingML protocol
  0362ul72 will bind separate historical/current source and lock identities;
  historical evidence is not relabeled.

  Cleanup wave 59 is authorized for four merged leaves with branches and
  ignored payloads retained; completion and measured savings are pending.
  Final integrated capture, complete gates, actual native/performance
  qualification and the signed-batch contract choice remain open. Nothing
  is pushed.

- Coordinator cleanup and composed consumer return (2026-10-25):
  Cleanup wave 59 is complete. Root verifies all 18 preserved payloads,
  actual sparse restoration, retained branches and absent registrations for
  the merged XLS, SpreadsheetDrawing, DOCX and PPT leaves. Record 0362ul73
  binds the receipts and the net allocation reduction of 8,465,526,784 bytes.
  External evidence and the protected constructor probe remain unchanged;
  the integration hold is released.

  Fresh composed source `b8897649d6` passes check and Clippy for fonts,
  crypto, DOC and DOCX, then fails PPT check because an integration-only
  correctness example still calls the now-private snapshot constructor.
  The failed capture and all nine receipts remain preserved and root-verified.
  Return 0362ul50a migrates both reads to the public options-based constructor
  at `200a3e5af3`; it preserves default parsing and outside-clock receipt
  behavior. Slot 4 continues on a fresh corrected capture; downstream checks
  and standalone compilation are still required. This failure is not accepted
  as an exception or reported as a full composed pass.

  Raw-coverage tooling 0362ul71b now passes independent review and eleven
  synthetic tests. Its manifest deliberately remains incomplete, and final
  mode rejects it. DrawingML's frozen-v1/current-v2 boundary passes review,
  but 0362ul72b still requires review of trusted-input ordering and at-use
  historical-library identity checks. Additional source workers prepare new
  signed and legacy/PPT qualification policies (0362ul74–75); no actual native
  execution or final source qualification is claimed. Nothing is pushed.

- Coordinator raw census integration checkpoint (2026-10-25):
  Reviewed 0362ul71b merges at `6b95d89554`; all 86 selected census,
  glossary, matrix and historical-adapter tests pass. Its final mode remains
  fail-closed until the complete frozen rustdoc census and dispositions are
  populated. Tooling-row registration is assigned as 0362ul76 without
  changing the approved runner used by active validations.

  PPT consumer return `200a3e5af3` passes independent source review:
  the public options-based constructor is exactly equivalent to the private
  default helper. Slot 4 continues the corrected composed check. DrawingML
  v2 `2b0bc89552` passes independent review and 28 source-only tests, including
  fixed historical input, disjoint paths and complete vendor-tree checks;
  candidate binding deliberately remains unset. It will compose with the
  RTF consumer change after actual validation. Further reviews return VBA,
  XLDM and formula changes for concrete retained-policy, typed-error and
  preallocation gaps. Those fixes and actual acceptance remain required;
  nothing is pushed.

- Coordinator feature coverage and lock-provenance return (2026-10-25):
  Census tooling registration merges at `8ea056c5dc`. The 53-step gate now
  includes 36 tooling modules; its changed runner line only adds the census
  test module. Root verifies the exact one-line delta, fourteen selected tests
  and synchronized external gate recipe. Active scoped runs keep their
  original approved runner; historical receipts remain unchanged.

  Markdown source `19767b4467` passes seven actual guarded Markdown, ops and
  facade rows. Root verifies logs, traces, target guards and empty terminal
  owned-process lists; independent review accepts its final compile/Clippy
  delta and populated-history regression. Its atomic packet awaits final
  coordinator composition. Slot 2 moves to reviewed VBA `604bce4fc5`.

  RTF's guarded suite used locally regenerated root lock `558fafc4fd`, rather
  than approved `4100b1fbb2`. Root independently confirms six registry-version
  substitutions and rejects that run as acceptance. The source and original
  receipts remain preserved; slot 3 reruns with the exact approved lock and
  a separate provenance correction. No dependency regeneration is accepted.
  Its standalone compiler failure also exposes two unmigrated PPT calls
  behind the non-encryption feature branch. Return 0362ul39a at `9d8b8e5a8c`
  passes the retained source limits to both raw parser calls. Feature-specific
  and standalone reruns remain required; all-feature passes did not cover
  those branches. Final gate and performance qualification remain open.

- Coordinator RTF, Markdown and benchmark integration (2026-10-25):
  Markdown atomic `e8b359eb6a` merges at `d86898161a`. Root verifies its
  seven passing guarded rows, source equality, retained replay/history policy
  and empty final process ownership. RTF atomic `d80cbb4fb6` merges at
  `ee7e9955d3` with the complete reviewed DrawingML-v2 packet and both
  standalone PPT benchmark repairs. The approved-lock RTF rerun passes all
  nine rows, including 1,539 tests; root verifies source, locks, logs, traces,
  target guards and terminal cleanup. The earlier regenerated-lock receipt
  remains rejected and preserved.

  Composed `200a3e5af3` passes all ten PPT/PPTX/XLSX/ops/facade check and
  Clippy rows across its preserved attempts. Minimal/non-encryption PPT check
  and Clippy pass on `9d8b8e5a8c`, whose standalone failure exposes two more
  benchmark callers. Return 0362ul39b supplies the required default parser
  profile and the equivalent public snapshot constructor; independent review
  confirms unchanged behavior and timer boundaries. Fresh composed `ee7e9955d3`
  then passes the actual guarded standalone all-target compiler check. Root
  verifies its complete receipt and 2,851,946,744-byte final target; record
  0362ul77 binds this result. These are compiler checks, not performance runs.

  DrawingML ancestry merges at `45c18b4e08`; the protocol uses the exact
  historical library rather than reversing later candidate migrations. The
  signed-native and current-native protocols merge at `2cce4e55c1` and
  `dcd5709b5c`; a standalone Python import failure is repaired at `0dc6c76634`
  and preserved in 0362ul75e. The candidate bindings remain unset. The tooling
  row now has 39 modules with 53 gate steps unchanged. The 113 selected
  integration tests and 47 protocol/wiring tests pass; a fresh complete
  tooling row is running. Crypto 0362ul52 receipts are explicitly classified
  as correctness/compiler evidence; all 3,700 measurement rows and 150 open
  misses remain unchanged.

  Further review returns PPTX replay for destination-profile enforcement:
  equal-content replay must apply the recipient's selected semantic limits.
  PPTX content also needs complete source-backed profile retention and
  fallible allocations. Numbers-wire 0362ul66a, XLDM 0362ul68b and formula
  0362ul69c pass source review; actual validation remains required. The four
  build slots now belong to ODF, VBA, PPTX properties and XLDM. A newly found
  output-symlink escape in the current-native validator is returned as
  0362ul75f for independent review; no native acceptance is claimed.

  Cleanup wave 60 is authorized for the merged font, crypto and Markdown
  worktrees after their owners retire all stale audit groups. Completion and
  measured savings remain pending. Full integration acceptance, complete
  gates, native/performance qualification and the signed-batch contract choice
  remain open. Nothing is pushed.

- Coordinator cleanup wave 60 completion (2026-10-25):
  The merged font, crypto and Markdown worktrees are removed non-forced.
  Root verifies all seven quarantined payloads, retained branch tips, absent
  registrations, actual sparse restore and the unchanged protected probe.
  Record 0362ul78 binds the receipts and conservative net reduction of
  7,873,409,024 allocated bytes including complete quarantine overhead.
  Cleanup releases its integration hold; the current tooling row retains
  its separate source hold until completion.

  ODF completes five actual guarded validation rows, including all 2,135
  library tests, strict XML and script-envelope regressions. Root verifies
  logs, traces, locks, source, target guards and empty terminal ownership;
  final source and runner-provenance reviews remain pending. Slot 1 moves
  to formula; the other build slots remain VBA, PPTX properties and XLDM.
  Current-native validator return 0362ul75f passes independent source review
  and awaits merge after the tooling hold. Final acceptance remains open.

- Coordinator ODF and VBA integration (2026-10-25):
  ODF atomic `677093e42f` merges at `cbd0a7587e`. Independent review accepts
  the final balanced-fragment, selected-limit and consumer corrections. Root
  verifies all 104 intended changed paths and exact leaf blobs outside the
  additive glossary merge. The five actual rows pass, including 2,135 library
  tests; final receipts bind source `32253785a5`, all three approved locks,
  logs, traces and empty terminal ownership. The captured runner's execution
  and guard functions match the approved runner; its differences are other
  plan rows. The record corrects decimal target units and explicitly accepts
  compiled-target reuse for these compiler/test rows. Native/performance
  freshness remains required.

  VBA atomic `4a22c3c136` merges at `ebc7425644`. All six actual guarded rows
  on `c68801c837` pass: 48 VBA tests, five PPT VBA writer tests, 1,268 PPT
  library tests with one ignored, affected check/Clippy and the feature-boundary
  check. Root verifies every log/trace, source and lock fence, target guard and
  terminal ownership sample. The three earlier compile failures remain
  recorded. Conflict resolution retains the reviewed public correctness
  projection and all font/ODF/VBA glossary profiles; the glossary now has 66
  profiles. The next composed performance-tool and native-tool compiler checks
  are queued for the next available build slot.

  Native-validator symlink repair 0362ul75f merges at `daf712bcc1` after
  independent review and all 18 focused tests. The full tooling row on
  earlier clean `0dc6c76634` finishes with 1,745 tests, two failures and one
  error. Root verifies its failed receipt, unchanged source/locks, empty
  Cargo trace, zero-byte target and terminal cleanup. The concrete returns
  are intentional policy strings rejected by the fixture guard, a lease test
  whose child event file is missing, and a fast-exit TMP test observing SIGTERM.
  Record 0362ul80 repairs are undergoing independent review; the failed row
  is preserved and must be rerun after integration. No full tooling pass is
  claimed for the current source.

  The final raw census plan covers every direct library package, but review
  requires receipt-bound profile outputs and trusted source/lock identities.
  Return 0362ul79a implements those checks and is under review; actual final
  rustdoc capture and human dispositions remain required. Current build slots
  are formula, Numbers wire, PPTX properties and XLDM. PPTX source replay and
  content aggregate-limit returns remain engineering work. Integration is
  clean with no source hold. Final gate/native/performance qualification and
  the signed-batch contract choice remain open. Nothing is pushed.

- Coordinator tooling repair and composed follow-up (2026-10-25):
  Record 0362ul80 merges at `317f648242` after independent review and all
  60 focused tests pass in the composed tree. The fixture guard uses exact
  policy-line exceptions with decoy negatives; lease and fast-exit tests now
  coordinate the intended paths deterministically. The actual complete
  39-module row is assigned a fresh detached capture of that source. The old
  1,745-test failure remains preserved; current full-row acceptance is pending.

  Both standalone performance-tool and native-tool all-target checks pass
  on clean `ebc7425644`, covering the ODF and VBA consumer composition. Root
  verifies logs, target traces, source/lock identities and empty terminal
  ownership; the target peaks at 3,095,985,461 bytes. Record 0362ul77 now binds
  both composed captures. Slot 3 is released.

  PPTX properties' nine guarded rows pass on `a75cf6419d`; root verifies the
  complete receipt. Later core replay/profile repairs and their new strict
  recipient regression remain source-only and require combined validation.
  Further audit confirms metadata public raw codecs still use fixed resource
  ceilings without required profiles. Active returns 0362ul81–82 cover those
  structural/history families, including every consumer and recipient replay
  policy; they are required work, not exceptions. Their source workers share
  the existing 61–63 profile-carrier design and receive no build slot yet.
  Census 0362ul79c is closing validator bytecode self-pollution after its
  receipt/guard/lock binding review. No final census or acceptance is claimed.

- Coordinator census merge and tooling provenance return (2026-10-25):
  Receipt-bound raw census 0362ul79c merges at `11048ed3cb` after independent
  review and all 76 selected census/fixture-guard/glossary tests pass. It binds
  package/features/command, source/tree, locks, toolchain, capped execution,
  target trace and rustdoc artifact identities; lock symlinks and ignored
  overlays are rejected. Bytecode generation is disabled before local imports.
  The final profile matrix, captures and reviewed dispositions remain pending;
  an incomplete manifest still fails final acceptance.

  The next full tooling test process reports 1,746 passing tests, but its f5
  wrapper correctly rejects the receipt because the worker ran against mutable
  integration instead of the requested detached source. The source advanced
  from `317f648242` to `11048ed3cb`; unchanged locks and empty final ownership
  do not make that run admissible. Root verifies and preserves the provenance
  failure. Return 0362ul80a requires an actual detached source checkout and
  coordinator verification of its root, commit, branch, locks and command
  before another run. Integration has no source hold.

  Formula's final source/Clippy cleanup passes independent review and its
  validation receipts await coordinator verification. PPTX core 63 replay
  repairs pass source review, with actual combined validation still required.
  Metadata returns 81–82 and the remaining PPTX replay/aggregate fixes remain
  active. No full gate, final census or native/performance acceptance is
  claimed; nothing is pushed.

- Coordinator verified tooling and XLDM checkpoint (2026-10-25):
  The complete 39-module tooling row passes on the independently verified
  detached source `11048ed3cb`: 1,752 tests. Root verifies the saved state,
  child log, empty Cargo trace, exact runner and locks, unchanged detached
  source and terminal ownership. Overall state is correctly `partial-passed`
  because only row 48 of 53 was selected. Record 0362ul80a preserves the
  earlier setup and mutable-source failures separately; this is a tooling-row
  pass, not a complete gate pass.

  XLDM final source `c5424e23e4` passes independent review. Root verifies ten
  accepted guarded rows across the `87e241633e` and `c5424e23e4` captures,
  including XLDM/XLSX/XLSB tests and affected format/ops/facade checks and
  strict Clippy. All log and trace digests, approved locks, source fences and
  terminal ownership match. Peak target usage is 3,957,146,335 bytes. The
  earlier failed rows remain preserved; an atomic integration proposal and
  precise decimal-unit receipt correction are being prepared.

  Formula's final source review is ready, but root has found only child logs
  and target traces in its stated capture directory. Saved row receipts must
  be located or the affected rows recaptured before acceptance. PPTX returns
  61e and 62b are active alongside metadata 81–82; combined source review and
  guarded validation remain required. Numbers wire continues its scoped
  validation. Cleanup wave 60 reclaimed 7,873,409,024 allocated bytes after
  final quarantine overhead; wave 61 is read-only preflight for the merged
  ODF, VBA and RTF leaves. Integration has no hold. Nothing is pushed.

- Coordinator XLDM integration and cleanup wave 61 (2026-10-25):
  XLDM atomic `c2495b728f` merges at `11d14a661d`. Root verifies the full
  31-path change set and all 29 non-shared leaf blobs; the package/glossary
  merges retain sibling changes and all 67 profiles. All 89 selected
  glossary, inventory, census and fixture-guard tests pass. A fresh detached
  composed compiler capture is being prepared; native and performance
  acceptance remain open.

  Formula's missing row receipts are confirmed. Three fresh actual guarded
  rows on `c937149a3b` now pass with saved manifests and per-row receipts:
  Formula tests, affected DOC/DOCX/ops/facade consumers, and strict Clippy.
  Root verifies source, locks, runner, every log and trace, target guards and
  empty terminal ownership. Peak target usage is 5,479,852,530 bytes. The
  earlier child logs remain historical evidence without reconstructed state;
  the complete atomic Formula proposal is being prepared.

  Cleanup wave 61 removes the merged ODF, VBA and RTF worktrees non-forced,
  retaining their source/tested/receipt/atomic refs and external validation
  evidence. Thirteen ignored files pass same-device sparse restoration with
  hashes, modes and inodes preserved. Root independently verifies the receipt;
  conservative net reclamation is 15,193,243,648 allocated bytes after all
  final quarantine overhead. Record 0362ul83 retains the exact evidence.
  The temporary integration and trunk holds are released.

  PPTX structural metadata 81 is committed and undergoing independent source
  review. Returns 61e/62c and history metadata 82 remain active. Review found
  processing-instruction offset and repeated-node-scan issues in collaboration
  replay; notes/media aggregates also still need affected-scope accounting.
  Those findings remain required engineering work. No complete gate, census,
  native or performance acceptance is claimed. Nothing is pushed.

- Coordinator Formula integration checkpoint (2026-10-25):
  Formula atomic `20aa27d913` merges at `f14235da95`. Root verifies the
  complete 44-path lane, exact tested Formula/consumer blobs, equivalent ops
  error-classification changes and additive glossary composition. All 89
  selected tooling tests pass; the glossary now has 68 profiles. Fresh
  persisted validation remains bound to `c937149a3b`, with the earlier
  receipt gap explicitly documented. No unrecorded rerun is treated as
  acceptance evidence.

  A partial detached compiler capture on `11d14a661d` is preserved after
  cancellation for retargeting; it is not a completed composed result. The
  next capture must use detached `f14235da95`, with coordinator verification
  of its manifest before launch. Numbers wire is validating `6417a19464`
  and its post-review compile repairs are undergoing independent review.

  History metadata 82 is committed and under independent review. Structural
  metadata 81 was returned for remaining compatibility wrappers, allocation
  before output-limit admission and typed limit errors. Shared PPTX replay
  accounting and no-op base matching remain under review. Those returns
  remain active engineering work, with no exception accepted. The workspace
  cleanup holds are released; the protected constructor probe is unchanged.
  Final gate, census and native/performance qualification remain open.

- Coordinator composed compiler and source-review checkpoint (2026-10-25):
  All nine composed compiler rows pass on detached `f14235da95`. The four
  spreadsheet/ops/facade checks pass first; the following row stops on an
  invalid DOCX feature flag in the command. The corrected five-row
  continuation passes Formula/DOC/DOCX, the scoped facade/ops consumers, and
  both standalone tools. Root verifies the actual receipts and preserves the
  failed attempt. Record 0362ul85 binds all commands, logs, traces and source
  identities. Final target usage is 3,687,603,132 bytes; the slot is released.
  These are compiler checks, with no native execution or timing claim.

  Numbers `6417a19464` reaches focused tests and exposes an output-preflight
  budget error in two mixed scalar batches. Source correction `67613f5dd0`
  passes independent review: it uses the selected tile output ceiling while
  retaining independent work/input/tail accounting. Its actual guarded rerun
  remains in progress; no final validation result is claimed yet.

  PPTX history return 82a at `fac5957f84` passes scoped source review, pending
  the shared typed-error migration. Record 0362ul84 owns that migration and
  every consumer of the root limit errors. The 61 physical accounting and
  no-op identity repair is under review. The 62 authoring-index repair passes
  review, while theme-edge accounting, missing-target behavior, history-chain
  cost and eager payload decoding remain required returns. Structural 81a
  continues alias and preallocation repairs. All PPTX work still requires
  combined review and actual validation before integration.

  Census reconciliation now derives 68 glossary profiles plus 16 package-only
  profiles, separately from the 68-package all-feature inventory. Its initial
  claims of unresolved ODF/CFB/ZIP limit twins were stale; root checked current
  signatures and returned the record for correction. No duplicate engineering
  scope or generic exemption is accepted from those historical claims. Only
  redaction audit 22 remains deferred. Nothing is pushed.

- Coordinator Numbers integration and cleanup checkpoint (2026-10-25):
  Numbers-wire atomic `0f098a25d9` merges at `6fe60b8d9a`. All 16 actual
  guarded rows pass on `67613f5dd0`, including 177 wire tests, 157 focused
  consumer tests, strict Clippy and seven glossary profiles. Root verifies
  every receipt, child log, target trace, source/lock fence and terminal
  ownership sample. Peak and final target usage is 4,660,470,149 bytes.
  The earlier mixed-scalar failure remains preserved. All 94 implementation,
  consumer, test and existing record blobs match the validated leaf; only
  the two additive glossary files and the new composition record differ.
  All 99 selected glossary, inventory, census and fixture-guard tool tests
  pass after correcting a mistyped test-module name in the initial command.
  The glossary now has 69 profiles; the raw-census matrix becomes 69 plus
  15 package-only profiles, still separate from the 68 all-feature inputs.

  Cleanup wave 62 removes the merged Formula and XLDM author worktrees,
  preserving refs, ignored payloads and external evidence. Root verifies
  the five payloads and same-device restoration receipt. Net allocation
  reclaimed is 15,120,351,232 bytes after final quarantine overhead; record
  0362ul88 binds the evidence. Trunk/integration holds are released and the
  protected constructor probe is unchanged.

  PPTX returns 61h (`6139650eea`), 62g (`790f694f88`) and 81c
  (`bf41322310`) pass independent scoped source reviews. The fixes cover
  recipient reservation transfer, lazy policy-aware laser accounting,
  reusable no-op proofs, persistent replay-index deletion, exact output
  admission, opaque custom-show preservation and deterministic guide edits.
  History 82a and core 63 remain ready in their reviewed scopes. Typed-error
  work 84 remains active: generic count dimensions and fabricated saturated
  observations must be replaced with actual typed dimensions and counts.
  Combined source integration, actual tests and performance remain required.

  Corrected census record 79d and shared OOXML audit 87 are merged. Audit
  87 confirms public semantic publication through `&mut OpcPackage` in
  properties/custom XML, web/Ribbon and VBA; workers 87a–87c must migrate
  those paths and every consumer to Edit/Commit/replay. Initial source
  proposals remain under review and are not accepted. ZIP alias work 86
  removes three confirmed compatibility aliases and migrates their callers.
  Its first guarded row stops on workspace formatting differences in 335
  files, including 112 vendored files. The failed capture is retained;
  changed-file formatting repair and fresh scoped validation are pending.
  Full workspace formatting remains required before the final gate.

  Final census, complete gate, native/corpus and performance qualification
  remain open. Only redaction audit 22 is deferred. Nothing is pushed.

- Coordinator ZIP acceptance, PPTX composition and cleanup checkpoint (2026-10-25):
  ZIP atomic `04879dd503` merges at `76b2d275f9`. All eight actual guarded
  rows pass on `e49a1062ec`, including 948 tests; root verifies every receipt,
  source/lock fence, log, target trace and terminal ownership sample. All 22
  implementation/prior-record blobs match tested source, glossary composition
  retains 69 profiles, and 99 selected tooling tests pass. Record 0362ul89
  binds acceptance. The earlier 335-path workspace formatting failure remains
  preserved and full formatting remains required before the final gate.

  Cleanup wave 63 removes the accepted ZIP author and atomic worktrees
  without force, preserving branches, two ignored payloads and external
  evidence/target. Root verifies same-device sparse restoration and final
  receipts. Net reclaimed allocation is 15,131,107,328 bytes; record 0362ul90
  binds the evidence. Trunk/integration holds are released and the protected
  constructor probe is unchanged.

  Private PPTX composition joins 61h/62g/63/81c/82a, retaining all 13 selected
  profiles, physical byte accounting and persistent replay indexes. Its 99
  tooling tests pass; runtime acceptance is not claimed. Composition review
  fixes metadata publication bookkeeping and policy loss during deferred/SVG
  reopening, with a committed regression still to run. The 89a/b laser/MCE
  return is composed: aggregate staging stays atomic and prepared-slide
  admission precedes allocation. Typed-error 84a/b remains in source review
  and compiler preparation; exact cross-copy transcript regeneration is open.

  Shared properties and Web/Ribbon pass scoped source review. Actual
  properties validation preserves a formatting failure, then a stale private
  helper compile failure; both have committed repairs. Web/Ribbon's initial
  preflight refuses an overlapping build before running a row, so isolated
  captures are serialized. VBA87c remains under recipient-policy and lexical
  fidelity repair. Final gate/docs, census, native/corpus and performance
  qualification remain open. Only redaction audit 22 is deferred. Nothing is
  pushed.

- Coordinator typed PPTX and shared OOXML compiler checkpoint (2026-10-25):
  Private typed composition `49f9b2fcac` and dependent-Part accounting merge
  `2915ef1762` are committed. All 99 selected source tool tests pass; actual
  combined compiler/tests/Clippy and performance remain pending. Review fixes
  exact resource dimensions, empty-element depth admission and the insertion
  capacity regression without an owner-wide allocation. Records 0362ul89c,
  0362ul89e and 0362ul91 bind the source preparation and failed captures.

  Shared properties removes obsolete publishers and migrates remaining test
  consumers; test-only codec helpers use cfg(test). Fresh source `841d46e353`
  follows a preserved Slot-field compile failure. Web/Ribbon's common check
  passes; leaf checks expose XLSX return-shape and DOCX SharedOpc adoption
  errors, repaired through `beac39a4c5`. Fresh guarded runs remain required.
  VBA87c continues mandatory source-fidelity, exact archive accounting and
  duplicate-scan repairs; no capability/performance exception is accepted.

  Isolated captures run serially. An accidental properties glossary CLI build
  is acknowledged and recorded, and its interference produces a retained
  zero-row Web/Ribbon refusal. Source/review work continues in parallel.
  Cleanup63's 15.13 GB reclamation is complete; later inventory is read-only.
  Final full formatting, gate/docs, census and native/corpus/ABBA remain open.
  Only audit22 is deferred. Nothing is pushed.

- Coordinator census preparation, compiler returns and cleanup checkpoint (2026-10-25):
  Current census preparation merges at `cd29aa8bda`, deriving 69 glossary plus
  15 package-only profiles and keeping the separate 68-package inventory.
  Dirty-input and stale-policy binding returns are fixed and independently
  reviewed. All 107 selected tooling tests and the matrix check pass on
  integration. Evidence arrays remain empty and state remains incomplete.
  Final inventory guidance requires actual F5 outer ownership/isolation and
  resource receipts. Record 0362ul92 binds this checkpoint.

  Combined PPTX formatting passes all 384 changed Rust paths; compiler errors
  expose three unnecessary OPC mutable bindings. Private repair `00a42a7d4c`
  is committed with record 0362ul89f; its fresh compiler run remains queued.
  Web/Ribbon passes every compiler row, then repairs two Clippy defaults at
  `c877dee369`; its fresh validation is in progress. Properties repairs two
  test result qualifications at `ac7b91e8d8`; its fresh plan is prepared.
  VBA `b62c2e5359` passes source review with exact generated-root metadata
  accounting and boundary regressions; corrected inert validation is prepared.
  No runtime or performance result is inferred from these source repairs.

  Cleanup64 removes the merged imgconv worktree without force, preserving its
  branch, four ignored payloads and both external validation roots. Actual
  restoration and root verification pass; net allocation reclaimed is
  148,025,344 bytes. Holds are released and the constructor probe is unchanged.
  Final full formatting, gate/docs, fresh inventories and dispositions,
  corpus/native and quiet ABBA remain open. Only audit22 is deferred.
  Nothing is pushed.

- Coordinator formatting, shared-error and cleanup checkpoint (2026-10-25):
  Formatting leaf `ad8f6465c7` merges at `ad15b44bff` after independent
  parent-preimage proof and an actual full `cargo fmt --all -- --check` PASS.
  The retained first retry found ten post-capture Numbers paths; record
  0362ul93a fixes them. Final formatting is still required after pending
  source merges. Census tooling registration merges at `16d3ea672b` and
  all 28 focused Python checks pass; the external gate row is updated.

  Private PPTX `237d6801a6` commits the reviewed compiler repairs and every
  caller. Its fresh isolated check clears prior type/borrow errors and
  exposes 113 dead-code errors; helpers are being audited for removal or
  test-only gating with every public capability preserved. Shared OOXML
  composition `6bf2a5f68a` joins Web/Ribbon and VBA. A coherent typed error
  return must fix visible corrupt_file misclassification. Properties
  `15863b6354` is returned for remaining source/invariant distinctions,
  tolerated-content refusal, stale XLSB assertions and canonical docs.

  Cleanup65 removes the merged census worktree without force, preserving
  its branch, ignored payload and external evidence. Actual restoration and
  root verification pass; net allocation reclaimed is 1,555,513,344 bytes.
  Holds are released and the constructor probe is unchanged. Record 0362ul95
  binds these results. Final gate/docs, inventories/dispositions, corpus/native
  and quiet ABBA remain open. Only audit22 is deferred. Nothing is pushed.

  Follow-up source composition: PPTX `6517f7f5c1` removes the reviewed
  private helpers; `2749ced58b` fixes two subsequent unused imports. Its
  actual check now reaches integration tests and reports missing limits in
  two transition writer callers. The complete test/example audit is active;
  no all-target success is claimed. OOXML `0db7ff3a6a` commits the reviewed
  properties/shared-error composition, including eight conflict resolutions,
  the five-variant properties bridge, source chains and typed patch errors.
  Custom XML audit 94c and actual combined validation remain required.

- Coordinator all-target consumers and cleanup checkpoint (2026-10-25):
  PPTX's actual keep-going check on `15d8d53a0c` passes formatting and
  exposes 130 library/integration-test diagnostics. Five workers repair the
  consumers and five independent reviews pass their scoped source changes.
  Private composition `096422a20a` commits all 36 Rust paths with records
  89p–89u; seven paths receive proven formatting-only normalization. Fresh
  compiler/runtime/Clippy/glossary and performance validation remain required.

  OOXML's corrected v2 packet on `0db7ff3a6a` passes formatting and fails
  on one unused OPC binding. Private fix `0395ca40f4` preserves the same
  generated-part predicate. Custom XML 94ca remains returned for bounded
  allocation, opaque/MCE fidelity, selected-limit propagation and scan-cost
  repairs; 94cb is active. The 65-row combined plan remains unexecuted.

  Cleanup66 removes the merged formatting worktree without force, retaining
  its branch, ignored lock, both captures and targets. Actual sparse restore
  and root verification pass; net allocation reclaimed is 7,565,520,896 bytes.
  Holds are released and the constructor probe is unchanged. Record 0362ul96
  binds this checkpoint. Final gate/docs, inventories/dispositions,
  corpus/native and quiet ABBA remain open. Only audit22 is deferred.
  Nothing is pushed.

  Subsequent actual compiler checks now pass: PPTX `02b2489ead` clears
  all-features/all-targets after 89v/89w repairs; OOXML common `5ec78a4858`
  clears all-features/all-targets after 94e/94f. Root verifies complete
  receipts and terminal ownership for both. Runtime, strict Clippy, glossary,
  downstream-consumer and performance qualification remain open. Custom XML
  94cb remains returned for paired-element indexing, typed limits and bounded
  MCE/attribute handling; 94cc is active and uncomposed.

- Coordinator qualification, typed-consumer and cleanup checkpoint (2026-10-25):
  PPTX 89x/89y/89z passes independent source review and composes at
  `3181342f0c` (records `2da172ece2`). Actual qualification clears the original
  76 Clippy findings, then reports three test-only findings. Independently
  reviewed `ae6f32d202` fixes them; fresh qualification is running. Prior
  failed captures and complete root receipt verification remain retained.

  Shared OOXML `95973e8c6a` composes reviewed Custom XML through 94ce,
  DOCX through 94gb and XLSX through 94hc, with all consumers. It preserves
  lexical/MCE fidelity, exact nested property-limit kinds and classification
  through XLSX joins. Intermediate classification, coverage and offset-type
  gaps are returned and fixed before composition. Actual combined validation
  remains required; neither private candidate is accepted yet.

  Cleanup67 removes two retired PPTX compiler/test worktrees without force,
  preserving their branches and exact changed blobs. Sparse restoration and
  root verification pass; full-wave net recovery is 61,091,840 bytes. Holds
  are released and the protected constructor probe is unchanged. Record
  0362ul97 binds this checkpoint. Final gates, census/dispositions,
  corpus/native and quiet ABBA remain open. Only audit22 is deferred.
  Nothing is pushed.

  Subsequent actual results: PPTX `ae6f32d202` passes strict Clippy, then
  its full-feature library reports 1,356 passed, 23 failed and one ignored.
  Root verifies the complete failed capture. Six reviewed repair chains
  compose privately at `7637126651` (89ak); notes recipient-proof and
  aggregate-limit returns remain active, and the digest golden still requires
  isolated regeneration/review. No runtime acceptance is claimed.

  OOXML's invalid inert command-kind packet is returned before GO. Corrected
  v2 on `95973e8c6a` passes formatting and fails with 16 compiler diagnostics.
  Independently reviewed 94j fixes them, and root fast-forwards the private
  candidate to `491137abe2`. Its fresh eleven-row capture is running;
  formatting and common compilation report PASS, with terminal verification
  pending. Final acceptance obligations remain unchanged.

- Coordinator notes admission, consumer repairs and cleanup (2026-10-25):
  OOXML `491137abe2` passes format/common/OPC, then the actual leaf check
  exposes three remaining consumer errors. Root verifies the terminal failed
  capture. Independently reviewed 94k `5e89400413` repairs them and is
  fast-forwarded privately; fresh isolated validation is active. The first
  prepared packet's author-worktree binding is returned before GO.

  PPTX 89aj passes source review for selected raw notes/master/theme aggregate
  admission without additional reads or copies. 89aca follow-ups repair
  selected physical policy, deferred root-only admission and exact proof
  identity, with positive-control regression coverage. Runtime and cold/warm
  performance acceptance remain pending; no exception is accepted.

  Cleanup68 removes two retired PPTX worktrees without force, retaining
  branches and all evidence. Root verifies full-chain restoration and an exact
  formatter preimage; net recovery is 15,132,827,648 bytes. Holds are released
  and the protected constructor probe is unchanged. Record 0362ul98 binds this
  checkpoint. Final gates, census/dispositions, corpus/native and ABBA remain
  open. Only audit22 is deferred. Nothing is pushed.

  Subsequent verified result: OOXML 94k passes all six formatting/compiler
  rows, including all four format crates, ops and facade. Strict Clippy then
  stops on two OPC findings; 94l is repairing them. Root verifies every
  completed receipt and terminal session. PPTX `3377583e2b` composes the
  independently reviewed final notes-proof chain and aggregate repair (89al).
  Fresh compiler/Clippy, actual golden generation/review and runtime/performance
  qualification remain required.

  Later compiler follow-ups: reviewed PPTX 89am clears the three production/
  library-test findings, and reviewed 89ao `e10039df1c` clears 18 integration
  test qualifications. Both actual failed captures are retained and verified.
  OOXML 94l clears the OPC findings, then common Clippy reports 16 more;
  reviewed 94m `305a7c0ee6` fixes them and its fresh six-row run is active.
  The PPTX compiler/Clippy/raw-golden run is queued; no golden is changed.

  Cleanup69 removes three retired OOXML author worktrees, preserving refs,
  all ten changed blobs and the ignored lock through actual restoration.
  Root verification passes and holds are released. Net recovery is
  22,698,254,336 bytes, or 37,831,081,984 bytes across cleanup68/69. The
  separate notes-memo performance harness proposal is under source review.
  Runtime, performance and final acceptance remain open; nothing is pushed.

- Coordinator golden and runtime qualification checkpoint (2026-10-25):
  PPTX `e10039df1c` passes formatting, minimal all-target compilation, strict
  Clippy and actual golden generation. Independent review confirms all 192
  row keys and all 140 digests are unchanged; exactly 18 typed limit records
  change. Root commits the reviewed artifact privately at `ef9470f7da` (89ap).
  Its full/minimal runtime and glossary packet is active after root preflight.

  OOXML `305a7c0ee6` passes common/OPC Clippy, then reviewed 94n/94na fix two
  XLSX findings and restore repository formatting. The actual `11c3f2fc8c`
  run reaches one XLSB exhaustive-match lint. Its complete failed receipt is
  root-verified; reviewed 94o is fast-forwarded privately. Unrun rows
  remain open. Performance harness review and cleanup70 audit continue.
  Record 0362ul99 retains the evidence. Final gates, census/dispositions,
  corpus/native and ABBA remain open; only audit22 is deferred. Nothing is pushed.

  Subsequent PPTX runtime result: 1,374 library tests pass, 12 fail and one is
  ignored. Root verifies the terminal capture and returns notes/presence/writer
  repairs as 89aq/89ar/89as. Reviewed 89as fixes the slide-sync opening quote
  and is fast-forwarded privately; other repairs remain active. OOXML 94o's
  fresh scoped Clippy capture is running. No runtime acceptance is claimed.

  Cleanup70 removes two retired PPTX worktrees without force, preserving
  refs and all four restored blobs. Root verification passes and holds are
  released. Net recovery is 15,133,167,616 bytes, or 52,964,249,600 bytes across
  cleanup68–70. The protected probe, all targets and evidence remain intact.

  OOXML `3774fe4374` now passes all four scoped format/leaf/ops/facade Clippy
  rows with complete root verification. Its 38-row runtime packet is active;
  the 13 glossary profiles remain separate. PPTX merge `437da0c7f5` composes
  reviewed notes 89aq and downstream consumers 89at without conflicts.
  Presence 89ar is returned for allocation, date-time, inactive-threading and
  physical-attribute defects; 89ara is active. No performance acceptance is
  claimed, and all final acceptance obligations remain open.

  The first OOXML runtime packet stops before compilation on an unsupported
  helper option. Root verifies the failed receipt and corrects the CLI audit;
  the fresh v2 packet is running. Reviewed notes performance harness 89and is
  merged privately at `cc2ab3f79d` (89av), with 23 Python tests passing after
  composition. Baseline source preparation continues; no benchmark or runtime
  acceptance is claimed. All changes are committed and nothing is pushed.

  Corrected OOXML runtime result: 511 common-library tests pass and 17 fail;
  root verifies the terminal capture. Four workers own the VBA, Custom XML,
  properties and ribbon returns (94p–94s). PPTX presence 89ara is returned for
  physical validation order and added scan cost; 89arb is required. Baseline
  `a4c28413cc` carries the identical reviewed performance harness on unchanged
  763 production source. No runtime or performance exception is accepted.

  Reviewed VBA 94p and Custom XML 94q compose privately at `8c766ac5c6`
  (94t), addressing fourteen reported failures pending rerun. Properties 94rb,
  Ribbon 94sa review and PPTX presence 89ard remain active. A fresh eight-row
  PPTX compiler/consumer packet is running on `cc2ab3f79d`; the initial inert
  packet was corrected before GO. Completed changes are committed; nothing
  is pushed. Final acceptance remains open.


- Coordinator reviewed runtime-repair composition (2026-10-25):
  All outstanding comments, properties and ribbon chains pass independent
  source review and merge privately without conflicts. PPTX current tip is
  `fc5ef75a9e` (89ay); OOXML is `b73ee18a69` (94u). The preceding PPTX packet
  stopped at formatting before Cargo. Root corrects the harness and comments
  layout and the premature composition formatting statements; all five PPTX
  and thirteen OOXML changed Rust paths now pass actual formatting checks.

  Baseline `84ba3b5c33` has all eight harness blobs identical to the PPTX
  candidate while retaining original 763 production/configuration/fixtures.
  Fresh source-bound qualification packets are being prepared; no new runtime
  or performance acceptance is claimed. Record 0362ul100 binds this checkpoint.
  Four retired worktrees are authorized for guarded cleanup71; results are
  pending. The protected probe is unchanged. Completed changes are committed;
  final acceptance remains open and nothing is pushed.


  The actual fc5 compiler packet passes formatting then reports three comments
  compiler diagnostics. Root verifies the terminal failure and independently
  reviews/fast-forwards compiler-only repair `2ac6bca410` (89az). Fresh compiler
  and runtime bindings are required. Cleanup71 is complete and root-verified:
  four retired worktrees removed, twelve blobs restored, branches retained,
  7,692,996,608 bytes recovered; cleanup68–71 totals 60,657,246,208 bytes.
  The live-holder 89at tree and all external captures/targets remain intact.


  Subsequent PPTX minimal all-target compilation passes; strict Clippy finds
  one redundant test-helper conversion. Reviewed one-line 89ba is fast-forwarded
  at `afede08087`, with fresh qualification prepared. OOXML's new publication
  code needs two imports; reviewed import-only 94v is fast-forwarded at
  `f167986c74` and its fresh compiler packet is active. Both failed captures
  are root-verified and retained. Baseline `6cb65f6358` now includes the
  previously missing vendored release inputs with unchanged production and
  identical harness. No runtime or performance exception is accepted.


- Coordinator joint source integration (2026-10-25):
  Reviewed PPTX `afede08087` merges privately at `75988e5eec` (101a), then
  reviewed OOXML `df57ad24df` joins it. Root resolves the one production
  conflict by preserving shared metadata Commit extraction and PPTX's fallible
  byte accounting. All 499 changed Rust paths pass formatting without edits;
  glossary JSON retains 69 unique profiles and 33 signature-bearing profiles.
  Records 0362ul101a/101 retain the composition details. Independent final-merge
  review and joint qualification follow; this private merge is not acceptance.

  The f167 OOXML Clippy run stops on redundant qualifications introduced by
  broad imports. Its complete receipt is root-verified; 94va qualifies only
  the new references and is composed after review. The standalone afede PPTX
  compiler packet continues unchanged; separate df57 packets remain inert.
  Subsequent validation will bind the combined source. Runtime, fidelity,
  performance and final acceptance remain open; nothing is pushed.


  Joint merge `e2a833d72e` passes independent source review: the metadata
  resolution and all fourteen overlap paths preserve both migration chains;
  the tree is clean. The next compiler packet is being bound to this exact
  source. This remains private integration pending full qualification.


  The standalone afede PPTX packet completes all eight compiler/consumer
  checks successfully, including strict leaf/ops/facade Clippy, combined
  DOCX/PPTX examples and standalone perf/native harnesses. Root verifies the
  complete terminal capture; target ends at 957,738,123 bytes. This result is
  scoped to afede. Joint e2a qualification is being prepared, with no runtime
  or performance acceptance inferred. Baseline `c5d2b30740` carries the same
  eight harness blobs as e2a after its test-only consumer migration.


  The joint e2a packet passes formatting and minimal PPTX compilation, then
  direct primary-package Clippy finds one unnecessary closure in common
  application properties. Root verifies the failed capture and independently
  reviews/fast-forwards 101b at `6964022294`; the fixed error and bounds are
  unchanged. Fresh compiler bindings are required. Joint runtime coverage is
  consolidated to 33 rows and glossary coverage to 13 unique profiles/26 steps.
  The existing signed-batch contract choice is explicitly surfaced to the owner;
  no answer is inferred, and independent qualification continues.


- Coordinator signed-batch decision and local consolidation (2026-10-25):
  The owner selects immediate signed-batch refusal. Reviewed 102 partitions
  signed corpus checks from unsigned equivalence and proves zero item reads;
  it is merged privately at `b4f83c95b9`. Native policy binds that explicit
  selection. The preceding joint compiler capture on 6964 passes formatting
  and minimal PPTX compilation, then finds one redundant ribbon test qualifier.
  Root verifies the failure and fast-forwards reviewed 101c at `a3aec01d5a`.
  Fresh combined compiler/runtime/glossary qualification is required.

  The owner also requests merging or deleting every local unified-derived
  branch/worktree, excluding main and office-format-completeness. Root is
  already on feat/unified-ops-api. The initial inventory finds 482 refs and
  300 worktrees; lineage review distinguishes independent office siblings
  from unified work. Already-integrated refs are being checked for guarded
  deletion; unique work and active measurement baselines remain under review.
  Record 0362ul103 tracks this consolidation. No push or final acceptance.

  The first cleanup is complete: 70 already-trunk local branch refs are
  deleted with exact-tip guards and all commits remain reachable. Cleanup72
  removes retired W7 docfix without force, preserves its ignored lock and
  reclaims 10,166,382,592 bytes; root verifies the unchanged protected probe.
  Sixty-six independent Office/iWork sibling branches are outside scope.
  Remaining unique rows receive semantic review against their squashed
  replacements. The 14 signed-policy tests pass and the fresh eight-row
  compiler packet is running on frozen joint source `ea5c95e212`.

  Joint source `ea5c95e212` passes all eight compiler/consumer rows, including
  strict six-package/ops/facade Clippy and standalone perf/native harnesses.
  Root verifies the complete receipt; the 34-row runtime packet is now active.
  Cleanup73 removes eleven unused worktrees without force, preserves eight
  locks and reclaims 62,694,273,024 bytes. Cleanup72–73 totals 72,860,655,616
  bytes. Local ref consolidation has deleted 271 refs with exact-tip guards;
  integrated commits remain reachable, and seven retired old reference
  branches are deliberately discarded. Sixty-eight office-lineage siblings
  are outside scope. Dirty payload and remaining worktree review continues.


  Runtime on ea5 stops with 526 common OOXML tests passing and five failing.
  A diagnostic continuation passes eight more rows, then DOCX signature
  policy reports eight passes and one stale Ribbon error assertion. Both
  terminal captures are root-verified. Reviewed 104c, 104a and 105 address
  those findings; reviewed 102a preserves individual signed invalid-target
  errors while keeping the owner's zero-read batch refusal. Repaired-source
  qualification remains required; no runtime/performance acceptance.

  Local consolidation now deletes 318 refs and removes 47 worktrees across
  cleanup72–76. The measured allocated ledger is 118,948,712,448 bytes before
  the small extra adapter archive, excluding cleanup75's unmeasured allocation.
  Cleanup75 separately removes about 31.8 GB of apparent size after reversible
  quarantine. Root verifies preserved payloads, source/registry proofs and the
  unchanged protected constructor probe. Remaining worktrees are still being
  audited; 68 independent office siblings remain outside scope. No push.


  Four reviewed runtime repairs compose at 3c291 (106). Its first runtime
  row exposes missing Result wrappers in Custom XML 104a before tests compile;
  root verifies the failed capture and fixes both at e130 (104aa), with
  independent review. The next common suite compiles: 528 pass, three VBA
  fixtures/oracle assertions still fail and are returned as 104d. Fresh full
  qualification remains required; no runtime/performance acceptance.

  Cleanup77–79 removes 34 more worktrees and deletes 23 refs after exact-tip,
  source/payload/registry checks. Totals are 81 worktrees and 341 refs; the
  measured allocation ledger is 120,572,608,512 bytes, with the same wave75
  allocation and small adapter-archive caveats. The DrawingML deletion-only
  checkout's signed invalid-target coverage is now preserved by composed102a.
  Remaining validation worktrees receive explicit disposition review; the
  protected probe is unchanged. Completed changes are committed; no push.


  Further diagnostic runtime rows expose DOCX web-extension consumer and
  conformance failures, ten PPTX library failures (1,386 pass, one ignored),
  and one signed ops ordering assertion. Seven repair chains pass source
  review: 104e, 107, 108a, 110, 111a, 112 and 113. The proposed 111 aggregate
  relaxation is returned; 111a preserves production limits and corrects the
  stale fixture boundaries. Signed batch zero-read refusal remains settled;
  113 tests the distinct closure/individual error order explicitly. The
  next continuation passes ten XLSX/XLSB rows, then ops metadata reports five
  passes and two missing-property-context failures, returned as 114. The last
  four diagnostic rows run on frozen e130. Composition and full repaired-source
  qualification remain required.

  Cleanup80–82 brings fully verified removals to 106 worktrees. Another 21
  are removed in cleanup83; its scope-hash difference is proved comment-only,
  but the failed ignored-payload inventory prevents a complete preservation
  claim. All committed source remains reachable, and local ref deletions
  total 370. Cleanup85 then removes three more verified leaves and their refs:
  totals are 373 refs and 109 fully verified removals, plus cleanup83's 21 with
  the recorded inventory gap. The measured allocated ledger reaches
  228,248,707,072 bytes with the recorded exclusions; cleanup83 allocation is
  unknown. Further leaf cleanup continues. The protected probe and active
  baselines remain intact. Final acceptance stays open; nothing is pushed.


  The final four diagnostic rows pass on e130. Across those diagnostic
  packets, 28 original rows pass and five fail; the minimal PPTX profile
  remains unrun. Reviewed metadata context repair 114a joins the other seven
  repairs at private integration `0da85f43e6` (115), without conflicts. Root
  verifies all twenty changed crate files against their reviewed leaves and
  refreshes only the common crate's test dependency in the ignored root lock.
  Compiler/runtime/glossary and final qualification must now bind this repaired
  source. No result from e130 is relabeled as a repaired-source pass.

  Fresh compiler qualification on fcc04996 stops at formatting of one new
  metadata helper signature, before any Cargo row. Root verifies the terminal
  capture and commits the whitespace-only correction as 0362ul116. Repaired
  source must receive fresh compiler/runtime/glossary qualification.

  Cleanup through91 deletes 402 refs and removes 195 worktrees: 149 fully
  verified, seventeen with a recorded mode-evidence limitation, eight with
  the disclosed late holder-guard failure, and cleanup83's 21 with its
  ignored-inventory gap. All recorded source/payload and retained controls
  receive root verification; the protected probe is unchanged. The measured
  allocation ledger is 488,902,098,944 bytes under record103's exclusions,
  including no allocation claim for waves88/90. Remaining ancestor snapshots
  and retired leaves are being consolidated. No push or final acceptance.

  Record117 removes the obsolete Web Extensions import found after the
  formatting return. Repaired integration `a8d32513b8` passes all eight joint
  compiler/consumer checks, including strict leaves/ops/facade Clippy and the
  performance/native harnesses. Root verifies the complete terminal capture;
  fresh runtime/glossary qualification follows on this frozen source (118).

  Consolidation reaches 422 deleted refs and 231 removed worktrees. Root
  verifies dirty overlays and old blob identities before deleting nineteen
  retired refs. The notes control is detached at measured `c5d2b30740`, with
  its old provenance branch archived/deleted and eight harness blobs verified.
  Only unified trunk, private integration and the active runner remain as
  unified-derived branches; main/office and 68 independent office siblings
  are untouched. The measured allocated ledger is 676,827,754,496 bytes under
  the recorded exclusions and evidence limitations. Further detached-candidate
  cleanup continues. Full runtime, workspace/wsdoc, API/fidelity and performance
  acceptance remain open. No push.

  The fresh a8 runtime packet stops in the common library with 529 passing
  tests and two VBA failures: relationship-size fixture accounting and source
  fixture publication. Record119 returns these for 119a diagnosis/repair;
  remaining groups continue diagnostically on frozen a8. No preservation or
  limit rule is waived. Compiler8 remains passing, full runtime stays open.
  Cleanup95's first seven snapshots are root-verified and removed, bringing
  totals to 422 refs and 238 worktrees. Its missing pre-allocation measurement
  is disclosed; the allocated ledger remains unchanged. Further stale after
  snapshots are reviewed while canonical before controls remain available.

  Common residual runtime exposes two Ribbon failures after the core/custom
  and markup/MCE groups pass; Web Extensions integration passes all five tests.
  Glossary finds the missing VBA Error export and an OPC rustdoc link. Returns
  120/120a are committed; 119b passes independent review after correcting the
  remaining ineligible fixture, and 121 is under replay/fidelity review (122).
  Diagnostics continue on frozen a8; no failed capture is relabeled. The
  full53 proposal is returned for using obsolete runner safety code. Cleanup95
  reaches 422 deleted refs and 247 removed worktrees, with archived dirty
  patches replayed and verified. The measured allocation ledger reaches
  730,280,955,904 bytes under the recorded exclusions. Three active repair
  branches remain for composition; final qualification and cleanup stay open.

  Further a8 diagnostics pass twelve common/OPC/DOCX groups, then return one
  DOCX durable tamper assertion. Follow-up123a uses typed errors throughout
  that test while preserving atomic publication checks. Three more DOCX
  groups pass; the PPTX helper passes 1,396 library tests (one ignored) and
  five integration chunks before a signed-policy consumer rejects typed
  VBA `Signed`. Reviewed124a corrects that consumer and its documentation.
  Ribbon121c passes replay, fidelity and ordering review after two further
  returns. Five reviewed repair leaves await composition; later diagnostics
  and all repaired-source acceptance remain open (125).

  Cleanup96 removes six archived DML/RTF diagnostic checkouts after actual
  payload restoration and patch replay. Canonical before controls remain.
  Root verifies 253 total worktree removals and 422 ref deletions, preserving
  prior evidence limitations and disclosing process visibility limits. The
  conservative allocated ledger reaches 737,962,319,872 bytes; this is not a
  filesystem free-space delta. Final API census requires fresh 84-profile
  captures and the 68-package inventory with the approved current lock.
  Root remains on `feat/unified-ops-api`; no push or final acceptance.

  Twenty-four further PPTX chunks pass before readonly replay exposes a real
  physical XML event-admission gap. Reviewed126 applies the existing raw
  scanner before semantic parsing and migrates readonly, math and metadata
  callers atomically. Its added bounded pass still requires performance
  qualification. Six source-reviewed leaves await composition. A guarded
  diagnostic continuation preserves aggregate failure while examining the
  remaining rows; primary-text-box and streaming-authoring failures are
  returned as128/129. No repaired-source pass is claimed (127).

  Cleanup97 removes three archived obsolete candidates. Root restores all
  seven payloads and verifies absent paths/registrations, bringing totals to
  422 refs and 256 worktrees. Process visibility limits remain explicit. The
  conservative allocated ledger is 753,344,364,544 bytes under the recorded
  exclusions. Raw census preparation is returned for three receipt/identity
  defects; no census execution is authorized by a plan alone. Final gates,
  fidelity, performance and consolidation remain open; nothing is pushed.

  The remaining43 packet terminates with 24 passing rows and three failures.
  The minimal PPTX profile reproduces readonly126 after chunks0–30 pass;
  its remaining tests stay unrun. Root verifies the complete capture and
  diagnostic admission decisions. Reviewed129 corrects exact InvalidLimit
  classification. Root returns128/128a for five adjacent memory/output
  dimension bugs and a regression-constructor error;128b is required before
  composition. The final sixteen outer diagnostic rows are separate (130).

  Census adapterv5 passes root/independent review and nine build-free tests
  after correcting the real inventory receipt schema; final-source binding
  and execution remain open. Cleanup98 removes one deletion-only obsolete
  DrawingML checkout after exact manifest/history and permission-aware guard
  checks. Totals reach257 worktrees and422 refs. Preserving its full deletion
  manifest costs57,864,192 net allocated bytes, reducing the ledger to
  753,286,500,352 bytes under the same exclusions. No push or final acceptance.

  All sixteen final diagnostic rows pass on a8. Eight reviewed repairs,
  including the completed 128c memory-dimension fix, compose without conflicts
  in private integration. Root verifies all sixteen changed crate files.
  Records131/131a and the refreshed census matrix produce frozen candidate
  C=780cf4c7ee; its matrix check passes with the same three locks. Fresh
  compiler qualification has started. No old result is relabeled as a C pass.

  Cleanup100 removes the obsolete DOC toggle candidate while preserving its
  lock and canonical comparison baseline. Cleanup99 then removes all eight
  merged repair worktrees and deletes their exact refs after fresh root
  guards. Totals are 266 worktrees and 430 refs; the allocated ledger reaches
  777,580,584,960 bytes under its recorded exclusions and evidence limits
  (132). Main/office and independent office branches remain untouched.
  Final qualification, performance and trunk integration stay open; no push.

  Frozen C passes all eight compiler/consumer rows; root verifies the terminal
  capture and starts the fresh 34-group runtime suite after preflight (133).
  Glossary26/full53, final census and fidelity/performance remain open.
  The reviewed 126 performance adapter and rebound notes packet pass their
  synthetic checks; actual measurements remain unrun. Census v7 is returned
  for capturing its binding hash too late; v8 must seal it before execution.

  Cleanup101 removes two archived obsolete ODP duplicates while preserving
  the canonical comparison endpoint. Root verifies 268 removed worktrees,
  430 deleted refs and a conservative allocated ledger of 785,182,834,688
  bytes under the recorded exclusions. A fresh branch census finds only
  trunk, integration and the active runner in scope. Main/office and their
  independent siblings are untouched. Final trunk integration is pending;
  nothing is pushed.

  C runtime passes sixteen common/OPC/DOCX groups, 1,398 PPTX library tests
  and integration chunks1–30, then returns the readonly no-op consumer (134).
  Physical event admission now passes; the later test incorrectly applies a
  pristine-package patch to different content. Return134a preserves strict
  base matching and corrects the consumer. Remaining runtime coverage is
  diagnostic; glossary26 runs separately on frozen C. No repaired-source
  runtime pass is claimed.

  Census v9 passes root/independent source review after binding-seal and
  recursion returns, but actual captures remain unrun. The 96-script/6,808-row
  corpus postcheck targets the existing full53 CLI row. Native policies
  72/74/75 still need executable v2 consumers, not just final identity pins;
  required engineering proceeds in separate protocol leaves137/136/135.
  Final gates, measurements, reviewed composition and trunk integration
  remain open. Nothing is pushed.

  C glossary passes nine rows, then returns fifteen DOCX minimal-profile
  MCE limits identity checks (138). Source signatures use the required public
  type; return138a investigates nominal/reexport proof without weakening the
  gate. Root verifies the failed capture and starts the separate 26-row PPTX
  full-feature runtime tail. Readonly134a is source-reviewed and ready; C
  stays frozen during diagnostics. Final qualification remains open.

  All 26 full-feature PPTX tail rows pass, including chunks32–53, six examples
  and doctests; root verifies the terminal capture (139). The minimal profile
  runs separately. Return138's explicit public-to-defining MCE type mapping
  passes independent review and 75 checker tests; it and134a await composition.
  Native135b/136a must repair actual execution/receipt paths;137a is under
  independent review after 43 focused tests. Final gates, measurements,
  trunk integration and leaf cleanup remain open. Nothing is pushed.

  The minimal PPTX diagnostic reproduces the known134 consumer failure after
  1,395 library tests and chunks1–30 pass. Root verifies it, then merges134a
  and138a without conflicts at private integration e7841dd222 (140). All five
  reviewed files and three locks match; 75 integrated checker tests pass.
  Final qualification uses full53 plus six compiler and five runtime profile
  supplements, preserving isolated wsdoc and complete glossary coverage.
  Native135c/136b/137c still need actual execution-path repairs before final
  composition; no native pass is claimed. Trunk and cleanup work continue.

  Cleanup103 removes both merged repair worktrees without force and deletes
  their exact refs after fresh root checks (141). Cumulative totals reach 270
  worktrees and 432 refs; the conservative allocated ledger is 800,318,455,808
  bytes under its existing exclusions, not a free-space measurement. All
  source blobs, three locks and the protected root probe are preserved.
  The 68 independent office refs stay outside scope; the 126 a8 baseline stays
  available. Three protocol leaves remain active. No push or final acceptance.

  Reviewed DrawingML137e composes at 2197a6c7f2; all 44 integrated tests pass
  (142). Separate protocol paths, ten imported tools and eight probe inputs
  are bound; final outer source pins and native execution remain pending.
  Cleanup104 retires its merged sparse leaf and exact ref, reaching 271
  removed worktrees and 433 refs. The allocated ledger is 800,323,493,888
  bytes under the existing exclusions. Final tooling coverage adds two new
  test modules to the existing forty. Native135d/136c and final qualification
  remain open; nothing is pushed.

  Reviewed current-native135e composes at bc5162cc97 without conflicts (143).
  The combined protocol/no-3rdparty suite passes 77 tests. Genuine baseline
  consumers, separate source builds and strict replay/output checks are in
  place; actual guarded builds and measurements remain pending. Integration's
  generated overlays are cleared after external preservation, with all three
  locks unchanged. The conservative allocated ledger reaches 800,926,519,296
  bytes under prior exclusions. Signed136d remains under repair; final source
  binding, qualification and trunk integration are still open. No push.

  The final format audit returns one PPT adapter import; 135f fixes it and
  migrates exact pins, with 33 focused tests passing (144). Final coverage is
  full53 plus seven compiler and five runtime supplements. Cleanup105 retires
  the merged current-native leaf, reaching 272 removed worktrees and 434 refs;
  the conservative ledger is 801,001,271,296 bytes under prior exclusions.
  The overwritten worker preflight failure is disclosed without attributing
  its unknown cause. Signed136d is reviewed; its identity label is being
  normalized. Final qualification and trunk integration remain open. No push.

  Signed136e is independently reviewed and merged at ee1bec60ec; the combined
  focused suite passes 67 tests (145). Candidate identity, all three lock
  graphs and generated consumer lock validation are bound; actual native
  execution still awaits final pins. Cleanup106 retires the last implementation
  leaf, reaching 273 worktrees and 435 refs removed. All 73 remaining heads
  are accounted for: three unified controls, main/office and 68 independent
  office refs. Required detached baselines remain. Final binding, qualification
  and trunk integration are open; nothing is pushed.

  Final candidate C=f8cd398a77 is frozen after the coordinator documentation
  merge and matrix refresh (145a/146). Its three locks and all 84 profiles
  remain unchanged. The reviewed raw-census v9 packet is bound to C and
  launched; the first two rustc identity steps pass. Capture, callable review
  and strict finalization remain open. Full65 qualification awaits its
  corrected executable driver. A separate sparse protocol worktree holds
  further native qualification fixes while C stays immutable. No acceptance
  or performance result is inferred; nothing is pushed.

  The C-bound raw census stops at inventory preflight after its two rustc
  identity steps pass (147). It supplied the outer tracing wrapper where
  inventory requires the coordinator cap shim; no profiles are captured.
  Root verifies the failed terminal receipts, unchanged source/locks and
  empty owned processes. Adapter v10 must repair the actual admission seam;
  the failed packet stays preserved and C stays frozen. All final gates and
  acceptance work remain open. Nothing is pushed.

  Reviewed135g commits on protocol P at 53451cc8d2; all 29 focused tests pass
  and Cargo outputs now match the recorded external binary paths (148).
  Full-gate v6-final passes independent review after command/path binding
  returns. Root starts its 65-step run on immutable C=f8cd398a77; changed-Rust
  formatting passes and the remaining rows are running or pending. Rawv10
  stays source-only while this build slot is occupied. Native measurements,
  census review, final trunk merge and cleanup remain open. Nothing is pushed.

  The first final65 attempt is incomplete: its driver is absent after three
  verified passing compiler rows, with no terminal result or known cause
  (149). Root preserves it and starts the same reviewed 65 checks from fresh
  state under a durable systemd user service. C and P remain unchanged.
  Notes and replay-scan126 packets now bind C, with 16 and eight focused
  tests passing; actual measurements remain unrun. Rawv11 source review
  passes, but its capture binding is still being prepared. Current-nativev3
  is returned for common-P admission; DrawingML repair continues. Final
  qualification, trunk integration and cleanup remain open. Nothing is pushed.

  Protocol137f commits at 4ab42bb967 with DrawingML's current candidate lock
  migrated across its policy and manifest; 26 integrated tests and six
  external mutation checks pass (150). The complete rawv11 C binding passes
  independent/root source preflight and remains inert behind the live gate.
  The gate's first four compiler checks pass. Native installer/driver returns
  remain under repair, and a capped source-archive job runs outside Git.
  Actual native measurements, final gates, census review, trunk integration
  and cleanup remain open. Nothing is pushed.

  Root verifies nine final65 rows: seven compiler supplements, workspace
  formatting and all-feature Clippy (151). DrawingML's C binding is generated;
  the exact 461,401-file C archive passes the reviewed verifier. Current-native
  v6 passes 22 focused tests and starts source binding against P4ab42bb.
  Signed v3 stops before archive generation on a lock permission assumption;
  reviewed v4 preserves C permissions, passes eight tests and starts fresh.
  Binding test migrations and DrawingML baseline preparation continue.
  The census remains 55 worktrees/74 heads, with four active unified controls.
  Native measurements, remaining gates, census review, final merge and cleanup
  remain required. Nothing is pushed.

  DrawingML's literal binding and consumer migration pass all 27 focused tests
  in an isolated P checkout (152); root verifies the exact combined patch.
  The signed test successor now requires production binding without a skip
  and keeps explicit pending negatives. Eleven tracked fixtures are restored
  in P, with unchanged content identities; its existing 14 policy tests pass.
  DrawingML v5 admits only the exact reviewed baseline overlays. Independent
  source checks and root's positive eleven-input check pass; the approved
  lock changes only the recorded VBA dependency edge. Seventeen final65 rows
  are verified; the gate and two source-generation services remain live.
  Common-P installation, actual measurements, final census, merge and cleanup
  remain open. Nothing is pushed.

  The signed v4 source-generation service subsequently dies at its 4G memory
  cap (152a); the journal confirms a kernel OOM kill. Its constructed tar is
  preserved but unverified. The protocol's whole-archive/member reads are
  returned for streaming repair, with a fresh verifier allowed to reuse the
  tar only after complete source checks. The current-native emitter and gate
  remain live. The baseline materializer is returned before real execution
  for full/sparse checkout metadata, unrelated-ref and completion-order fixes.
  P stays frozen until the live emitter finishes. No acceptance or push.

  The signed streaming repair passes independent review and 16 focused tests
  (153); it waits for P installation. The archive-reuse v5 recipe is returned
  for standalone finalizer path containment and immediate reuse identity
  checks. DrawingML materializer v5a passes 13 source/seam checks and root's
  actual read-only ten-profile plan. Its capped source-only execution starts;
  terminal verification is pending. P/C remain frozen during both source jobs.
  Root verifies 32 passing final65 rows; the gate remains live. Common-P
  installation, final gates, measurements, census review, merge and cleanup
  remain required. Nothing is pushed.

  Completion audit 154 returns C=f8cd398a77 for a mandatory BatchCommit
  accessor mismatch and 22 public deprecated format/support routes. Three
  isolated repair leaves (155–157) must preserve capabilities and migrate
  every consumer; stale detection/RTF prose also needs correction. The
  current-native binding and ten DrawingML baseline roots finish and pass
  root verification, but C's binding is held as evidence for that revision.
  Signed streaming repair 154a passes 25 integrated tests in P, with candidate
  bindings still pending. The live C gate remains useful reference evidence;
  final gates/census/native acceptance must bind the corrected candidate.
  No additional deferral, capability-loss waiver or push.

  The C gate then terminates with a real PPT hyperlink-add failure (154b):
  row 25's CRUD matrix reports corrupt_file; six later hyperlink cells are
  unreached. The same chunk's 13 DOC tests pass. Root verifies terminal logs,
  unchanged C/locks, target bounds and empty owned processes. Repair 158 has
  the single build slot for focused diagnosis. The three contract leaves and
  their follow-ups are committed and entering candidate integration, with
  compiler/fidelity/performance acceptance still required. Old-C packets
  remain historical; final qualification must bind the corrected candidate.

  The three contract repairs and their review follow-ups merge into candidate
  f17dda6e6f (154c), together with bounded signed-archive verification. All
  151 combined Python contract/protocol tests pass. Root retires the three
  merged leaves and refs after preserving their ignored locks, reaching
  276 worktrees and 438 refs removed; 62 worktrees and 75 heads remain.
  The 71-profile glossary requires a matrix refresh. PPT repair 158 retains
  the build slot; Rust, native and final acceptance remain open. C1 evidence
  stays historical, and the corrected candidate is not yet frozen or pushed.

  Census preparation 159 commits at ec478ed84f: all 68 packages remain covered
  by 84 profiles, now split into 71 glossary and 13 package-only profiles.
  All eight preparation tests pass, including checked-in determinism; capture
  evidence remains empty. The post-repair source audit finds no further
  mandatory API defect. Historical OfficeCLI v3 passes its source review and
  correctly refuses the stale source identity. PPT158 review and the complete
  corrected-source qualification remain open; no final acceptance or push.

  Candidate86a12758 includes coordinator records through159. Independent PPT
  review returns158 for a record/body unit mismatch, dropped payload limits
  and remaining default-policy resets (160); source-only158a repair proceeds.
  Earlier direct-Cargo observations lack retained approved-runner receipts
  and are not accepted validation. Root starts isolated, source-bound Rust
  checks for the merged IWA graph/index and owned PPT record changes. C stays
  fixed for that run; final source freeze, full gates and measurements remain
  open. Nothing is pushed.

  Focused Rust verification catches one graph test-consumer compile error;
  157b fixes the optional-iterator assertion without a production change.
  The fresh approved run on 49fc1721 passes 18 graph tests, 26 index tests and
  the owned PPT record fidelity test (160a). Root verifies logs, target
  traces, unchanged source/locks and clean process exits. PPT 158a and final
  corrected-source qualification remain open; no acceptance or push.

  PPT158–158h merges at ac36f3099c and the accumulated integration fast-forwards
  onto feat/unified-ops-api (161). Approved runs pass 1,249 PPT library tests
  (one existing ignored), nine OfficeArt tests, five integration tests, typed
  Ops limit classification and all-format CRUD with zero unreached claims.
  Earlier compiler failures remain preserved; no oracle is changed. Root
  retires the merged PPT checkout/ref and its test cache, reaching 277
  worktrees and 439 refs removed; 61 worktrees and 74 heads remain. The local
  untracked probe and old ignored lock are preserved before trunk integration.
  Rawv12a passes independent source review and 18 tests. Corrected C2 freeze,
  full65/isolated rustdoc, final census and native performance remain open.
  The remaining unified refs are merged controls. Nothing is pushed.

  The fresh C2 full65 run stops at its first formatting row (162): two
  BatchCommit consumer tests need assertion reflow. No build or test runs.
  Root preserves and verifies the terminal failure, applies the formatting
  correction and passes the same check across all 588 changed Rust files.
  Native and census recipes wait at their source-identity boundary. A fresh
  committed successor and gate state are required; no final acceptance or push.

  C2a freezes at 33a0a63d35 and its fresh full65 run passes all seven compiler
  supplements, workspace formatting and all-feature Clippy (163). Root
  verifies the initial nine rows; the gate remains live. Current, signed,
  DrawingML, raw-census, OfficeCLI, notes and scanner126 preparations pass
  source review after metadata/pin corrections. No archive, final census
  capture or native measurement is claimed. C and pending P stay frozen;
  coordinator documentation continues on trunk. Full qualification, protocol
  installation and remaining control cleanup remain required. Nothing is pushed.

  Root verifies 29 passing rows of the same live C2a gate (164). A completion
  checklist finds no additional source defect but preserves all outstanding
  census, native, fidelity and performance requirements; residuals are not
  accepted exceptions. The execution checklist corrects current-native to
  176 total ABBA legs across its two lanes. The refreshed 74-head census finds
  all four unified control refs already merged and still in qualification
  use; 61 worktrees remain, with no additional unused unified leaf. C and P
  stay frozen. Full65 and the final proofs remain open; nothing is pushed.

  Native scope reconciliation restores 25 required execution rows (165):
  17 historical groups become 18 rows after splitting OLE/XLS properties,
  then seven later rows are added. The recent three-policy shorthand did
  not close the earlier lanes; no final-C2a native receipt exists. W8 File's
  109-route C2a preparation exposes a stale six-edge dependency policy;
  its exact twelve-edge successor is being prepared. DOC/Ink preparation
  retains a returned historical-path test, and ODF raw codec still needs
  its paired protocol. The gate's 1,022 Ops library tests pass while its
  integration chunks continue. Root preserves and removes six generated
  bytecode files; frozen C has only its approved ignored lock. Full gate,
  census, all 25 qualification rows and final control cleanup remain open.
  Only memo 22 is deferred; nothing is pushed.

  Protocol admission review returns retired DOC HEAD checks, replaceable
  fidelity expectations and then recursive bind delegation (165a). DOC/Ink
  v5 passes independent source review with original oracles and positive
  bind/CLI tests. W8 v4 passes after path/hash repairs; its accidentally
  overwritten initial packet is restored exactly from retained tool-call
  patches and all original seals are independently verified. Neither packet
  is native acceptance. The isolated ODF raw-codec leaf commits at 48d68bcc29
  with seven Python checks and rustfmt passing; independent review and Rust
  builds remain open. The same full65 gate continues its Ops integration
  chunks. All 25 qualification rows remain required; nothing is pushed.

  Independent review returns the ODF raw-codec source leaf for eleven
  concrete findings (165b), including contradictory source/overlay bindings,
  a lazy-reader negative that panics, inverted no-op fidelity, incorrect
  timing boundaries, incomplete preservation proofs and unsafe preparation
  ordering. Seven Python checks did not cover these seams. The author repairs
  the isolated leaf; it is not merged or execution-ready. C/P remain frozen
  and the same full65 run stays active. No native exception or push.

  iWork constructor/owner preparations pass independent source review (165c);
  owner fixture provenance needs only one future-P metadata entry repaired.
  ODF raw-codec follow-ups fix source admission, fidelity and timing seams,
  then a Rust/Python descriptor mismatch; successor 8e400cbaee passes twelve
  source-only tests and independent review, with compilation still pending.
  Sheet v2 is returned for an incompatible e63 allocation consumer. OLE owner
  preparation runs against the real sources and passes 25 focused tests;
  its independent review and a new five-lane legacy review remain pending.
  Root verifies 32 completed gate rows while Ops chunks continue. The census
  retains 62 worktrees/75 heads, including the active ODF leaf and required
  qualification inputs; no unused unified leaf remains. C/P stay frozen, all
  25 actual qualification rows remain open, and nothing is pushed.

  The complete Ops suite passes all 132 chunks (165d), including 1,022 library
  tests and 217 doctests; root verifies 33 completed gate rows while facade
  tests run. Sheet v3 is returned for changed allocator controls; v4 reuses
  the exact preserved e63 adapter and awaits independent review. Legacy v2
  is returned because its tracked test consumer was not migrated. OLE owner
  source review passes. ODG/ODF/ODP require concrete final-C2a binders, and
  OPC preparation is returned for an outdated runner and missing proof
  checks. A rewritten three-lane review receipt is disclosed as unrecovered;
  its readiness claim is withdrawn. Root rechecks 75 heads/62 worktrees;
  required controls/baselines and the ODF leaf remain in use. C/P stay frozen,
  full qualification remains open, and nothing is pushed.

  Legacy v3 migrates the omitted tracked consumer and passes independent
  source review (165e), including all 18 existing tests. Root composes the
  reviewed protocol patches in an external source copy, reproduces a signed
  test-header conflict and commits the exact corrected patch; independent
  review proves equivalent final test bytes. Actual bindings/archives remain
  pending. Sheet v4 is returned for a stale XLSF3 harness pin, and OPC v3 for
  a real report producer/consumer mismatch and missing binary linkage. The
  ODF reuse leaf b3ad6000d7 passes source admission; ODG/ODP preparations are
  returned for incomplete build-to-binding transitions. The
  same gate continues facade chunks with 33 checks complete; C/P stay frozen.
  All 25 actual qualification rows remain required. Nothing is pushed.

  OPC v5 then passes independent source review with 18 tests, including the
  complete public prefix-to-policy flow, failed builds, binary drift and
  terminal status checks. Nineteen of 25 rows now have reviewed source
  preparations; every actual native qualification remains pending.

  A coverage audit restores three omitted qualification groups (165f): common
  flat ODF parts, the original ODT set/flat-open performance residuals, and
  four W6 ODS source-backed routes. The corrected inventory has 28 required
  rows; ODP tail append is already represented. Sheet v6 migrates its runner
  and tracked test together and passes independent source review, bringing
  source-ready preparations to 23. Root composes the reviewed legacy, sheet
  and ODF reuse patches into a 26-file external preview; actual bindings and
  integrated tests remain pending. ODG/ODP v3 is committed on the active leaf
  and under review. The same full65 gate continues facade tests with 33 checks
  verified. Root reclaims 3.86 GB of retired crypto debug caches, preserves
  their evidence, and confirms 75 heads/62 worktrees with no further unused
  unified leaf. All 28 actual qualifications remain open; nothing is pushed.

  Independent public-flow review returns ODG/ODP v3 (165g): the ODG fixture
  is inside a source root forbidden by final admission, and ODP's build
  producer records the shared helper's 64 MiB output cap instead of its
  required 4 MiB cap. Forty-five focused tests did not cover these composed
  failures. The author repairs both; flat-parts preparation enters independent
  review and original W6 preparation continues. Root commits a read-only
  final-gate receipt verifier, checks the existing 33 rows and exercises ten
  negative controls. The same gate remains active; C/P stay frozen and all
  28 final qualifications remain open. Nothing is pushed.

  Flat-parts v1 is returned for unbound candidate-lock admission, unsealed
  materialized probe inputs, unbounded child output and detached reports;
  its allocation-only probe also leaves native timing unprepared. W6 v1 is
  returned for incompatible historical lock treatment, a deterministic
  candidate-overlay assertion failure and incomplete bounded build/run proof.
  Root corrects the new receipt verifier's rejection of valid zero-periodic-
  sample rows while preserving initial-size and error checks. All prior
  packets and review receipts remain preserved; qualification stays open.

  The corrected receipt verifier passes independent review. The facade suite
  then finishes all 79 chunks: root verifies every chunk log and 34 completed
  top-level gate checks. Facade logs report 209 passing tests, including 19
  library tests and 20 doctests. The same gate proceeds to core tests; no
  full65 or native acceptance is claimed.

  The core suite passes six chunks and 332 tests (165h), including 253 library
  tests and 59 doctests. Root verifies 35 completed gate checks while the same
  run continues. ODG/ODP v4 repairs pass their producer checks, but ODG is
  returned because its sealed fixture path fails the later runtime validator;
  its handoff test stopped at builder validation. V5 must preserve the fixed
  reviewed harness fixture and test runtime admission. Original flat/W6
  repairs continue, with a retained e4a lock located for dependency review.
  C/P remain frozen and all 28 final qualifications remain open. No push.

  ODG/ODP v5 then passes independent source review through runtime admission,
  including the fixed tracked fixture and real producer-receipt interfaces.
  Source-ready preparations reach 25 of 28; integrated-P rebinding and all
  actual qualification remain pending. A one-digit review-hash transcription
  error is corrected in a separate receipt, preserving the original. Root's
  expanded 32-file consumer preview applies cleanly. W6 v2 is returned because
  its guarded producer is still an unimplemented placeholder; v3 must supply
  it. Flat-parts v2 has entered independent review. Nothing is pushed.


  The gate then terminates at leaf Clippy: 35 checks pass, seven DrawingML
  test lints fail one check, and 29 checks remain unrun (165h). An isolated
  leaf commits the seven fixes as 472a89a3fa; production code is unchanged.
  Root starts a bounded, isolated Clippy diagnostic with `--keep-going`.
  The failed full65 receipt is preserved and no successor acceptance is
  claimed. Composition v3 is returned for stale pristine-input paths; v4
  independently reproduces all 32 outputs from 23 sealed pristine inputs.
  W6 v3 and flat-parts v2 are returned for executable producer/admission
  defects; both are in repair. C/P remain frozen, 25 of 28 preparations are
  source-ready, and all actual qualification remains pending. No push.


  The isolated Clippy diagnostic passes after all seven test-only fixes
  (165i). Root verifies its logs, unchanged source/locks and resource records,
  then merges Q472a into trunk. A separate remaining-29 diagnostic is being
  prepared; final full65 acceptance is still required. Root starts bounded
  ordinary/allocator compilation of the ODF raw example after materializing
  its tracked dependencies. W6 v4 enters independent review, and flat-parts
  v3 remains in repair. C/P stay frozen; no native acceptance or push.


  ODF raw compilation exposes a missing lock migration (165j). Root verifies
  that its repair adds only seven local ODF packages/edges; leaf16744934d6
  updates all live lock consumers and the optimized build profile. The next
  locked release check reaches the probe and finds three Rust errors; the
  allocator check is unrun. A candidate consumer-lock overlay also remains
  necessary. ODF raw is returned, leaving 24/28 source preparations ready.
  Root repairs and tests the remaining-29 diagnostic's actual runner-source
  handoff; its reviewed v2 is ready but not launched. W6 v4 and flat-parts v3
  are returned for concrete public-flow defects, with successors in repair.
  All final qualification remains required; nothing is pushed.

  The ODF tooling leaf fixes its consumer-lock overlay, actual ZIP API calls,
  buffer borrow and generic closure dispatch (165k). Both current release
  compiler variants pass at 340191f1ef; all 82 affected Python tests pass
  after materializing a tracked fixture omitted by the sparse checkout.
  Independent review passes and root merges the leaf without conflicts.
  Historical compilation and native qualification remain pending. The
  remaining-29 driver fixes OfficeCLI symlink admission, passes 17 controls
  and independent review, then starts its bounded diagnostic. Flat-parts v5
  passes independent source review; W6 v5 repair continues. The branch audit
  confirms 76 heads/63 worktrees and 70 unchanged excluded refs. Every retained
  unified leaf branch is now merged; required validation inputs remain.
  The unused ODF design worktree/ref is removed, bringing cumulative cleanup
  to 278 worktrees/440 refs. C/P stay frozen; final qualification remains open.
  Nothing is pushed.

  The remaining-29 diagnostic passes ODS Clippy, the full ODS suite and
  workspace Clippy, but finds two stale DOCX metadata assertions (165l).
  An isolated leaf corrects them and strengthens serialized empty-core and
  independent replay coverage; focused Rust validation remains pending.
  Final-source admission review also finds that ODF preparation cannot admit
  the probe and lock now committed on trunk. Its preparer, runner, manifest
  and tests migrate together on a new leaf; review returns incomplete P trust,
  adapter and report/input-drift checks. Protocol previews
  separate P1/P2, with successive guards returned for actual producer and
  provenance defects; v9 repair continues. W6's real source preparation passes,
  but remaining path/child-invocation defects require v6, now under review. The
  exact 70d ODF baseline is prepared for later lock resolution/compilation.
  Root removes 8.35 GB of obsolete gate caches and preserves inputs before
  retiring failed W6 source copies. All final qualification remains required;
  C/P stay frozen and nothing is pushed.

  ODF final-source admission is merged at leaf 172c803ebd (165m). It verifies
  committed candidate inputs without writes and checks the actual C → P1 → P2
  identity chain. Independent review findings are repaired; root verifies the
  final seal-equality correction and all 16 affected tests pass. Protocol
  composition v10 passes source review and 20 controls; v11 corrects only a
  stale test count and dependent receipt hashes. Actual P installation and
  rebinding remain pending. The branch census confirms 78 heads/66 worktrees
  and 70 unchanged excluded heads; DOCX is the remaining unmerged repair.
  Root retires the merged ODF admission worktree/ref, leaving 77 heads and
  65 worktrees (279 worktrees/441 refs retired cumulatively). Cleanup removes
  all four failed W6 source copies (about 111.66 GB apparent), preserving
  logs and recovery inputs. After 157 passing chunks, the workspace suite
  reproduces the same two DOCX assertions and proceeds to isolated wsdoc.
  Focused DOCX validation passes source preflight and is queued to take the
  Cargo slot only after the complete diagnostic becomes terminal. W6 repair
  and all final qualification remain required. Nothing is pushed.
