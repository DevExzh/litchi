# Performance program phase report

Date: 2026-08-12
Branch: `feat/office-format-completeness`
Production base for the latest measured tranche: `59d1a17d85df086e90d7b0fd8cfb18267db106ed`

This report summarizes the measured implementation tranches to date. It is not a
claim that the end-to-end performance program or CRUD scenario matrix is
complete. The reproducible environment, original substrate baseline, corpus
definitions, commands, and profiler limitations are in
[`BASELINE.md`](BASELINE.md); raw reports are under [`results/`](results/).

## Current stable tranche

The original stage-1 results below remain historical evidence. The current
harness contains **128 selectable cases**: 36 default cases and 198 default
records, plus six opt-in simulated-range cases, two opt-in scaling cases, one
opt-in XLSX commit/read attribution case, four opt-in opaque-heavy common OLE2
stage/control cases, one opt-in source-backed OPC one-Part publication case,
one opt-in source-backed DOCX semantic publication case, one opt-in
source-backed media-rich PPTX semantic publication case, one opt-in media-rich ODT
paragraph-publication case, six opt-in matched XLSX
calculation-metadata/page-break/page-margin publication cases, 16 opt-in DOCX/PPTX
semantic cases, seven opt-in RTF semantic case names across four
capability-bounded variants (25 tiny / 44 tiny-plus-large rows), 23
shape-selected ODT/ODS/ODP semantic cases, three fixed media-rich ODF cases,
and 21 opt-in native DOC/XLS/PPT semantic cases. It
is still not broad program or CRUD coverage.

| Change | Current evidence | Scope / limitation |
|---|---|---|
| XLSX row-start index | ABBA p50 geomean **-80.499%**, mean geomean **-79.962%**; full scan **+0.03%** mean; first cell **-1.31%** mean | Heap allocations **+17**, RSS **+0.25%**; narrow-range query only |
| Targeted OPC raw publication | Four-cell ABBA p50 geomean **-84.98%**; few-large/incompressible **-71.70%**; matched cycles **-69.21%** | Initial peak heap **+37.18%**, one-shot RSS **+22.26%** from retained source/provenance and a changed-payload copy; the copy is removed by the shared-payload follow-up below |
| Positional CFB/ZIP and explicit execution | Large-task p50 scaling at 12 CPUs: OPC **4.52x**, CFB **5.93x**; no hidden global Rayon | Many-small tasks regress at high worker counts; default/legacy paths remain serial |
| Source-backed OPC and DOCX/XLSX/PPTX facades | EOCD structural-open source bytes **-73.6% to -98.5%**; ordinary payload overlap zero | No latency claim: later EntryId/cache-diagnostic changes confound comparison and some cells exceed 5% variance |
| Source-backed PPTX selected-slide publication | Media-rich one-edit/save p50 **-97.12%**; atomic eight-shape batch p50/mean **-97.45%**, allocation calls **-39.80%**; materializations **229 -> 2**; byte-identical output | One operation in one existing slide only, bounded to 256 unique nonoverlapping shape-text replacements; MCE rewrites, topology changes and changed signed packages refuse before output |
| Source-backed XLSX calculation-metadata publication | Media-rich one-edit/save p50 **-99.2519%** (133.67x), mean **-99.2507%**; instructions **-77.78%**; materializations **12 -> 1**; byte-identical output | Existing `xl/workbook.xml` calculation properties/features only; cells, formulas, cached results, chains, relationships and topology remain outside the capability |
| Source-backed XLSX page-break publication | Media-rich one-edit/save p50 **-97.86%** (46.65x), mean **-97.86%**; materializations **12 -> 2**; byte-identical output | One existing normal worksheet's page-break collections only; cells, formulas, styles, relationships, topology, and changed signed sources remain outside the capability |
| Source-backed XLSX page-margin publication | Media-rich one-edit/save p50 **-97.93%** (48.26x), mean **-97.93%**; materializations **12 -> 2**; byte-identical output | One existing normal worksheet's direct six-value page-margin set only; cells, formulas, styles, relationships, topology, and changed signed sources remain outside the capability |
| Deterministic range simulation | XLSX listing has zero timed requests; selected reads have zero unselected-sheet overlap; full physical size distributions recorded | Synthetic latency model, not a cold filesystem or ambient network |
| DOCX/PPTX semantic selectors and edits | DOCX one paragraph **-4.72%** p50; PPTX 1% edit/save **-9.37%** p50 and mean; PPTX one-edit guardrail +0.28% p50 (neutral) | Generated text corpora; complete transaction capture dominates one edit; no ODF/iWork implication |
| Coalesced DOCX paragraph edits | Large 100-edit/save p50 **-94.99% (19.97x)** and mean **-95.02%**; medium two-edit/save p50 **-12.98%**; scalar one-edit guardrail neutral | Direct-body, strictly ordered paragraph text replacement; generated corpus; scalar API remains separate |
| ODF semantic baselines and ODS snapshot reuse | Medium/large ODS no-op edit-save p50 **-7.45% / -11.78%**; one-cell edit-save **-3.57% / -2.06%** | Generated ODT/ODS/ODP baseline corpora; focused ODP/ODT publication follow-ups are listed below |
| RTF semantic baseline and text paths | Medium/large full-text p50 **-38.39% / -27.08%**; one-edit/save **-33.40% / -25.79%** | Generated native RTF text corpus; open guard +0.96% / +3.41%; formatting/media/security matrices remain missing |
| RTF retained story length | Large paragraph-list p50 **-15.04%**, mean **-13.71%**; middle-paragraph p50 **-27.19%**, mean **-25.23%** | Already-open generated 10,000-block story queries only; exact parser-derived length, all allocation/peak-heap/RSS metrics flat, and open/full-text/save/no-op guards remain within 5% |
| RTF sparse paragraph selection | Large middle-paragraph p50 **-47.87%**, mean **-47.95%**, p95 **-49.42%** | Explicit `Paragraphs::nth` only; remains linear and allocation-free, constructs the selected view once, preserves iterator state/formatting, and leaves open/list/full-text/save/edit guards within policy |
| ODT shared transaction bytes | Medium/large no-op edit-save p50 **-27.05% / -18.51%**; exactly two allocations and one archive copy removed per snapshot | Existing-document snapshot handoff only; changed edit/save and open guardrails remain within 3%; changed publication still rewrites the package |
| ODT consuming full-text blocks | Repeated large full-text p50 **-3.25%**, mean **-4.81%**; allocation calls **-15.48%**, temporary allocations **-45.52%** | Private full-text mode only; structured queries remain near neutral; unchanged open +3.94% p50/+4.17% mean and +10.95% p99 disclosed |
| ODT indexed paragraph selector | Large middle-paragraph p50 **-48.56%**, mean **-48.33%**; allocation calls **-27.05%**; peak heap **-24.74%**; RSS **-10.93%** | Complete XML/limit validation remains; retains one paragraph, excludes headings from the index, and leaves the established list path neutral |
| ODT content-only unchanged-media publication | Media-rich paragraph edit/save p50 **-95.58%**, mean **-95.63%**, p95 **-95.43%**; allocation calls **-6.71%**; peak heap flat and RSS **-0.59%** | Exactly one paragraph in a fixed 16 MiB-media package; structural/mixed operations and regenerated content over the common 16 MiB limit retain the established rebuild |
| ODT direct snapshot byte sharing | Media-rich direct paragraph edit/save p50 **-75.84%**, mean **-73.84%**, p95 **-75.41%**; peak heap/RSS flat | Removes two 16 MiB archive copies from direct snapshot validation/rehydration; complete XML parsing, publication, reopen/readback, patch and inverse remain |
| ODT compact-audit package sharing | Media-rich paragraph edit/save p50 **-30.44%**, mean **-31.36%**, p95 **-32.41%**; allocations **-0.57%**, peak heap/RSS flat | Removes three archive-sized audit copies (50.36 MB/operation); compact validation, final materialization and readback remain; exact no-op +39 ns p50 is disclosed |
| ODT envelope-classification package sharing | Media-rich paragraph edit/save p50 **-11.40%**, mean **-11.95%**, p95 **-12.19%**; two allocations/commit removed, peak heap/RSS flat | Removes one 16.79 MB envelope copy; archive/manifest and signed/encrypted classification remain; large exact no-op +152 ns p50 is disclosed |
| ODT final changed-result byte handoff | Media-rich paragraph edit/save p50 **-22.74%**, mean **-22.56%**, p95 **-21.48%**; final 16.79 MB copy removed; allocation calls **-3.46%** | Snapshot remains byte-only and one independent final reopen remains; medium one-paragraph +2.77% p50/+1.29% mean is within the 3% gate; peak heap/RSS flat |
| Coalesced ODT paragraph publication | Large 100-edit/save p50 **-98.28% (58.05x)**, mean **-98.27%**; medium two-edit/save p50 **-27.62%**; allocation calls **-96.13%** | Consecutive plain-text replacements only; ordinary durable operations, ordered duplicate semantics, atomic refusal, compact audit, full reopen and scalar path remain |
| Native DOC/XLS/PPT semantic baseline | Large one-edit/save p50: XLS **1.722 ms**, DOC **1.416 ms**, PPT **0.357 ms**; large XLS open **1.383 ms** | Generated writer corpora; accepted XLS and DOC follow-ups are listed below |
| Native XLS validated-editor reuse | Large one-cell edit/save p50 **-7.72%**, mean **-7.90%** | Final exact owner parse, public Workbook reopen and typed readback remain; peak heap/RSS flat |
| Native XLS fixed-width numeric inventory carry-forward | Large one-cell edit/save p50 **-7.83%**, mean **-7.37%**, p95 **-7.20%** | Exact byte-range proof plus complete public Workbook validation/readback remain; peak heap -5.54%, RSS flat; all nonnumeric/structural/resource edits retain full parse |
| Rejected XLS terminal-render handoff | Tiny changed save p50 **-7.55%**; large changed save **-0.39%** (neutral) | Fully reverted: repeated large exact no-op p50 **+22.00%**, mean **+16.69%** |
| Common OLE2 publication stages and rejected handoffs | Current open/publication/finish/end-to-end p50: **1.382 / 7.979 / 5.473 / 26.086 ms**; inline recapture prototype end-to-end **-2.61%** p50 | Stages are non-additive; shared-payload, validated-render and inline recapture prototypes are all fully reverted |
| Native DOC batched stream publication | Large one-paragraph edit/save p50 **-10.52%**, mean **-10.48%** | Ordinary two-stream replacement only; final strict revision and independent document reopens remain |
| Native DOC PieceTable physical index | Large open p50 **-55.91%**, mean **-55.78%**; changed edit/save p50 **-31.08%** | Private FC-ordered/prefix-max index only; exact scalar mapping, full FKP validation and strict/public reopens remain; peak heap/RSS flat |
| Native DOC paragraph-style baseline cache | Large open p50 **-11.44%**, mean **-11.87%**; changed edit/save p50 **-4.01%** | One private resolved baseline only; direct PAPX, piece modifiers, direct style switches and complete readbacks remain; allocation calls -18.61%, peak heap/RSS flat |
| Native DOC CHPX range index | Large paragraph-list p50 **-21.07%**, mean **-20.93%**, p95 **-20.00%** | Private monotonic slice query only; exact run identity/order, property cascading and complete readbacks remain; allocations and peak heap/RSS flat |
| Native PPT root snapshot CFB reuse | Repeated large root open p50 **-8.78%**, mean **-10.58%**; allocation calls **-5.01%** | Reuses only the validated CFB index; independent stream/current-user/live-document, slide-order, review-history and public-reader checks remain |
| Native PPT text-edit resolver reuse | Direct large edit/save p50 **-14.12%**, mean **-15.39%**; allocation calls **-3.53%** | Reuses the full editor preflight for persisted-record resolution; exact error precedence, fresh commit editor and complete readback remain; minor-fault increase disclosed |
| Native PPT root text-publication adoption | Large root one-shape edit/save p50 **-18.59%**, mean **-17.83%**, p95 **-16.58%**; allocation calls **-6.54%** | Exact source and selected-slide persist identity gate a private output-Arc handoff; custom limits and structural edits retain complete root reopen; peak heap/RSS flat |
| Bounded XLSX validated-store handoff | Medium one-cell commit + first read p50 **-23.23%**, mean **-23.15%**; allocation calls **-21.01%** | At most 4,096 cells / 1 MiB XML with exact byte and lineage identity; peak heap +4.29%; unrestricted dense-wide candidate rejected at +8.99% peak heap |
| Rejected direct XLSX action-plan flattening | Best formal p50 **-1.61%**; dense commit **-0.27%** p50 with mean interval crossing zero | Fully reverted; process allocation calls -0.0623%, peak heap flat, medium commit p99 +4.33% |
| XLSX no-extension worksheet scan | Medium commit/save p50 **-19.31% to -20.74%**; cold reads about **-35%**; dense 1% commit p50 **-19.62%** | `dyDescent`-free success path only; rejected inputs rerun the original collector for error precedence; allocation calls -25.24%, peak heap flat |
| ODS row-local publication | Large/medium one-cell edit-save p50 **-9.54% / -7.22%**; allocation calls **-5.85%**, peak heap **-27.18%** | Same-topology modeled rows only; structural edits fall back and touched opaque rows refuse |
| ODS unchanged-media publication | Media-rich one-cell edit/save p50 **-4.73%**, mean **-5.73%**, p95 **-7.65%**; peak heap **-8.78%** | Compact `content.xml` replacements in ordinary unsigned/unencrypted ZIPs; every unproved layout/member retains logical rebuild or comparison fallback |
| ODS shared durable-patch blobs | Media-rich one-cell edit/save p50 **-8.80%**, mean **-9.07%**, p95 **-13.85%**; 33.58 MB copy site removed; peak heap **-1.92%** | Shares only already retained immutable source/target package bytes with the forward/reverse semantic bundles; patch wire, limits, final reopen and media verification remain |
| ODS row-splice raw publication | Media-rich one-cell edit/save p50 **-74.16%**, mean **-74.17%**, p95 **-74.11%**; instructions **-69.04%**; peak heap/RSS flat | Same-topology compact row replacements only; exact checked range provenance reaches raw ZIP emission, while structural, signed/encrypted and unsupported layouts retain established fallback/policy |
| ODS shared worksheet archive handoff | Media-rich one-cell edit/save p50 **-21.32%**, mean **-21.30%**, p95 **-21.15%**; peak heap **-22.03%**, RSS **-20.57%** | Private nested worksheet snapshot/package/unified staging only; exact source lineage, failure rollback, durable patches and final readback remain |
| ODP content-only unchanged-media publication | Media-rich text-box edit/save p50 **-94.44%**, mean **-94.43%**, p95 **-94.29%**; allocation calls **+0.52%**; peak heap/RSS flat | Source-backed content-only operations reuse accepted checked-splice/raw-copy publication; resource additions and unsupported/security-sensitive layouts retain logical rebuild |
| ODS exact no-op handoff | Large exact-no-op p50 **-23.26%**, mean **-23.21%**; instructions **-10.54%**; peak heap flat | Exact no-op only; changed commits retain complete audit, preservation and readback paths; read-only link-layout trigger disclosed |
| ODP indexed slide selector | Large middle-slide p50 **-4.09%**, mean **-4.20%**, p95 **-5.18%**; allocation calls **-3.86%**; peak heap/RSS flat | Full style/content EOF validation remains; tiny is neutral, medium p50 -1.55%, and unchanged list/save guards remain within thresholds |
| ODP snapshot slide-projection reuse | Large exact-no-op edit/save p50 **-59.96% (2.50x)**, mean **-59.92%**; large changed edit/save p50 **-20.78%**; allocation calls **-20.13%** | Reuses only the snapshot-validated slide projection for detached staging; package/security reopen, auxiliary parsing, raw page coverage, complete publication/readback and peak heap/RSS remain |
| ODP final slide-snapshot handoff | Large one-slide edit/save p50 **-32.35%**, mean **-32.92%**, p95 **-35.95%**; allocation calls **-16.71%** | Exact slide-only commits move the already parsed candidate projection only after the independent final package/audit/media pipeline; compound domains retain ordinary final parsing; peak heap/RSS flat |
| ODS adaptive cell locator | Large public cell sweep p50 **-81.74%**, mean **-80.72%**; full cell text p50 **-52.65%** | Builds lazily at 64 calls, requests 3,216 bytes on the dense corpus and is capped at 4 MiB; peak heap/RSS flat |
| RTF parser-state specialization | Large open p50 **-20.09%**; large/medium one-edit-save **-11.54% / -14.16%**; cycles **-10.50%** | Ordinary body text only; insertion/deletion metadata retains the full state; allocation count, peak heap and RSS flat |
| RTF ASCII transport batching | Large open p50 **-26.67%**; large/medium one-edit-save **-6.26% / -10.07%**; instructions **-18.40%** | ASCII source tokens only; byte-valued non-ASCII and invalid-Unicode fallback unchanged; allocation count, peak heap and RSS flat |
| RTF byte delimiter scanning | Large open p50 **-17.23%**, mean **-17.99%**; one-edit/save p50 **-14.65%**, mean **-14.84%**; instructions **-21.27%** | Ordinary-text lexer only; plain/CP-1252/LZFu opens improve; prepared LZFu no-op segment +0.290 us/+6.41% p50 is disclosed while complete open improves 19.39%; peak heap/RSS flat |
| RTF retained body source span | Large one-edit/save p50 **-10.72%**, mean **-10.11%**, p95 **-8.76%**; instructions **-10.64%** | Direct uncompressed ASCII ordinary bodies only; cached range is proven during full parser preflight, while ambiguous/binary/non-ASCII/LZFu inputs keep the established locator/refusal and candidate parse/readback |
| RTF bounded body-block reservation | Large open p50 **-21.17%**, mean **-21.00%**, p95 **-21.04%**; one-edit/save p50 **-1.46%**; peak heap **-29.73%** | Sources >=64 KiB only; exact root-text count, token/source/16 MiB caps, lazy fallible allocation, and table/deletion fallback retain semantic behavior; medium plain/CP-1252 +0.49%/+2.84% p50 disclosed |
| Rejected RTF decoded-body ownership | Broad raw CP-1252 open **-3.08% p50 / -3.28% mean**; allocation calls **-20.15%** | Fully reverted: plain large open **+25.53% p50 / +22.45% mean**; owned-only variants were compiler-layout sensitive at -1.41% and +1.02% p50 |
| OPC shared changed-Part payload | Few-large compressible targeted save **-20.73%** p50 / **-18.49%** mean; cache misses **-31.12%** | Removes one 4.19 MiB handoff copy; peak heap -3.42%, uninstrumented RSS +0.22% (flat); the remaining local-span copy is removed by the follow-up below |
| ZIP generated local-span move | Few-large compressible/incompressible targeted save **-4.09% / -2.70%** p50; means **-4.08% / -2.25%** | Removes the separate 4.20 MiB post-validation local-span copy; peak heap -3.20%, uninstrumented RSS -0.10% (flat); required compressor/archive buffer remains |
| Source-backed OPC one-Part publication | Fixed four-Part save p50 **-73.12%**, mean **-73.58%**; semantic materializations **4 -> 1**; instructions **-65.42%** | Low-level consuming same-topology replacement only; raw-copies all unselected ZIP members; signed real changes and unsupported layouts refuse before output; complete physical input/output bytes remain |
| Source-backed DOCX semantic publication | Fixed media-rich one-edit/save p50 **-97.43%**, mean **-97.41%**, p95 **-97.27%**; materializations **17 -> 1**; instructions **-74.91%** | Exact raw main-document transactions only; MCE rewrites, dependency transfers and signed real changes refuse; physical archive input/output remains and eager DOCX guard p50 is +0.25% |

Raw evidence: [`XLSX before A`](results/abba-xlsx-range-before-a.json),
[`after A`](results/abba-xlsx-range-after-a.json),
[`before B`](results/abba-xlsx-range-before-b.json),
[`after B`](results/abba-xlsx-range-after-b.json); [`EOCD before A`](results/abba-eocd-before-a.json),
[`after A`](results/abba-eocd-after-a.json), [`before B`](results/abba-eocd-before-b.json),
[`after B`](results/abba-eocd-after-b.json); and
[`source-versus-eager`](results/stage3-source-vs-eager-many-small.json). The
committed positional XLSX record is
[`xlsx-source-positional.json`](results/xlsx-source-positional.json): p50 open
is 33.881 us/56.493 us/139.897 us (tiny/medium/dense), listing after open has
zero timed source reads, and first/range reads physically overlap only the
selected worksheet member (zero unselected worksheet read calls). These are
physical-overlap counts, not materialization counts.

The semantic raw evidence is [`before A`](results/abba-semantic-before-a.json),
[`after A`](results/abba-semantic-after-a.json),
[`after B`](results/abba-semantic-after-b.json), and
[`before B`](results/abba-semantic-before-b.json). The dedicated 60-sample
one-edit guardrail is linked from
[`change 0010`](changes/0010-docx-pptx-semantic-queries-and-edits.md).

The ODF/ODS raw evidence is [`before A`](results/abba-odf-before-a.json),
[`after A`](results/abba-odf-after-a.json),
[`after B`](results/abba-odf-after-b.json), and
[`before B`](results/abba-odf-before-b.json). Pooled statistics and the
allocation/RSS guardrails are in
[`change 0011`](changes/0011-odf-semantic-baseline-and-ods-snapshot.md).

The coalesced-DOCX raw evidence is
[`before A`](results/abba-docx-batch-before-a.json),
[`after A`](results/abba-docx-batch-after-a.json),
[`after B`](results/abba-docx-batch-after-b.json), and
[`before B`](results/abba-docx-batch-before-b.json). Pooled statistics and the
allocation/RSS guardrails are in
[`change 0012`](changes/0012-docx-coalesced-paragraph-edits.md).
That record also links the dedicated four-leg large-corpus scalar one-edit
guardrail (p50 -1.28%, mean +0.79% with overlapping intervals), which is
treated as neutral.

The RTF raw evidence is
[`text before A`](results/abba-rtf-text-before-a.json),
[`text after A`](results/abba-rtf-text-after-a.json),
[`text after B`](results/abba-rtf-text-after-b.json), and
[`text before B`](results/abba-rtf-text-before-b.json). The independent open
guard, complete seven-case matrix, allocation/RSS evidence, and rejected first
candidate are in
[`change 0013`](changes/0013-rtf-semantic-baseline-and-text-paths.md).

The ODT shared-snapshot raw evidence is
[`before A`](results/abba-odt-shared-before-a.json),
[`after A`](results/abba-odt-shared-after-a.json),
[`after B`](results/abba-odt-shared-after-b.json), and
[`before B`](results/abba-odt-shared-before-b.json). Allocation attribution,
pooled statistics, open/changed-publication guardrails, and RSS evidence are in
[`change 0014`](changes/0014-odt-shared-snapshot-bytes.md).

The ODT media-publication raw evidence is
[`before A`](results/abba-odt-media-paragraph-before-a.json),
[`after A`](results/abba-odt-media-paragraph-after-a.json),
[`after B`](results/abba-odt-media-paragraph-after-b.json), and
[`before B`](results/abba-odt-media-paragraph-before-b.json). The ordinary ODT
guard ABBA, allocation/RSS/counter profiles, binary identity and common-limit
fallback proof are indexed in
[`change 0035`](changes/0035-odt-content-only-paragraph-publication.md).

The native OLE2 semantic baseline is
[`ole2-semantic-baseline-a57506d23-2026-08-11.json`](results/ole2-semantic-baseline-a57506d23-2026-08-11.json).
Its complete latency table, Heaptrack/RSS evidence, hardware counters, and
ranked next target are in
[`change 0015`](changes/0015-native-ole2-semantic-baseline.md).

The accepted native XLS follow-up reuses its validated object editor; its
primary raw reports are
[`before A`](results/abba-xls-commit-reuse-one-edit-before-a.json) and
[`after A`](results/abba-xls-commit-reuse-one-edit-after-a.json), with pooled
statistics and all four legs in
[`change 0016`](changes/0016-xls-commit-editor-reuse.md). The later fixed-width
numeric follow-up carries the private BIFF inventory only after exact
field-range certification and keeps the complete public Workbook validation
boundary. Its record and pooled evidence are
[`change 0059`](changes/0059-xls-fixed-numeric-inventory-carry.md) and the
[`primary summary`](results/xls-inventory-carry-primary-summary.json). The DOC follow-up
batches ordinary stream replacement; its primary raw reports are
[`before A`](results/abba-doc-stream-batch-one-edit-before-a.json) and
[`after A`](results/abba-doc-stream-batch-one-edit-after-a.json), with the
complete record in
[`change 0017`](changes/0017-doc-batched-stream-publication.md).

The later XLS terminal-render handoff was measured and fully reverted. Its
large changed-save p50 improved only 0.39%, while four repeated exact-no-op
cycles regressed 22.00% p50 and 16.69% mean. The profile, equality prototype,
allocation evidence and rejection gate are in
[`change 0028`](changes/0028-xls-terminal-render-handoff-rejected.md).

The native PPT root-snapshot evidence is retained as four short ABBA cycles
under `results/abba-ppt-slide-order-root-repeat-*.json`. Reader/edit guards,
allocation attribution, RSS, counters, the disclosed initial selected-shape
tail and its neutral repeat are summarized in
[`change 0024`](changes/0024-ppt-slide-order-open-reuse.md).

The later native PPT root text-publication adoption is summarized in
[`change 0062`](changes/0062-ppt-root-text-publication-adoption.md), with exact
pooled latency, guard, allocation, RSS, and counter values in its linked JSON.

The bounded XLSX commit/read evidence is retained under
`results/abba-xlsx-store-handoff-*.json`; the exact identity gates, primary
latency, allocation/RSS/counter attribution and rejected unrestricted
dense-wide prototype are summarized in
[`change 0025`](changes/0025-xlsx-validated-store-handoff.md).

The direct XLSX writer-regrouping prototype was also fully reverted. Its
medium and dense-wide 1% commit/save ABBA reports are under
`results/abba-xlsx-action-plan-*.json`; matched allocation evidence and the
rejection rationale are in
[`change 0030`](changes/0030-xlsx-action-plan-flattening-rejected.md).

The accepted XLSX no-extension scan evidence is under
`results/abba-xlsx-x14ac-*.json`. Medium and dense-wide latency, read/no-op
guards, allocation/RSS/counter attribution and malformed-input precedence are
summarized in
[`change 0032`](changes/0032-xlsx-no-extension-scan.md).

The new common OLE2 publication evidence is under
`results/abba-ole-common-*.json`. It retains the deterministic opaque-heavy
case, but both measured production handoffs were reverted: direct shared
writer payloads regressed 32.02% p50, while retaining the validated render
improved the target 34.06% but regressed DOC open 21.64%. The full rationale
and DOC/XLS guards are in
[`change 0033`](changes/0033-ole-common-publication-handoffs-rejected.md).

The common OLE2 stage/recapture reports are
[`before A`](results/abba-ole-recapture-before-a.json),
[`after A`](results/abba-ole-recapture-after-a.json),
[`after B`](results/abba-ole-recapture-after-b.json), and
[`before B`](results/abba-ole-recapture-before-b.json). The stage profile,
non-additivity finding and fully reverted inline allocation-reuse prototype are
documented in
[`change 0036`](changes/0036-ole-common-stage-attribution.md).

The ODS row-local publication evidence is
[`before A`](results/abba-ods-row-splice-one-edit-before-a.json),
[`after A`](results/abba-ods-row-splice-one-edit-after-a.json),
[`after B`](results/abba-ods-row-splice-one-edit-after-b.json), and
[`before B`](results/abba-ods-row-splice-one-edit-before-b.json). Medium,
guardrail, allocation, RSS and hardware-counter evidence is summarized in
[`change 0018`](changes/0018-ods-row-local-publication.md).

The adaptive ODS cell-locator ABBA evidence starts at
[`before A`](results/abba-ods-cell-locator-before-a.json) and
[`after A`](results/abba-ods-cell-locator-after-a.json); the complete profile,
guard, memory and counter record is
[`change 0027`](changes/0027-ods-adaptive-cell-locator.md).

The ODS unchanged-media publication evidence is
[`before A`](results/abba-ods-media-preservation-before-a.json),
[`after A`](results/abba-ods-media-preservation-after-a.json),
[`after B`](results/abba-ods-media-preservation-after-b.json), and
[`before B`](results/abba-ods-media-preservation-before-b.json). Raw-member
proofs, the no-media guard, fallback semantics, memory and counter attribution
are summarized in
[`change 0031`](changes/0031-ods-unchanged-media-preservation.md).

The ODS durable-patch ownership evidence starts with the balanced primary
[`before`](results/ods-shared-patch-blobs-primary-forward-1-before.json) and
[`after`](results/ods-shared-patch-blobs-primary-forward-1-after.json) legs.
All four primary pairs, medium/large guards, profiles, counters, memory, RSS,
wire-identity checks and binary provenance are indexed in
[`change 0054`](changes/0054-ods-shared-durable-patch-blobs.md).

The ODS row-splice raw-publication evidence retains all 300 samples per state
in the
[`primary summary`](results/ods-row-splice-raw-publication-primary-summary.json).
Tiny, medium and large ordinary CRUD distributions, matched profiles,
counters, Heaptrack, GNU Time and exact binary provenance are indexed in
[`change 0057`](changes/0057-ods-row-splice-raw-publication.md).

The ODS shared worksheet-ownership evidence pools 2,000 ABBA/reverse-BAAB
samples per state in the
[`summary`](results/ods-worksheet-shared-ownership-summary.json). Large
ordinary guards, matched Heaptrack/GNU Time/perf evidence, exact binary
provenance, and the rejected 4.01% intermediate are indexed in
[`change 0068`](changes/0068-ods-shared-worksheet-archive-handoff.md).

The ODP content-only publication evidence is
[`before A`](results/abba-odp-media-textbox-before-a.json),
[`after A`](results/abba-odp-media-textbox-after-a.json),
[`after B`](results/abba-odp-media-textbox-after-b.json), and
[`before B`](results/abba-odp-media-textbox-before-b.json). Raw-member proofs,
ordinary ODP guards, patch/inverse checks, memory, and hardware counters are
summarized in
[`change 0034`](changes/0034-odp-unchanged-media-preservation.md).

The ODP snapshot-projection evidence pools 4,000 large exact-no-op samples per
state in the
[`primary summary`](results/odp-slide-projection-primary-summary.json).
Tiny/medium scaling, large changed-edit and read/media guards are in the
[`guard summary`](results/odp-slide-projection-guard-summary.json); matched
profiles, counters, Heaptrack, GNU Time and exact binary provenance are indexed
in [`change 0060`](changes/0060-odp-snapshot-slide-projection-reuse.md).

The ODP final-snapshot evidence uses a drift-gated warmed 2,000-sample pool per
state in the [`summary`](results/odp-final-snapshot/summary.json). Tiny/medium
scaling, ineligible no-op/media guards, repeated read-only tails, matched
Heaptrack, GNU Time, counters and binary provenance are indexed in
[`change 0065`](changes/0065-odp-final-snapshot-handoff.md).

The RTF parser-state follow-up evidence is
[`before A`](results/abba-rtf-state-clone-one-edit-before-a.json),
[`after A`](results/abba-rtf-state-clone-one-edit-after-a.json),
[`after B`](results/abba-rtf-state-clone-one-edit-after-b.json), and
[`before B`](results/abba-rtf-state-clone-one-edit-before-b.json). Open/save
guardrails, profiles, hardware counters, memory results and the rejected ODS
candidate are summarized in
[`change 0019`](changes/0019-rtf-parser-state-specialization.md).

The RTF transport-batching evidence is
[`before A`](results/abba-rtf-ascii-transport-primary-before-a.json),
[`after A`](results/abba-rtf-ascii-transport-primary-after-a.json),
[`after B`](results/abba-rtf-ascii-transport-primary-after-b.json), and
[`before B`](results/abba-rtf-ascii-transport-primary-before-b.json). Medium,
save-only, profile, counter and memory guardrails plus the rejected ODT
candidate are summarized in
[`change 0020`](changes/0020-rtf-ascii-transport-batching.md).

The RTF byte-delimiter evidence is
[`before A`](results/abba-rtf-byte-delimiter-final-before-a.json),
[`after A`](results/abba-rtf-byte-delimiter-final-after-a.json),
[`after B`](results/abba-rtf-byte-delimiter-final-after-b.json), and
[`before B`](results/abba-rtf-byte-delimiter-final-before-b.json). Plain,
CP-1252 and LZFu guards, the prepared LZFu no-op disclosure, profiles,
counters, memory and complete correctness gates are summarized in
[`change 0040`](changes/0040-rtf-byte-delimiter-scanning.md).

The rejected RTF decoded-body ownership evidence includes two broad-prototype
ABBA cycles, plain/CP-1252/LZFu and prepared-operation guards, two owned-only
refinements, Heaptrack, `perf record`, and GNU Time summaries. The raw JSON
digests and full rejection rationale are in
[`change 0043`](changes/0043-rtf-decoded-body-ownership-rejected.md).

The retained RTF body-source-span evidence is
[`before A`](results/abba-rtf-body-span-before-a.json),
[`after A`](results/abba-rtf-body-span-after-a.json),
[`after B`](results/abba-rtf-body-span-after-b.json), and
[`before B`](results/abba-rtf-body-span-before-b.json). Tiny/medium scaling,
open/list/no-op guards, allocation attribution, counters, RSS, capability
smoke and artifact hashes are summarized in
[`change 0048`](changes/0048-rtf-retained-body-source-span.md).

The source-backed PPTX selected-slide publication evidence is
[`before A`](results/abba-pptx-source-edit-before-a.json),
[`after A`](results/abba-pptx-source-edit-after-a.json),
[`after B`](results/abba-pptx-source-edit-after-b.json), and
[`before B`](results/abba-pptx-source-edit-before-b.json). The eager semantic
guard, CPU/allocation/RSS attribution, exact preservation/refusal matrix and
frozen binary hashes are summarized in
[`change 0044`](changes/0044-pptx-source-backed-semantic-publication.md).

The source-backed XLSX calculation-metadata publication evidence is
[`before A`](results/abba-xlsx-calculation-metadata-edit-before-a.json),
[`after A`](results/abba-xlsx-calculation-metadata-edit-after-a.json),
[`after B`](results/abba-xlsx-calculation-metadata-edit-after-b.json), and
[`before B`](results/abba-xlsx-calculation-metadata-edit-before-b.json).
Counters, allocation/RSS attribution, exact workbook/media preservation,
refusal coverage and frozen binary/input/output hashes are summarized in
[`change 0046`](changes/0046-xlsx-source-backed-calculation-metadata-publication.md).

The coalesced ODT paragraph-publication evidence is
[`before A`](results/abba-odt-paragraph-batch-before-a.json),
[`after A`](results/abba-odt-paragraph-batch-after-a.json),
[`after B`](results/abba-odt-paragraph-batch-after-b.json), and
[`before B`](results/abba-odt-paragraph-batch-before-b.json). Scalar/no-op
guards, CPU/allocation/RSS attribution, durable replay, media preservation,
over-limit fallback and frozen binary hashes are summarized in
[`change 0045`](changes/0045-odt-coalesced-paragraph-publication.md).

The ODT compact-audit package-sharing evidence is
[`before A`](results/abba-odt-compact-audit-final-before-a.json),
[`after A`](results/abba-odt-compact-audit-final-after-a.json),
[`after B`](results/abba-odt-compact-audit-final-after-b.json), and
[`before B`](results/abba-odt-compact-audit-final-before-b.json). Ordinary
open/edit/no-op guards, the dedicated 10,000-sample/state no-op disclosure,
profiles, counters, memory, allocator policy and complete correctness gates are
summarized in
[`change 0041`](changes/0041-odt-compact-audit-package-sharing.md).

The ODT envelope-sharing evidence comprises two balanced ABBA cycles:
[`cycle 1 before A`](results/abba-odt-envelope-sharing-rerun-before-a.json),
[`after A`](results/abba-odt-envelope-sharing-rerun-after-a.json),
[`after B`](results/abba-odt-envelope-sharing-rerun-after-b.json),
[`before B`](results/abba-odt-envelope-sharing-rerun-before-b.json), plus the
four matching `final2` reports. Ordinary edit/open/no-op guards, the discarded
exploratory run, profiles, counters, memory and complete correctness gates are
summarized in
[`change 0042`](changes/0042-odt-envelope-package-sharing.md).

The ODT final changed-result byte-handoff evidence comprises two balanced
execution cycles with four 500-sample legs per state. Primary raw reports use
the `odt-final-handoff-cycle*` prefix; the matched medium/large read/no-op/edit
matrix uses `odt-final-handoff-guards*`. Profiles, counters, allocation/RSS,
the byte-only ownership distinction and complete correctness gates are
summarized in
[`change 0052`](changes/0052-odt-final-result-byte-handoff.md).

The OPC shared-payload evidence is
[`before A`](results/abba-opc-shared-regeneration-primary-before-a.json),
[`after A`](results/abba-opc-shared-regeneration-primary-after-a.json),
[`after B`](results/abba-opc-shared-regeneration-primary-after-b.json), and
[`before B`](results/abba-opc-shared-regeneration-primary-before-b.json).
No-op/edge guardrails, allocation attribution, RSS and hardware counters are
summarized in
[`change 0021`](changes/0021-opc-shared-regenerated-payload.md).

The generated-local-span evidence is
[`before A`](results/abba-opc-local-span-move-before-a.json),
[`after A`](results/abba-opc-local-span-move-after-a.json),
[`after B`](results/abba-opc-local-span-move-after-b.json), and
[`before B`](results/abba-opc-local-span-move-before-b.json). Repeated small,
edge, tiny and exact-no-op guardrails, allocation attribution, RSS and hardware
counters are summarized in
[`change 0022`](changes/0022-zip-generated-local-span-move.md).

The source-backed overlay evidence is
[`before A`](results/abba-opc-source-overlay-before-a.json),
[`after A`](results/abba-opc-source-overlay-after-a.json),
[`after B`](results/abba-opc-source-overlay-after-b.json), and
[`before B`](results/abba-opc-source-overlay-before-b.json). Source/sink
counters, CPU and memory attribution, failure boundaries and binary/evidence
digests are summarized in
[`change 0037`](changes/0037-opc-source-backed-one-part-publication.md).

The ODT full-text ownership evidence is retained as four short ABBA cycles
under `results/abba-odt-full-text-single-repeat-*.json`. Structured-query,
open, size, exact-no-op and edit guardrails, rejected broad-parser evidence,
allocation attribution, RSS and hardware counters are summarized in
[`change 0023`](changes/0023-odt-full-text-owned-blocks.md).

The ODT indexed-selector evidence is retained as four headline ABBA cycles
under `results/abba-odt-indexed-paragraph-repeat-*.json`, with separate
size/guard reports, Heaptrack attribution, GNU Time RSS and `perf stat`
counters. The complete validation contract and rejected shared-parser design
are summarized in
[`change 0047`](changes/0047-odt-indexed-paragraph-selector.md).

The RTF block-reservation evidence pools six balanced pairs and retains every
sample in the [`primary summary`](results/rtf-body-block-reservation-primary-summary.json).
The [`medium guard summary`](results/rtf-body-block-reservation-medium-guards-summary.json)
covers plain, raw CP-1252 and LZFu with the same six-pair protocol. Allocation,
RSS, profile, counter, tiny-variant and binary-provenance artifacts are indexed
in [`change 0055`](changes/0055-rtf-body-block-reservation.md).

The retained RTF story-length evidence pools two 1,000-sample legs per state
for the paragraph-list and middle-paragraph queries. A reverse-order
2,000-sample pool covers open, full-text, exact stream-save and no-op guards;
allocation, RSS and process-wide profile records are indexed in
[`change 0064`](changes/0064-rtf-retained-story-length.md).

The sparse RTF paragraph-selection evidence pools two 1,000-sample legs per
state for the already-open middle-paragraph query. Reverse-order read/save,
4,000-sample no-op and changed-edit guard pools, iterator-equivalence tests,
variant verification, allocation, RSS and process-wide profile records are
indexed in [`change 0066`](changes/0066-rtf-sparse-paragraph-nth.md).

The DOC PAPX-containment evidence pools five balanced pairs for both the
already-open snapshot paragraph list and complete one-edit/save path, retaining
every sample in the
[`primary summary`](results/doc-papx-containment-primary-summary.json).
Ordinary-reader/no-op and tiny direct distributions are retained in the
[`guard summary`](results/doc-papx-containment-guards-summary.json); profiles,
counters, Heaptrack and GNU Time artifacts are indexed in
[`change 0056`](changes/0056-doc-papx-containment-index.md).

Source-backed cache bytes are bounded by `SourceCacheLimits` but are not yet
charged to hierarchical `Budget`. Raw ZIP preservation is integrated for owned
same-topology OPC mutations and the narrow consuming source-backed one-Part
publisher; broad source-backed semantic editing remains pending.
See [`0005`](changes/0005-xlsx-row-start-index.md),
[`0006`](changes/0006-positional-containers-and-explicit-execution.md), and
[`0007`](changes/0007-source-backed-opc-and-facades.md),
[`0008`](changes/0008-targeted-opc-preservation.md), and
[`0009`](changes/0009-range-source-and-scaling.md), and
[`0010`](changes/0010-docx-pptx-semantic-queries-and-edits.md), and
[`0011`](changes/0011-odf-semantic-baseline-and-ods-snapshot.md), and
[`0012`](changes/0012-docx-coalesced-paragraph-edits.md), and
[`0013`](changes/0013-rtf-semantic-baseline-and-text-paths.md), and
[`0014`](changes/0014-odt-shared-snapshot-bytes.md), and
[`0015`](changes/0015-native-ole2-semantic-baseline.md),
[`0016`](changes/0016-xls-commit-editor-reuse.md),
[`0017`](changes/0017-doc-batched-stream-publication.md), and
[`0018`](changes/0018-ods-row-local-publication.md), and
[`0019`](changes/0019-rtf-parser-state-specialization.md), and
[`0020`](changes/0020-rtf-ascii-transport-batching.md), and
[`0021`](changes/0021-opc-shared-regenerated-payload.md), and
[`0022`](changes/0022-zip-generated-local-span-move.md), and
[`0023`](changes/0023-odt-full-text-owned-blocks.md), and
[`0024`](changes/0024-ppt-slide-order-open-reuse.md), and
[`0025`](changes/0025-xlsx-validated-store-handoff.md).

Consolidated changed-crate tests passed, along with focused changed-crate
warning-denied Clippy and formatter checks. The latest XLSX batch passes 732
unit tests, all integration suites, two doctests and the 32-test harness.
Warning-denied public rustdoc remains blocked by pre-existing broken/private
links, and all-target XLSX Clippy retains the three unrelated findings named in
change 0046. The broad crate-boundary checker likewise retains existing
unclassified workspace edges; no manifest or dependency edge changed. A
workspace all-target/all-feature gate was not run because iWork was explicitly
excluded while its crates are changing independently.

## Accepted results

All latency figures below are warm-memory release-build p50 results from
matched before/after binaries. Each linked change record contains raw-sample
counts, ABBA ordering, mean or interval context, hashes, and memory profiles.

| Workload group | Before | After | Result | Memory result |
|---|---:|---:|---:|---|
| Targeted OPC mutation, four synthetic cells | individual rows in record | individual rows in record | **-84.98% p50 geometric mean**; range -58.24% to -96.41% | Few-large/incompressible peak heap +37.18%; one-shot RSS +22.26% |
| Shared changed-Part handoff, few-large compressible | 1.342 ms | 1.063 ms | **-20.73% p50 / -18.49% mean** | One 4.19 MiB allocation removed; peak heap -3.42%; uninstrumented RSS +0.22% (flat) |
| Exact owned OPC no-op, 16.78 MB incompressible archive | 211.531 ms | 3.443 ms | -98.37% | Peak heap +22.6%; profiler RSS +25.5% because the compressed source is retained alongside eagerly inflated Parts |
| Exact owned OPC no-op, six named many-Part/large-Part cells | individual rows in record | individual rows in record | -99.93% p50 geometric mean | Many-small allocation calls -93.7%; large memory tradeoff above |
| CFB final-root-stream lookup, four 256/2,048-sibling cells | 1.067-7.596 us | 0.451-0.486 us | -84.70% p50 geometric mean | Wide-root peak heap +1.5%; profiler RSS +7.6% for retained exact comparison keys |
| CFB open, four 256/2,048-stream cells | 141.1-963.1 us | 136.8-974.9 us | -1.42% p50 geometric mean | Allocation calls -6.1% to -8.8%; temporary allocations -20.6% to -27.7% |
| Rejected common OLE2 inline recapture allocation reuse, 16 MiB opaque streams | 26.086 ms | 25.404 ms | **-2.61% p50 / -2.30% mean** | Fully reverted as immaterial; p95 +0.54%; isolated publication p50 -6.49% but stages are non-additive |
| OPC rewritten publication, eight named cells | individual rows in record | individual rows in record | -1.65% mean geometric mean; best intended cell -5.49% | Allocation calls -37.0%; peak heap -2.3% |
| Payload-heavy PPT fresh writer | 6.312 ms | 5.035 ms | -20.23% | Peak heap -12.4%; profiler RSS -12.9% |
| Payload-heavy XLS fresh writer | 4.126 ms | 4.065 ms | -1.48%, treated as latency-neutral | Peak heap -9.5%; profiler RSS -12.6% |
| DOCX one paragraph, 10,000-paragraph corpus | 2.945 ms | 2.805 ms | -4.72% p50 / -4.99% mean | 10 collection-growth allocations removed per selector invocation; process peak unchanged |
| DOCX 1% edit/save, 10,000 paragraphs / 100 edits | 487.542 ms | 24.418 ms | **-94.99% p50 (19.97x) / -95.02% mean**; scalar one-edit neutral | Allocation calls -94.11%; peak heap flat; uninstrumented RSS +0.37% (flat) |
| PPTX 1% edit/save, 10,000 text boxes | 399.320 ms | 361.915 ms | -9.37% p50 / -9.37% mean | Allocation calls -11.67%; peak heap flat; profiler RSS +1.28% |
| ODS no-op edit/save, 32,768 cells | 76.894 ms | 67.838 ms | -11.78% p50 / -12.08% mean | Peak heap flat; profiler RSS -0.13% |
| ODS one-cell edit/save, 32,768 cells | 384.150 ms | 376.237 ms | -2.06% p50 / -2.19% mean | Changed package rewrite/readback still dominates |
| ODS row-local one-cell edit/save, 32,768 cells | 359.011 ms | 324.774 ms | **-9.54% p50 / -9.32% mean** | Allocation calls -5.85%; peak heap -27.18%; uninstrumented RSS improved |
| ODS media-rich one-cell edit/save, 2,048 cells + 16 MiB media | 325.902 ms | 310.472 ms | **-4.73% p50 / -5.73% mean** | p95 -7.65%; peak heap -8.78%; existing no-media guard p50 -0.77% |
| ODS durable-patch sharing, 2,048 cells + 16 MiB media | 326.694 ms | 297.958 ms | **-8.80% p50 / -9.07% mean** | p95 -13.85%; redundant package SHA stack absent; 33.58 MB copy site removed; peak heap -1.92%; RSS flat |
| ODS checked row-splice raw publication, 2,048 cells + 16 MiB media | 287.766 ms | 74.365 ms | **-74.16% p50 / -74.17% mean** | p95 -74.11%; instructions -69.04%; unchanged-media rebuild/deflate subtree absent; peak heap/RSS flat |
| ODS shared worksheet archive handoff, 2,048 cells + 16 MiB media | 76.440 ms | 60.140 ms | **-21.32% p50 / -21.30% mean** | p95 -21.15%; peak heap -22.03%; uninstrumented RSS -20.57% |
| ODP media-rich text-box edit/save, 12 slides + 16 MiB media | 227.606 ms | 12.665 ms | **-94.44% p50 / -94.43% mean** | p95 -94.29%; allocation calls +0.52%; peak heap/RSS flat |
| ODS public cell sweep, 32,768 cells | 2.049 ms | 0.374 ms | **-81.74% p50 / -80.72% mean** | Lazy 3,216-byte dense index; peak heap/RSS flat; allocation calls +0.0004% process-wide |
| ODS full cell text, 32,768 cells | 3.047 ms | 1.443 ms | **-52.65% p50 / -52.30% mean** | Existing string clones/join remain; lookup work only is indexed |
| RTF full text, 10,000 paragraphs | 33.095 us | 24.134 us | -27.08% p50 / -25.37% mean | One fragment-vector allocation removed per first materialization |
| RTF one paragraph edit/save, 10,000 paragraphs | 12.408 ms | 9.208 ms | -25.79% p50 / -25.53% mean | Allocation calls -707 over 100 samples; peak heap flat; RSS +0.32% (flat) |
| RTF parser-state follow-up, one paragraph edit/save, 10,000 paragraphs | 8.630 ms | 7.634 ms | **-11.54% p50 / -11.71% mean** | `State::clone` profile frame removed; allocation calls, peak heap and RSS flat |
| RTF transport batching, open, 10,000 paragraphs | 3.159 ms | 2.316 ms | **-26.67% p50 / -26.56% mean** | Per-byte `SmallVec::extend` frame falls from 15.37% to 2.56%; allocations and peak heap flat |
| RTF transport batching, one paragraph edit/save, 10,000 paragraphs | 7.795 ms | 7.307 ms | **-6.26% p50 / -5.73% mean** | Instructions -18.40%; allocation count, peak heap and RSS flat |
| RTF byte-delimiter scan, open, 10,000 paragraphs | 2.479 ms | 2.052 ms | **-17.23% p50 / -17.99% mean** | `tokenize_with_spans` share 17.36% -> 11.06%; instructions -21.27%; peak heap/RSS flat |
| RTF byte-delimiter scan, one paragraph edit/save, 10,000 paragraphs | 7.554 ms | 6.447 ms | **-14.65% p50 / -14.84% mean** | p95 -16.34%; allocations effectively flat; complete edit/save readback unchanged |
| RTF retained body span, one paragraph edit/save, 10,000 paragraphs | 6.053 ms | 5.404 ms | **-10.72% p50 / -10.11% mean** | p95 -8.76%; 588 locator-subtree allocation calls over 20 edits removed; peak heap/RSS flat; candidate parse/readback unchanged |
| RTF bounded body-block reservation, open, 10,000 paragraphs | 2.073 ms | 1.634 ms | **-21.17% p50 / -21.00% mean** | p95 -21.04%; body-vector allocations 264 -> 22 over 22 parses; peak heap -29.73%; uninstrumented RSS flat |
| RTF bounded body-block reservation, one paragraph edit/save | 5.585 ms | 5.503 ms | **-1.46% p50 / -1.75% mean** | p95 -1.87%, p99 -4.11%; complete candidate parse/readback unchanged |
| RTF paragraph list, already-open 10,000-block story | 29.692 us | 25.225 us | **-15.04% p50 / -13.71% mean** | p95 -8.64%; reuses the parser-owned exact text length; allocations and peak heap/RSS flat |
| RTF middle paragraph, already-open 10,000-block story | 18.926 us | 13.780 us | **-27.19% p50 / -25.23% mean** | p95 -14.46%; paragraph boundaries, formatting, exact no-op and complete verification unchanged |
| ODT no-op edit/save, 10,000 paragraphs | 3.950 us | 3.219 us | -18.51% p50 / -29.58% mean | Exactly two allocations and one 28.42 KiB archive copy removed per snapshot; peak heap/RSS flat |
| ODT full text, 10,000 blocks | 4.127 ms | 3.993 ms | **-3.25% p50 / -4.81% mean** | Allocation calls -15.48%, temporary allocations -45.52%; peak heap/RSS flat; open guard disclosed |
| ODT middle paragraph, 10,000 paragraphs | 3.202 ms | 1.647 ms | **-48.56% p50 / -48.33% mean** | Allocation calls -27.05%; peak heap -24.74%; uninstrumented RSS -10.93%; complete EOF validation retained |
| ODP middle slide, 100 slides | 1.019 ms | 0.977 ms | **-4.09% p50 / -4.20% mean** | p95 -5.18%; allocation calls -3.86%; peak heap/RSS flat; complete style/content EOF validation retained |
| ODP exact no-op transaction/save, 100 slides | 1.728 ms | 0.692 ms | **-59.96% p50 (2.50x) / -59.92% mean** | Large changed edit/save p50 -20.78%; allocations -20.13%; complete package/security and final readback retained; peak heap/RSS flat |
| ODP one-slide edit/save, 100 source slides | 3.573 ms | 2.417 ms | **-32.35% p50 / -32.92% mean** | p95 -35.95%; allocations -16.71%; final package reopen/audits/media checks retained; peak heap/RSS flat |
| ODT media-rich paragraph edit/save, 200 paragraphs + 16 MiB media | 249.177 ms | 11.001 ms | **-95.58% p50 / -95.63% mean** | p95 -95.43%; allocation calls -6.71%; peak heap flat; RSS -0.59% |
| ODT direct snapshot sharing, 200 paragraphs + 16 MiB media | 32.270 ms | 7.798 ms | **-75.84% p50 / -73.84% mean** | Two archive-sized copies removed; p95 -75.41%; peak heap/RSS flat |
| ODT compact-audit package sharing, 200 paragraphs + 16 MiB media | 7.773 ms | 5.407 ms | **-30.44% p50 / -31.36% mean** | Three archive-sized audit copies removed; p95 -32.41%; allocations -0.57%; peak heap/RSS flat; exact no-op +39 ns disclosed |
| ODT envelope-classification sharing, 200 paragraphs + 16 MiB media | 5.555 ms | 4.921 ms | **-11.40% p50 / -11.95% mean** | One archive-sized envelope copy and two allocations/commit removed; p95 -12.19%; peak heap/RSS flat; large exact no-op +152 ns disclosed |
| ODT final changed-result byte handoff, 200 paragraphs + 16 MiB media | 5.216 ms | 4.030 ms | **-22.74% p50 / -22.56% mean** | One 16.79 MB result copy and redundant parse removed; p95 -21.48%; allocation calls -3.46%; independent final reopen and peak heap/RSS retained |
| ODT 1% paragraph edit/save, 10,000 paragraphs / 100 replacements | 906.439 ms | 15.615 ms | **-98.28% p50 (58.05x) / -98.27% mean** | One mutable candidate/publication/reopen/audit replaces 100; allocations -96.13%; peak heap and uninstrumented RSS flat; tool-inclusive RSS +9.93% disclosed |
| Native XLS one-cell edit/save, 8,192 cells | 1.777 ms | 1.639 ms | **-7.72% p50 / -7.90% mean** | Allocation calls -1.19%; peak heap and uninstrumented RSS flat |
| Native XLS fixed-width numeric edit/save, 8,192 cells | 1.582 ms | 1.458 ms | **-7.83% p50 / -7.37% mean** | Complete public Workbook validation retained; peak heap -5.54%, RSS flat |
| Native DOC one-paragraph edit/save, 512 paragraphs | 1.506 ms | 1.348 ms | **-10.52% p50 / -10.48% mean** | Duplicate publication-site allocations nearly halved; peak heap and uninstrumented RSS flat |
| Native DOC open, 512 paragraphs | 790.727 us | 348.679 us | **-55.91% p50 / -55.78% mean** | Physical PieceTable scan self cycles 36.89% -> 4.17%; allocation calls +0.009%; peak heap and uninstrumented RSS flat |
| Native DOC one-paragraph edit/save after PieceTable index, 512 paragraphs | 1.379 ms | 0.950 ms | **-31.08% p50 / -31.68% mean** | Same private index accelerates mandatory candidate/public readbacks; patch/inverse and exact output checks unchanged |
| Native DOC open after PieceTable index, 512 paragraphs | 343.503 us | 304.199 us | **-11.44% p50 / -11.87% mean** | Paragraph-style validation 4.44% -> 0.83% self cycles; allocation calls -18.61%; peak heap and uninstrumented RSS flat |
| Native DOC one-paragraph edit/save after style cache, 512 paragraphs | 912.288 us | 875.736 us | **-4.01% p50 / -4.23% mean** | Same one-entry cache accelerates mandatory candidate/public readbacks; patch/inverse and exact output checks unchanged |
| Native DOC paragraph list after style cache, 512 paragraphs | 454.100 us | 358.414 us | **-21.07% p50 / -20.93% mean** | CHPX range query changes from a full scan per paragraph to binary start plus matching slice; p95 -20.00%; allocations and peak heap/RSS flat |
| Native DOC exact-source paragraph list after CHPX index, 512 paragraphs | 206.644 us | 168.142 us | **-18.63% p50 / -19.04% mean** | Ordered piece/PAPX containment uses predecessor binary search; instructions -26.13%; allocations and peak heap flat |
| Native DOC one-paragraph edit/save after PAPX containment index | 888.602 us | 817.424 us | **-8.01% p50 / -7.88% mean** | p95 -7.71%, p99 -8.37%; patch/inverse, candidate owner and independent public readback unchanged |
| Native PPT root snapshot open, 144 shapes | 37.522 us | 34.227 us | **-8.78% p50 / -10.58% mean** | Allocation calls -5.01%, temporary allocations -12.22%; peak heap and uninstrumented RSS flat |
| Native PPT direct text edit/save, 144 shapes | 206.209 us | 177.089 us | **-14.12% p50 / -15.39% mean** | Allocation calls -3.53%, temporary allocations -6.05%; peak heap/RSS flat; minor faults +315.43% with zero major faults |
| Native PPT root text edit/save, 144 shapes | 352.306 us | 286.805 us | **-18.59% p50 / -17.83% mean** | p95 -16.58%; allocation calls -6.54%; peak heap and uninstrumented RSS flat; custom limits retain full reopen |
| XLSX one-cell commit + first read, 4,096 cells | 4.431 ms | 3.402 ms | **-23.23% p50 / -23.15% mean** | Allocation calls -21.01%; peak heap +4.29%; unrestricted dense-wide retention rejected |
| Rejected XLSX 1% commit + save, 4,096 cells | 15.235 ms | 14.990 ms | -1.61% p50 / -1.26% mean | Fully reverted as immaterial; p99 +0.18%, peak heap flat |
| Rejected XLSX 1% commit + save, 131,072 cells | 514.926 ms | 511.407 ms | -0.68% p50 / -0.66% mean | Fully reverted as immaterial; process allocation calls -0.0623% |
| Source-backed XLSX calculation-metadata publication, 12 Parts + 16 MiB media | 215.457 ms | 1.612 ms | **-99.2519% p50 (133.67x) / -99.2507% mean** | Materializations 12 -> 1; allocation calls -10.81%; peak heap flat; uninstrumented RSS -1.20% |
| Source-backed XLSX page-break publication, 12 Parts + 16 MiB media | 216.789 ms | 4.647 ms | **-97.86% p50 (46.65x) / -97.86% mean** | Materializations 12 -> 2; allocation calls -15.95%; peak heap and uninstrumented RSS flat |
| Source-backed XLSX page-margin publication, 12 Parts + 16 MiB media | 216.799 ms | 4.492 ms | **-97.93% p50 (48.26x) / -97.93% mean** | Materializations 12 -> 2; allocation calls -12.10%; peak heap and uninstrumented RSS flat |

The underlying records are:

- [`0001-opc-publication-plan.md`](changes/0001-opc-publication-plan.md)
- [`0002-cfb-lookup-and-sector-buffers.md`](changes/0002-cfb-lookup-and-sector-buffers.md)
- [`0003-legacy-owned-stream-handoff.md`](changes/0003-legacy-owned-stream-handoff.md)
- [`0004-opc-exact-owned-source.md`](changes/0004-opc-exact-owned-source.md)
- [`0008-targeted-opc-preservation.md`](changes/0008-targeted-opc-preservation.md)
- [`0009-range-source-and-scaling.md`](changes/0009-range-source-and-scaling.md)
- [`0010-docx-pptx-semantic-queries-and-edits.md`](changes/0010-docx-pptx-semantic-queries-and-edits.md)
- [`0011-odf-semantic-baseline-and-ods-snapshot.md`](changes/0011-odf-semantic-baseline-and-ods-snapshot.md)
- [`0012-docx-coalesced-paragraph-edits.md`](changes/0012-docx-coalesced-paragraph-edits.md)
- [`0013-rtf-semantic-baseline-and-text-paths.md`](changes/0013-rtf-semantic-baseline-and-text-paths.md)
- [`0014-odt-shared-snapshot-bytes.md`](changes/0014-odt-shared-snapshot-bytes.md)
- [`0015-native-ole2-semantic-baseline.md`](changes/0015-native-ole2-semantic-baseline.md)
- [`0016-xls-commit-editor-reuse.md`](changes/0016-xls-commit-editor-reuse.md)
- [`0017-doc-batched-stream-publication.md`](changes/0017-doc-batched-stream-publication.md)
- [`0018-ods-row-local-publication.md`](changes/0018-ods-row-local-publication.md)
- [`0019-rtf-parser-state-specialization.md`](changes/0019-rtf-parser-state-specialization.md)
- [`0020-rtf-ascii-transport-batching.md`](changes/0020-rtf-ascii-transport-batching.md)
- [`0021-opc-shared-regenerated-payload.md`](changes/0021-opc-shared-regenerated-payload.md)
- [`0022-zip-generated-local-span-move.md`](changes/0022-zip-generated-local-span-move.md)
- [`0023-odt-full-text-owned-blocks.md`](changes/0023-odt-full-text-owned-blocks.md)
- [`0024-ppt-slide-order-open-reuse.md`](changes/0024-ppt-slide-order-open-reuse.md)
- [`0025-xlsx-validated-store-handoff.md`](changes/0025-xlsx-validated-store-handoff.md)
- [`0026-ppt-text-edit-resolver-reuse.md`](changes/0026-ppt-text-edit-resolver-reuse.md)
- [`0027-ods-adaptive-cell-locator.md`](changes/0027-ods-adaptive-cell-locator.md)
- [`0028-xls-terminal-render-handoff-rejected.md`](changes/0028-xls-terminal-render-handoff-rejected.md)
- [`0029-rtf-transport-and-producer-coverage.md`](changes/0029-rtf-transport-and-producer-coverage.md)
- [`0030-xlsx-action-plan-flattening-rejected.md`](changes/0030-xlsx-action-plan-flattening-rejected.md)
- [`0031-ods-unchanged-media-preservation.md`](changes/0031-ods-unchanged-media-preservation.md)
- [`0032-xlsx-no-extension-scan.md`](changes/0032-xlsx-no-extension-scan.md)
- [`0033-ole-common-publication-handoffs-rejected.md`](changes/0033-ole-common-publication-handoffs-rejected.md)
- [`0034-odp-unchanged-media-preservation.md`](changes/0034-odp-unchanged-media-preservation.md)
- [`0035-odt-content-only-paragraph-publication.md`](changes/0035-odt-content-only-paragraph-publication.md)
- [`0036-ole-common-stage-attribution.md`](changes/0036-ole-common-stage-attribution.md)
- [`0037-opc-source-backed-one-part-publication.md`](changes/0037-opc-source-backed-one-part-publication.md)
- [`0038-odt-direct-snapshot-sharing.md`](changes/0038-odt-direct-snapshot-sharing.md)
- [`0039-docx-source-backed-semantic-publication.md`](changes/0039-docx-source-backed-semantic-publication.md)
- [`0040-rtf-byte-delimiter-scanning.md`](changes/0040-rtf-byte-delimiter-scanning.md)
- [`0041-odt-compact-audit-package-sharing.md`](changes/0041-odt-compact-audit-package-sharing.md)
- [`0042-odt-envelope-package-sharing.md`](changes/0042-odt-envelope-package-sharing.md)
- [`0043-rtf-decoded-body-ownership-rejected.md`](changes/0043-rtf-decoded-body-ownership-rejected.md)
- [`0044-pptx-source-backed-semantic-publication.md`](changes/0044-pptx-source-backed-semantic-publication.md)
- [`0045-odt-coalesced-paragraph-publication.md`](changes/0045-odt-coalesced-paragraph-publication.md)
- [`0046-xlsx-source-backed-calculation-metadata-publication.md`](changes/0046-xlsx-source-backed-calculation-metadata-publication.md)
- [`0047-odt-indexed-paragraph-selector.md`](changes/0047-odt-indexed-paragraph-selector.md)
- [`0048-rtf-retained-body-source-span.md`](changes/0048-rtf-retained-body-source-span.md)
- [`0049-odp-indexed-slide-selector.md`](changes/0049-odp-indexed-slide-selector.md)
- [`0050-doc-piece-table-physical-index.md`](changes/0050-doc-piece-table-physical-index.md)
- [`0051-doc-adjacent-style-baseline-cache.md`](changes/0051-doc-adjacent-style-baseline-cache.md)
- [`0052-odt-final-result-byte-handoff.md`](changes/0052-odt-final-result-byte-handoff.md)
- [`0053-doc-chpx-range-index.md`](changes/0053-doc-chpx-range-index.md)
- [`0054-ods-shared-durable-patch-blobs.md`](changes/0054-ods-shared-durable-patch-blobs.md)
- [`0055-rtf-body-block-reservation.md`](changes/0055-rtf-body-block-reservation.md)
- [`0056-doc-papx-containment-index.md`](changes/0056-doc-papx-containment-index.md)
- [`0057-ods-row-splice-raw-publication.md`](changes/0057-ods-row-splice-raw-publication.md)
- [`0058-ods-exact-noop-handoff.md`](changes/0058-ods-exact-noop-handoff.md)
- [`0059-xls-fixed-numeric-inventory-carry.md`](changes/0059-xls-fixed-numeric-inventory-carry.md)
- [`0060-odp-snapshot-slide-projection-reuse.md`](changes/0060-odp-snapshot-slide-projection-reuse.md)
- [`0061-xlsx-source-backed-page-break-publication.md`](changes/0061-xlsx-source-backed-page-break-publication.md)
- [`0062-ppt-root-text-publication-adoption.md`](changes/0062-ppt-root-text-publication-adoption.md)
- [`0063-pptx-atomic-source-backed-shape-text-batch.md`](changes/0063-pptx-atomic-source-backed-shape-text-batch.md)
- [`0064-rtf-retained-story-length.md`](changes/0064-rtf-retained-story-length.md)
- [`0065-odp-final-snapshot-handoff.md`](changes/0065-odp-final-snapshot-handoff.md)
- [`0066-rtf-sparse-paragraph-nth.md`](changes/0066-rtf-sparse-paragraph-nth.md)
- [`0067-xlsx-source-backed-page-margin-publication.md`](changes/0067-xlsx-source-backed-page-margin-publication.md)
- [`0068-ods-shared-worksheet-archive-handoff.md`](changes/0068-ods-shared-worksheet-archive-handoff.md)

The DOC ownership-transfer variant was rejected and removed after a 58.42%
p50 regression. The earlier full-rewrite mutated-OPC guardrail was neutral on
incompressible data; targeted raw publication supersedes it only for the
strictly proved same-topology owned-source case. Fallback still uses that
validated full rewrite. Rejected, fallback and memory results are retained
rather than hidden in an aggregate.

An ODS target-package adoption candidate was likewise removed after large
one-cell edit/save improved only 0.44% p50 and p95 regressed 0.30%. The existing
package/readback boundary remains; no production or test code from that
candidate is retained.

An ODT final-document adoption candidate was also fully reverted. It improved
large one-edit/save p50 5.70%, but a dedicated medium one-paragraph read guard
regressed 6.33% mean and 17.64% p95. The accepted snapshot-byte sharing remains;
the rejected parsed-document retention contributes no production or test code.
Change 0052 is deliberately narrower: it shares only immutable final bytes and
retains a fresh independent final reopen; its same guard stays within 3% p50
and mean with a better p95.

The first ODT full-text ownership candidate also moved strings for structured
list and one-paragraph callers. Their large-corpus p50 regressed 5.71% and
5.30%, respectively, so that broad version was removed. The accepted private
full-text mode retains the original structured path; the rejected raw reports
remain linked from change 0023.

## Work removed

- Exact unchanged owned OPC publication no longer regenerates manifests,
  reconstructs ZIP records, or recompresses logical Parts. It copies the
  complete validated source to the caller's sequential sink in writes bounded
  to 64 KiB and verifies complete output in the benchmark.
- Targeted same-topology OPC publication no longer recompresses unchanged
  Parts. It audits the ordinary publication plan, regenerates only changed
  payload/relationship/content-type closures, and raw-copies unchanged local
  spans and central records, including unknown non-part members.
- The low-level source-backed one-Part publisher no longer converts the
  positional package into an eager owning package or recompresses every Part.
  It materializes and validates the selected original payload, regenerates that
  member, and raw-copies every other member while monitoring source version.
- The changed ordinary Part now shares its already-owned immutable logical
  payload with ZIP regeneration rather than allocating and copying it again.
  Generated XML and the required compressor/archive buffer stay owned.
- After the generated member has passed complete ZIP validation, its local span
  now moves into the prepared entry instead of being allocated and copied a
  second time. Central-directory framing remains separately retained.
- Rewritten OPC publication constructs and audits generated XML and stable
  Part order once before emission rather than once for validation and again
  for writing.
- CFB lookup follows the validated sibling-tree ordering with SID-aligned
  cached comparison keys rather than scanning the complete sibling tree.
- CFB FAT/DIFAT/MiniFAT parsing reuses a bounded sector buffer, MiniFAT decodes
  into its final table, and directory sectors read into their final buffer.
- Fresh XLS and PPT writers transfer already-owned generated stream buffers to
  CFB without a second payload copy. DOC deliberately retains its measured
  faster exact-sized copy.
- Native XLS changed commit reuses the already rendered/reopened object editor
  instead of discarding one BIFF owner parse and reopening/capturing the CFB a
  second time before final validation.
- Native DOC applies ordinary WordDocument and table-stream replacements to
  one isolated object-editor candidate and renders/reopens the CFB once rather
  than once per replacement.
- Native DOC paragraph FKP parsing reuses one resolved initial-style baseline
  across repeated source runs instead of reconstructing and revalidating the
  same inheritance chain. Direct properties, piece modifiers and direct style
  switches still execute independently for every PAPX.
- DOCX one-paragraph selection no longer constructs the complete paragraph
  collection, and source-backed paragraph counts no longer construct any
  paragraph views. Complete XML validation and limits still run.
- Canonical multi-paragraph DOCX text edits now plan disjoint replacements and
  emit/reparse one candidate XML document instead of rebuilding and reparsing
  the complete main document once per paragraph. Durable patches remain
  ordinary source-checked paragraph operations with complete final readback.
- Repeated PPTX shape-text edits no longer parse the selected slide scene a
  second time solely to map the already selected shape to its raw XML span.
- DOCX plaintext package output exposes the underlying forward-only OPC sink
  instead of imposing an unused `Seek` bound.
- ODS unified snapshot construction reuses its one validated package for full
  facade readback instead of cloning package bytes and parsing the package a
  second time.
- Eligible same-topology ODS worksheet commits serialize only changed modeled
  rows instead of regenerating every worksheet row; untouched XML source spans
  are copied exactly and structural edits retain full-table fallback.
- Eligible compact ODS `content.xml` commits raw-copy every other validated ZIP
  member. Source/target effect checks use exact local and central member bytes
  to skip unchanged payload inflation only while the manifest is also exact;
  every unproved case retains logical comparison and established rebuild.
- Eligible same-topology ODS row edits now carry their already checked exact
  source ranges through that raw package publisher. They no longer fall back
  to recompressing unchanged media merely because the flattened result cannot
  be rediscovered as one conservative maximal diff.
- The adjacent unified ODS worksheet handoff now moves and shares its exact
  archive allocation across nested worksheet snapshots, package parsing,
  commit readback and candidate validation instead of repeatedly copying it.
  Failure paths restore the original bytes and allocation; durable patch and
  final validation boundaries are unchanged.
- RTF first full-text materialization retains only a byte count during parse,
  then allocates the final string once and copies blocks in one pass instead of
  allocating and joining a temporary fragment vector.
- RTF borrowed stories now receive that already validated byte count instead
  of rescanning every retained style block to establish paragraph and inline
  iterator endpoints.
- RTF canonical text emission writes ordinary ASCII spans in chunks instead of
  one formatted write per character. Text-only commits skip paragraph-property
  vectors/scans, and a successful paragraph selector stops at its target.
- Ordinary RTF body-text flushes no longer clone the complete parser state.
  They copy the effective encoding and block properties; insertion/deletion
  runs alone retain full state for revision author/date and exact range data.
- All-ASCII RTF source tokens now enter transport buffers in one extension
  rather than one generic `SmallVec::extend` call per character. The checked
  byte-valued non-ASCII and invalid-Unicode fallback is unchanged.
- ODT transaction snapshots created from an already validated `Document` clone
  its private immutable package handle instead of allocating and copying the
  complete archive. Direct snapshot byte ingress keeps independent validation.
- ODT full-text extraction moves each parser-created validated block string
  into the element and consumes it into final output instead of cloning the
  string at both private handoff boundaries. Structured block queries retain
  their original ownership behavior.
- Eligible XLSX changed sheets move their exact commit-validated semantic
  store into the published snapshot after byte and style/shared-string lineage
  checks. Retention is capped at 4,096 cells and 1 MiB of worksheet XML; larger
  sheets keep the cold-cache path.
- Direct PPT text-edit setup uses its full editor preflight to resolve the
  selected persisted record instead of opening and capturing the CFB a second
  time. Commit still opens a fresh editor and performs exact source comparison,
  publication, complete snapshot reopen and semantic readback.
- Repeated public ODS cell queries lazily build one private, sheet-aligned
  locator after 64 successful lookups. Direct runs retain compact row
  descriptors; repeated runs add cumulative endpoints under a 4 MiB cap and
  any build failure permanently falls back to linear lookup.

No unsafe code, ambient I/O, dependency edge, public archive type, or global
synchronization primitive was introduced. Exact-source authorization is
revoked conservatively on every mutable OPC entry point, including failed and
semantic no-op calls. Borrowed ingress, topology-changing edits, and unsupported
ZIP layouts use the fully validated owning rewrite path; the narrow
source-backed publisher instead returns a typed zero-output refusal.

## Evidence and verification

The standalone harness provides 128 selectable cases and a 198-record default
matrix across deterministic ZIP/OPC, positional CFB/OPC, source-backed XLSX,
public DOC/XLS/PPT writer and semantic corpora, and DOCX/PPTX/RTF/ODT/ODS/ODP
semantic corpora. RTF includes deterministic raw CP-1252 and LZFu inputs plus
a content-addressed producer watermark; its separate native `relsize` chain is
an offline correctness gate rather than a timed paragraph case.
It records
p50/p95/p99, raw samples, mean, sample deviation, Student's-t 95% mean interval,
corpus/output hashes, environment, bounded sequential-write behavior,
deterministic logical/physical range distributions, and exact execution
tasks/bytes. CI runs a non-gating deterministic smoke check and a
scheduled/manual release matrix.

The current local evidence includes consolidated changed-crate tests with
byte/hash checks, focused changed-crate warning-denied Clippy, formatter and
diff checks, YAML parsing, and JSON parsing. The umbrella all-feature `litchi`
attempt exhausted local disk, so it is not represented as a passing umbrella
gate. The historical stage-1 all-feature gate and its pre-existing Cargo
warning about DOCX/PPTX example output name `owner_native_smoke` remain scoped
to that earlier capture.

The repository-wide warning-denied rustdoc command remains blocked by existing
broken/private intra-doc links in unchanged OPC, DOC, XLS, and PPT files. The
dependency-direction checker unit suite passes, while the live policy check
reports existing unclassified edges (including `litchi-opc -> xml-minifier`
and several dev-only `-> soapberry-zip` edges); this tranche changes no Cargo
manifest or dependency edge. These pre-existing gate failures are not counted
as passing verification.

During the stage-1 capture, hardware counters were unavailable because that
host had `perf_event_paranoid=4`. The later targeted-OPC capture ran after the
environment reported `1`: matched process counters show cycles -69.21% and
instructions -69.85% for that one save cell. No counter claim is retroactively
made for stage 1 or generalized to other workloads. Later XLS, DOC and ODS
change records also retain matched process counters; the row-local ODS workload
reports cycles -5.47%, instructions -6.92%, and cache misses -6.58%.
The later RTF parser-state workload reports cycles -10.50%, instructions
-9.28%, and cache references -8.61%; its profiler removes the former 8.53%
exclusive state-clone frame. The subsequent RTF transport workload reports
cycles -11.22%, instructions -18.40%, and branches -14.04%; its per-byte
`SmallVec::extend` share falls from 15.37% to 2.56% on open. The OPC
shared-payload follow-up removes one 4.19 MiB allocation, cuts peak heap 3.42%,
task clock 21.08%, cycles 19.41% and cache misses 31.12% on its matched
few-large compressible process. The local-span follow-up removes the next 4.20
MiB allocation, cuts peak heap another 3.20% and task clock 2.11%; its other
major hardware counters stay within 5%. Uninstrumented RSS is flat for both.
The source-backed one-Part publisher removes three unnecessary Part
materializations and recompressions on its four-Part corpus: operation p50
falls 73.12%, instructions 65.42%, allocation calls 6.41%, peak heap 3.20% and
maximum observed uninstrumented RSS 3.26%. Physical source bytes remain flat
because every unchanged compressed span is still copied to the output.
The ODT full-text follow-up removes 420,019 allocation calls over ten samples,
cuts temporary allocations 45.52%, task clock 2.39%, instructions 2.51% and
cache misses 13.05%; peak heap and uninstrumented RSS remain flat.
The PPT root-snapshot follow-up removes 45 allocation calls per open, cuts
task clock 6.56%, instructions 9.57% and cycles 6.85%, and keeps peak heap and
uninstrumented RSS flat. Its 15.00% cache-miss increase is disclosed rather
than presented as a locality improvement.
The bounded XLSX changed-sheet handoff removes the duplicate public first-read
parse on eligible commits: task clock falls 24.25%, instructions 23.05%,
cycles 24.29% and allocation calls 21.01% on the medium attribution process.
Peak heap rises 4.29% under the bound and uninstrumented RSS is flat. The
unrestricted dense-wide prototype's 8.99% peak-heap increase triggered its
rejection and the retained cold-cache fallback.
The direct PPT text-edit follow-up removes one repeated editor open and cuts
task clock 3.60%, allocation calls 3.53% and temporary allocations 6.05%.
Peak heap and uninstrumented RSS remain flat. Its +315.43% minor-fault trigger
has zero major faults and is disclosed rather than presented as a
memory-locality improvement.
The PPT root text-publication handoff removes the immediate second root reopen
after the validated text owner: scoped p50 falls 18.59%, task clock 8.76%,
instructions 6.75%, and allocation calls 6.54%. Peak heap and uninstrumented
RSS remain flat; nondefault limits retain the prior complete reopen.
The ODS cell-locator follow-up reduces the large public sweep's p50 81.74% and
the full-text aggregate's p50 52.65%. Matched sweep-process task clock falls
10.28%, cycles 9.74%, instructions 6.72% and cache misses 7.90%; peak heap,
Heaptrack RSS and uninstrumented RSS remain flat. The 318 added allocation
calls across 105 snapshots (+0.0004%) and the retained 3,216-byte dense index
are disclosed.
The ODS unchanged-media follow-up reduces its 16 MiB media-rich edit/save p50
4.73%, mean 5.73% and p95 7.65%. Peak heap falls 8.78%, task clock 3.54% and
cache references 5.92%; allocation calls rise 0.11%, while branch/cache misses
move +0.42%/+1.29% and are disclosed. The existing medium no-media case remains
slightly better at -0.77% p50.
The ODS row-provenance follow-up reduces the same media-rich edit/save p50
74.16%, mean 74.17% and p95 74.11%. Instructions fall 69.04%, branches 72.96%
and branch misses 94.07%; allocation calls fall 1.26%, while peak heap and
uninstrumented RSS remain flat. Tiny/medium/large open, read and no-op p50
guards remain within the 3% gate or improve.
The ODS shared-worksheet ownership follow-up reduces the remaining media-rich
transaction p50/mean/p95 by 21.32%/21.30%/21.15%. Peak heap falls 22.03%,
uninstrumented RSS 20.57%, cache misses 23.50% and page faults 27.31%; large
ordinary open/no-op/one-edit guards remain within 1.6%.
Lock-wait evidence remains missing.

## Remaining highest-impact work

The largest remaining limitation is the incomplete migration from eager OPC to
source-backed CRUD: selective open, source versions, finite cache,
single-flight, and a low-level consuming one-Part publisher now exist, but
cache bytes are not yet charged to the hierarchical budget and broad semantic
edit/patch coverage is incomplete. Raw ZIP preservation is integrated for
eager owned same-topology mutation and this narrow source-backed case; format
facades, topology changes, signatures and real-producer/media matrices remain.
The changed-Part handoff and post-validation local-span copies are removed;
the required selected-Part/compressor buffer remains to be attributed and
reduced independently.

Other high-priority gaps are cold-filesystem and real range-source matrices,
threshold tuning/contention work beyond the committed explicit scaling curves,
and broad format-semantic CRUD coverage beyond the generated text/grid slices
(bulk action distinctions, dependency-copy, merge/split, patch timing, repair,
security, malformed and real-producer corpora, plus broader ODF and RTF
coverage). Native DOC/XLS/PPT semantic baselines now have accepted XLS
editor-reuse, DOC batched-publication, PPT root-open reuse, direct PPT
text-edit resolver reuse, and checked PPT root text-publication adoption.
Native DOC also indexes physical
PieceTable intervals and reuses one resolved PAPX initial-style baseline after
distinct profiles attributed 36.89% of large-open self cycles to scalar FKP
range mapping and 6.94% to repeated style resolution/validation; full
validation remains.
Remaining native work requires new attribution inside the retained final
owner/public-reader validation layers. The rejected XLS terminal-render
handoff is not a reusable shortcut for those checks. The new opaque-heavy
common OLE2 case rejects direct shared writer payloads (+32.02% p50), an
editor-wide validated-render cache, and inline recapture allocation reuse. The
last improves isolated publication 6.49% p50 but only 2.61% end to end, so it
too was reverted. ODT full-text block
ownership is accepted, and repeated ODS facade cell lookup now has a bounded
lazy index. ODP one-slide lookup now retains only the selected semantic
projection, and its editing snapshot reuses its validated complete slide
projection during transaction staging. Compact ODS and content-only ODP/ODT
edits preserve unchanged ZIP members, and eligible ODS row-range provenance
now survives through raw package emission. Broader ODF source-backed reads, repeated independent
ODT/ODP scans, resource-adding/structural publications, package-parse reuse and
structural-edit profiles remain open.
XLSX changed-sheet validation can now seed a bounded first-read cache. Direct
writer-local action regrouping was immaterial and reverted; distinct bulk
actions, any larger planning/emission coalescing, large-sheet retention,
source-backed editable publication, structural changes and broad preservation
matrices remain independent work.
The rejected direct ODS target-package and parsed ODT final-document adoptions
are not evidence that those broader paths are complete or that validation
should be weakened. Change 0052 shares final bytes only and retains the
independent parse boundary.
iWork work is deliberately deferred while the `iwa-*` crates are changing
independently.
The scenario-by-scenario gap map and next case queue are in
[`CRUD_COVERAGE.md`](CRUD_COVERAGE.md).
The ranked source-level queue and path maps are maintained in
[`HOTSPOTS.md`](HOTSPOTS.md), and architectural gates are in
[`ADR_COMPLIANCE.md`](ADR_COMPLIANCE.md).
