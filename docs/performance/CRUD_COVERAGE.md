# Performance CRUD coverage

Date: 2026-08-22

This is a coverage map, not a completion claim. It compares the 344 selectable
benchmark cases and the explicitly labeled correctness-only APIs with
`docs/CRUD_Scenario_Checklist.md`. Generic ZIP/OPC/CFB substrate measurements
do not certify format-semantic CRUD behavior, and API-only coverage is not a
performance claim.

| Required scenario | Current status | Coverage evidence |
|---|---|---|
| Open and identify format | Partial | ZIP/OPC/CFB plus owned DOC/XLS/PPT/RTF/XLSX and source-backed XLSX/ODT open; the separate opt-in XLSB binary now checks facade identification and owner open over a fixed real POI fixture. Opt-in bounded RTF/XLS/DOCX/PPTX/generic-ODF reports exercise format validation, while RTF still covers plain, raw CP-1252, LZFu and a real-producer watermark input. Change 0120 adds matched eager/source-path PPTX ordinary-root open controls using `litchi::Presentation::open(path)` and full untimed parity checks. Change 0187 routes high-level `litchi::Workbook::open(path)` for XLSX through its source-backed owner and adds open/open-plus-projection evidence; change 0188 adds matched warm DOCX/PPTX fresh-open-plus-query lifecycle evidence over fixed media-rich corpora but accepts no latency statistic because drift gates fail. Change 0191 routes high-level `litchi::Document::open(path)` for ODT through one retained source-backed owner while preserving OOXML precedence and typed source errors; change 0192 closes the withheld open-only warm p50/p99 latency evidence for that path on a bit-identical-binary rerun while mean/p95 remain withheld. Broader smart-detection handoff coverage remains incomplete |
| List semantic children without payloads | Partial | XLS/XLSX/XLSB/ODS sheets, DOC/RTF/DOCX/ODT paragraphs and PPT/PPTX/ODP slides; the XLSB case checks exact worksheet count/name parity outside timing. Opt-in `docx_section_inventory` now lists the exact source-backed DOCX section descriptors, PPTX ordinary-root `slide_count` proves catalog-only source replay while `list_slides` intentionally materializes all owned slide payloads, change 0122 adds ODP source-backed open/one-middle-slide logical-read guards over a media-rich package, change 0123 adds unified-root ODP filesystem open/list parity guards, and change 0124 adds unified-root ODS filesystem worksheet-name/count/text parity plus direct typed source-read evidence; broader section/edit matrices remain missing |
| Query one property or named object | Partial | XLS/XLSX/ODS cells, one DOC/RTF/DOCX paragraph, an indexed fully validated ODT paragraph, one PPT shape and one PPTX/ODP slide; native PPT now also has matched eager/source-backed selected-shape query-only and fresh-open-plus-query controls, change 0120 adds eager/source-root `slide_count` plus selector-first slide-100 controls with independent exact range evidence, change 0122 adds matched eager/source-backed ODP middle-slide query evidence over retained media, change 0123 adds matched unified-root ODP middle-slide filesystem queries, and change 0124 adds matched typed ODS eager/source selected-cell queries after owner preparation; the already-open RTF paragraph query reuses its parser-derived exact story length and cardinality, while explicit sparse `nth` skips discarded-view construction; change 0188 adds warm PPTX open-plus-count/selected-slide and DOCX open-plus-paragraph-count evidence but accepts no latency statistic because drift gates fail; change 0254 adds accepted plan-only DOCX evidence for eight repeated `Snapshot::plan_target_urls` calls on a prepared immutable 49-story/1,152-link snapshot (88.225897%-91.592577% lower across p50/mean/p95/p99), not an end-to-end speedup; broader properties/images remain missing |
| Read one cell/paragraph/slide/image/Part | Partial | XLS/XLSX/ODS cells, DOC/RTF/DOCX paragraphs, indexed ODT paragraphs, PPT/PPTX/ODP text objects and generic OPC Part. Change 0120 adds a source-path PPTX slide-100 read with no unselected-slide/media payload overlap in its separate untimed replay; change 0122 adds an explicit ODP selected-media replay that proves one complete compressed `Pictures/` range and reports non-Pictures bytes separately; change 0123 adds root-path ODP archive/member/hash parity and direct typed source replay evidence; change 0124 adds typed ODS selected-cell and selected-media controls, with compressed-range and uncompressed-payload evidence kept distinct. Three opt-in OPC source-cache selectors add exact managed-Budget boundary evidence plus finite-control/managed same-Part and fixed-work disjoint-Part contention across `1/2x`, `1x`, and `2x` capacities and capped worker widths; their deterministic source delay is correctness evidence, not a latency claim. Semantic image selection remains missing |
| Scan all cells/paragraphs/slides | Covered for generated native/OOXML/RTF/ODF text corpora | XLS/XLSX/ODS cell scans, including the isolated ODS public-cell scalar and ordered batch sweeps; DOC/RTF/DOCX/ODT paragraph enumeration and PPT/PPTX/ODP slide/text enumeration. Change 0120 adds an ordinary-root PPTX full owned-slide list; its source replay reads all slide payloads and no media, while eager/source parity and names/text are checked outside timing |
| Full text extraction | Covered for generated DOC/PPT/RTF/DOCX/PPTX/ODT/ODS/ODP; partial for XLSB | Complete deterministic text or row-major cell text is checked; the separate XLSB binary adds exact bytes/digest parity for one real POI workbook. RTF additionally verifies raw CP-1252/LZFu text and the body plus public header-shape projection of a real-producer watermark; ODT consuming block ownership is measured in change 0023, and change 0191 records zero additional logical source reads plus accepted warm p50/mean/p95/p99 reductions for the retained high-level source-backed ODT open-plus-full-text lifecycle; change 0188 measures DOCX open-plus-full-text but withholds latency; broader real-producer/media-heavy corpora remain missing |
| Semantic conversion to sequential sink | Partial | Measured: RTF writes semantic body text to a bounded forward-only, non-seek UTF-8 sink with configurable paragraph separators; output and sink counters are fully verified outside timing. Production correctness tests, rather than the timing sink, cover partial/interrupted/write-zero behavior. No other-format semantic sink benchmark exists |
| Create a small document | Partial | Fresh DOC/XLS/PPT plus DOCX/PPTX/ODT/ODS/ODP public authoring, and narrow XLSX/RTF forward-only creation with selectable timing evidence; broader streaming authoring remains missing |
| Create or append a very large stream | Partial | XLSX one-sheet scalar rows and plain-run RTF paragraphs have opt-in small/medium/large fixed-window creation cases through non-seek hashing sinks, deterministic untimed reopen, exact counters, and output hashes. Existing-document RTF logical tail append is now covered by both the original pair and four matched Commit-versus-PublicationPlan selectors measured at the pre-staged publication-call interval over tiny/medium/large plain corpora, with a 16 KiB non-seek sink window, exact byte counters, separate planning/publication/reopen/lifecycle scopes, and untimed patch/reopen/source-conflict/failure gates. Planning/publication vectors are per-sample; reopen/lifecycle vectors are one-element preflight-only gates run once outside the sample loop, and the retained candidate snapshot is not claimed to be window-bounded. Large fresh legacy writers still accumulate before final output |
| Exact no-op edit and commit | Covered for generated DOC/XLS/PPT/RTF/XLSX/DOCX/PPTX/ODT/ODS/ODP | Public semantic transaction plus save/reopen; RTF also proves exact raw CP-1252, LZFu and producer-watermark publication, and the dedicated logical-tail no-op case proves shared snapshot identity plus exact sequential bytes; signed/extension corpora remain missing |
| One semantic edit and save | Covered for generated DOC/XLS/PPT/RTF/XLSX/DOCX/PPTX/ODT/ODS/ODP | Cell/paragraph/shape edit or supported ODP slide append, then save/reopen; XLSX now also has six matched opt-in eager/source-backed scalar-cell publication controls over deterministic medium and dense/sparse four-sheet media-rich corpora (one cell, `ceil(1%)`, and exact 256-cell batch), one unmanaged two-worksheet control, and four managed source-backed selectors (one cell, `ceil(1%)`, exact 256-cell batch, and two worksheets). Managed evidence records separate open/plan/commit/publication/reopen vectors plus bounded PartData Budget/cache diagnostics, but has no release ABBA or performance claim. The other timed interval covers open/stage/commit/sequential publication, while lifecycle, reopen, topology, media identity, hashes and source counters are untimed; change 0096 accepts the source-backed provenance-reuse result without making an eager, physical-I/O, allocation or RSS claim. Native XLS also has matched eager/source-backed existing-comment and worksheet-visibility controls over deterministic opaque-heavy CFB corpora. Change 0174 adds matched owned/source-backed ODS existing-cell one-edit controls over the fixed media-rich corpus; change 0177 accepts the one-cell source-backed complete-lifecycle latency result while retaining the 21-cell workload as correctness/phase evidence only. Change 0193 caches the per-owner edit protection parse and accepts the new four-transaction repeated-edit total (9.31%-10.68% lower) and stage-phase (67.87%-71.61% lower) latency on that corpus, leaving single-transaction lifecycles unchanged by design. Change 0171 fuses DOC paragraph, PPT shape-text, and XLS visibility owner validation into the existing final CFB fingerprint fence, removing one complete source scan per effective transaction; only the XLS bounded-batch total and scalar/batch plan-phase latency statistics pass the retained release gates. Native PPT now has a correctness-only immutable-source transaction for one existing equal-encoded-length shape-text atom through a checked CFB byte-range splice; change 0102 resolves its semantic selector through positional metadata and selected-slide reads, while complete artifact fingerprinting/publication remain, so no end-to-end latency or memory claim is made. Source-backed DOC now adds one correctness-only Word97+ main-story paragraph selector/replacement when the paragraph is one uncompressed Unicode piece and replacement width is unchanged; bounded positional chunk scanning, candidate reopen/readback, exact no-op/source/fingerprint/stale checks, inverse and typed partial output are covered, while complete CFB fingerprint/validation/publication scans remain and no DOC end-to-end latency, I/O/range, allocation/RSS, cold/high-latency, real-producer or broad CRUD claim is made (change 0105). ODT separately measures paragraph replacement, inline line-break/run/hyperlink insertion, structural paragraph insertion/removal, and matched existing-image replacement while ODT, ODS and ODP verify exact retained resources and manifest media types; generic packaged ODF chart-definition replacement now uses opt-in verified raw preservation for unchanged members, with correctness-only evidence in change 0101. Source-backed DOCX/PPTX and narrow XLSX calculation-metadata/defined-name/sheet-protection/data-validation/auto-filter/conditional-formatting publication likewise verify eight exact 2 MiB media Parts after one semantic edit. Change 0151 freezes managed source-backed constructors and resource-accounted correctness for eleven XLSX editors—calculation properties, defined names, tab state, print options, page breaks, page margins, page setup, sheet protection, data validation, auto filter, and conditional formatting—with exact no-op/signed/MCE/stale/cancellation/unknown-owner protections, one-byte-under Budget checks, and raw preservation of unselected Parts; it adds no latency, allocation, RSS, copy, decompression, cold-I/O, total-memory, or real-producer claim. RTF now separately measures bounded existing-document logical-tail append, and change 0162 adds selectable phase/correctness evidence for same-length standalone PNG/JPEG payload replacement; their sequential sink, full reopen, durable apply/inverse and foreign-source refusal checks are untimed gates, not speedup claims. Correctness-only additions include source-backed XLSX existing scalar-cell set/clear/remove and row visibility, native XLS cross-workbook Number/Blank scalar transfer, ODS plain-scalar worksheet transfer, ODP dependency-free blank-slide transfer, package-wide DOCX story hyperlink redaction, RTF passive external-reference redaction, XLS worksheet visibility, RTF direct paragraph-layout updates, and ODT plain-paragraph moves. ODS content-validation catalog CRUD and direct PPTX transition edits retain their previously documented narrow closures. A separate RTF correctness gate proves real-producer root-shape edit plus checked LibreOffice resave/readback without presenting it as a paragraph benchmark |
| About 1% semantic update and save | Covered for XLSX/DOCX/PPTX/RTF/ODT/ODS generated corpora; partial for XLSB | Deterministic evenly spaced `ceil(1%)` cell, paragraph and shape changes; the separate XLSB case edits deterministic existing scalar cells on one selected worksheet, then requires reopen/readback, deterministic output, and unchanged unrelated Part digests. It is warm standalone baseline coverage, not ABBA or workbook-wide evidence. The matched XLSX scalar-cell control pairs eager and source-backed selector-first multi-sheet publication over both fixed media-rich shapes, and change 0096 accepts the source-backed provenance-reuse result. DOCX and RTF use one canonical atomic paragraph batch, ODT coalesces ordinary scalar durable replacements internally, and ODS partitions flat cell positions by worksheet into bounded atomic `set_cells` calls; each commits once and reopens the package. Change 0174 adds matched owned/source-backed ODS 21-cell selectors with a common bounded sink and untimed full semantic/media/raw-member gates. Change 0177 withholds the ODS 21-cell latency result because mean and tail stability gates fail; the RTF batch case has a matched scalar-loop comparison recorded in change 0081 |
| Bulk update matching objects | Partial | Selectable: XLSX has a matched eager/source-backed exact-256 existing scalar-cell batch over all four sheets; PPTX has a bounded atomic batch across up to 32 existing slides, with up to 256 unique nonoverlapping shape-text selectors per slide; RTF measures one bounded ordered paragraph batch plus 1/7/63 same-length standalone-picture payload replacements; native XLS measures the exact 256-existing-comment and 64-worksheet-visibility limits through matched eager/source-backed controls; XLSX also replaces a complete three-owner core conditional-formatting collection through matched eager/source-backed controls; ODP measures eight fixed-name existing text boxes across eight slides; ODT replaces 64 fixed existing image owners through matched scalar/bounded-batch controls. Change 0096 accepts only the XLSX source-backed scalar-cell provenance-reuse result; the other evidence additions still make no latency claim pending frozen ABBA. API-only correctness coverage: ODS updates up to 4,096 cells on one selected sheet; native PPT updates up to 256 persisted shape-text targets; XLSX also clears the complete conditional-formatting collection; ODT correctness extends to the exact 256-change embedded object/image bound; ODP correctness extends to the exact 256-owner batch bound. These remaining APIs are not selectable timing evidence |
| Clear/remove/hide/detach/GC distinctions | Partial | Measured: ODT exact paragraph removal intentionally preserves orphaned resources; RTF removes one exact middle ordinary paragraph on its narrow generated plain-source closure and now removes 1/4/32 exact standalone PNG/JPEG picture groups on its separate generated ASCII-hex closure. API-only: OLE2 deletes one or a bounded stream set while retaining storages and unrelated streams; XLS shows/hides/very-hides worksheets, while XLSX additionally shows/hides/very-hides/activates tabs, hides/unhides existing rows, clears scalar cell values without deleting owners, or physically removes supported existing scalar `<c>` owners; ODS clears only unbound validation definitions; DOCX removes direct plain main-document paragraphs; PPTX and ODP remove only dependency-free supported slides; ODT removes selected resource owners while retaining payloads. DOCX reversible hyperlink-wrapper detachment remains distinct from irreversible relationship/URL redaction. Opt-in ODT GC remains exact-source and explicit-name only. General cascading delete, orphan collection and dependency-aware removal remain missing |
| Sanitization and irreversible redaction | Partial | API-only: DOCX retains two deliberately separate exact-source flows. Reversible detachment unwraps selected main-document external `w:hyperlink` owners but retains relationship records and target URLs. The newer forward-only story flow inventories relationship-reachable main, header, footer, footnote, endnote, comments, and glossary stories, selects exact target URLs, unwraps their visible owners, removes the corresponding external hyperlink relationships, and publishes selected story/rels Parts while raw-copying untouched members; it intentionally provides no inverse. Both flows are bounded and fail-closed; the story flow still refuses fields/DDE/mail-merge/templates, MCE, protection, macros/embedded content, signatures, non-story external references, unknown ownership, and encryption, so neither is general personal-data cleanup or package-wide external-reference removal. RTF now inventories only passive top-level `nextfile`/`template` references and can forward-redact authenticated exact source spans; active fields/forms/objects/pictures/shapes, protection, filetable/hyperlink-base/mail-merge, opaque or unknown syntax, and compressed inputs without source spans remain explicit strict refusals |
| Copy object with dependency closure | Partial (bounded PPTX timing; otherwise correctness/counter evidence) | PPTX atomically applies a bounded in-package slide-copy plan for its supported finite owned dependency graph while reusing shared layout dependencies and refusing unsupported/shared ownership. The two owned-source plain/media-rich cross-presentation selectors separate plan, commit, and sequential OPC publication timing. Change 0158 accepts only their canonical generated owned-source prepared-operation result after the additive-topology publisher landed. Change 0159 adds one independent source-backed plain selector over the exact same existing plain corpus; it reports only source-backed plan/publication phases plus separate source/destination logical `ReadAt` call/byte counters, with setup/reopen/raw ZIP/topology/semantic/typed-refusal gates untimed. It is correctness/counter evidence only and makes no eager/source speedup, cache, physical-I/O, allocation/RSS, real-producer, or broader dependency claim. Native XLS transfers bounded cross-workbook rectangles or single cells, but only standalone BIFF8 `Number` and formatting-only `Blank` owners with canonical default XF; ODS transfers one bounded plain-scalar worksheet while retaining its lexical table fragment apart from the destination name; ODP transfers one dependency-free self-closing blank `draw:page` across immutable presentations with deterministic name collision mapping. DOCX copies only direct plain main-document paragraphs. These APIs have source-bound reopen, exact preservation, and reversible-patch tests. Formulas/SST, styles, ranges/merges/validation, arbitrary charts/media/themes/notes, scripts/events/external references, name-collision reconciliation beyond the narrow closures, and complete dependency closure remain missing |
| Merge and split | Partial (selectable correctness evidence) | XLSX now has two opt-in eager commit-plus-bounded-save cases over one deterministic sparse `Sheet1` A1:B2 fixture: merge and unmerge prepare the transaction outside timing and time only commit plus `Workbook::write_to`; untimed checks reopen merge membership, anchor retention, covered/uncovered and unrelated-cell semantics, exact durable patch apply/inverse restoration, and stale-source refusal. RTF adds two opt-in selectors across tiny/medium/large plain lifecycle corpora for ordinary-body split and adjacent merge, with separate open/stage/commit/16-KiB publication/lifecycle vectors; independent exact raw splice, semantic reopen, no-op, volatile/durable forward/inverse, stale/foreign, forged result-artifact, bounded refusal, partial/zero-sink and hash gates remain outside timing, while focused tests retain exact boundary-byte/forged-boundary coverage. These cases make no latency claim without controlled ABBA evidence; broader format merge/split coverage remains missing |
| Patch encode/apply/invert/merge | Partial | Durable encode/decode/apply/inverse coverage now also includes exact DOCX plain-paragraph copy/removal, package-wide DOCX story hyperlink redaction's source-bound forward publication, ODT plain-paragraph move, ODS cross-document scalar-sheet transfer, ODP cross-presentation dependency-free blank-slide transfer and dependency-free slide removal, PPTX dependency-free slide removal, ODS worksheet move/copy, XLS worksheet visibility and cross-workbook Number/Blank scalar transfer, RTF picture-payload/removal edits, the logical-tail append's source-checked durable replay/inverse/reopen/foreign-source gates, and RTF plain paragraph split/adjacent merge with exact boundary bytes and result-artifact SHA-256 digest checks. [Change 0189](changes/0189-xlsx-edit-composition-evidence.md) adds four opt-in XLSX edit-composition selectors for disjoint `Edit::join`, recoverable overlap refusal, disjoint three-way planning, and explicit conflict resolution over one-cell branches; exact no-op, empty-join identity, and `Left`/`Right`/`Neither` are untimed gates for the applicable cases, while durable JSON, forward/inverse replay, stale/foreign refusal, reopen, and media identity additionally gate the save-bearing selectors. Change 0151 adds source-bound managed constructors, direct selected-Part publication, fallible managed-to-owned Arc refusal, and source-bound apply/inverse checks for eleven XLSX editors; it is correctness/resource-accounting evidence without a durable-wire or performance claim. Source-backed XLSX scalar-cell and row-visibility edits, native PPT equal-length shape text and RTF paragraph-layout edits have source-bound apply/inverse but no durable wire claim. The selectable RTF paragraph remove/move and logical-tail cases keep durable work outside timing; split/merge and the cross-document transfer APIs are correctness-only. DOCX irreversible hyperlink redaction and the DOCX story flow intentionally have no inverse. Broader patch merge/join/three-way coverage remains incomplete and unmeasured |
| Validate without mutation | Partial (bounded opt-in evidence) | Opt-in `rtf_validation_report`, `xls_validation_report`, `docx_validation_report`, `pptx_validation_report`, and `odf_validation_report` retain deterministic report hashes, check IDs/statuses, issue codes/counts, source hashes, and bounded source-read counters for ReadAt validators; `docx_section_inventory` additionally retains exact source-backed section topology. Change 0182 accepts only the existing large generated PPTX validation shape after fusing catalog/graph metadata traversal (`2 -> 1` package and `4 -> 1` per-Part relationship-list passes): complete p50 is 7.08%-11.50% lower with distribution/stability gates passing. Tiny/medium PPTX latency, every other validation selector, physical-I/O/resource/cold-cache/producer breadth, decryption, signature verification, macro execution, link fetching, general repair, and broader semantic/security matrices remain open |
| Explicit repair plan | Partial (selectable correctness evidence) | Generic ODF exposes one bounded typed `RepairPlan<NonDestructive>` for an otherwise-valid first/stored `mimetype` with one recognized Extended Timestamp local-header extra. The opt-in `odf_mimetype_repair_plan` case binds source length and SHA-256 evidence, rejects stale/foreign reports and unsafe ZIP/security/semantic cases before output, previews deterministic changed-member/archive effects without source bytes, removes only that local extra, reopens the candidate, proves member digest equality and raw preservation, checks exact forward/inverse recovery and partial progress, and publishes to a zero-retained-output sink. Planning still performs a bounded full-candidate preflight, so no latency or total-memory claim is made; destructive plans and unsupported structural/XML/encryption/signature/macro repairs remain unconstructible. |
| Preserve unknown extension during understood edit | Partial | Targeted OPC raw-copy framing/unknown-member tests, exact untouched opaque ODS rows/members, and exact raw ODT/ODS/ODP auxiliary/media members during neighboring edits; ODT resource replacement retains frame attributes and unknown children while removal retains payload files; ODP model batches retain unselected producer content; the XLSX closures raw-copy unselected Parts and refuse selected MCE/unknown owners. Change 0151 extends this preservation/refusal boundary to managed snapshots and direct selected-Part publication; unknown or unsupported selected owners fail before output, while unrelated members retain raw identity. This is narrow fail-closed preservation, not general extension-aware editing |
| Replace one or a bounded set of low-level Parts, preserve the rest | Covered for owned OPC, bounded source-backed OPC, guarded DOCX main-document and package-wide story-hyperlink semantics, guarded PPTX selected/multi-slide semantics, and guarded XLSX worksheet/workbook semantics | Changes 0008/0021/0022 test owned raw framing, fallback and payload ownership; changes 0037/0077 add consuming one-Part/bounded multi-Part publishers; changes 0039, 0044/0063/0077, 0046, 0061, 0067, 0070, 0073, 0076, 0078, 0079 and 0080 integrate accepted measured guarded semantic transactions while refusing unsafe MCE, signatures, stale closure, topology, relationship, style-reference and printer-reference cases before output. Change 0082 adds matched selectable XLSX conditional-formatting publication through one selected worksheet after workbook/relationship/styles validation, with no latency claim before ABBA. Change 0151 freezes the managed ownership handoff for eleven XLSX source editors, including typed fallible Arc escape and Budget retention/release; the direct publisher materializes only the proven selected Part(s) and raw-copies the rest, with no performance claim. Correctness-only guarded closures add direct standard PPTX transition set/replace/clear through one selected slide Part, XLSX tab-state publication through the workbook alone or the workbook plus old/new active worksheets, DOCX main-document-only hyperlink-wrapper detachment, and package-wide DOCX story hyperlink redaction through selected story XML/.rels overlays with raw preservation of untouched members. General XLSX cell/formula/table-filter, printer graph and structural/inherited/extension-transition PPTX editing remain outside the closure |
| Same-length OLE2 stream edits and metadata moves | Substrate plus bounded native-XLS/PPT/DOC consumers | `litchi-cfb` resolves existing logical streams through validated FAT/MiniFAT chains for bounded same-length whole-stream overlays and byte-range splices, and can move existing stream directory entries without copying payload sectors. Its writer also moves storage subtrees atomically while retaining descendant payload allocations and order. Exact source/version and target checks, overlap/duplicate/path/topology limits, complete reopen, direct sequential publication and atomic path output are covered where exposed. `litchi-ole-common` retains signing/encryption/DRM refusals. Change 0094 proves proportional MiniFAT exact-range reads at the generic substrate; change 0125 adds a distinct 4095-byte MiniFAT boundary control; change 0146 adds correctness/counter evidence for one-shot and repeated public `open_stream` calls, change 0148 adds different-SID A-B-A, public-bulk A-B-A, and overlapping same-target selectors, and change 0149 accepts only the configured-simulator aggregate same-target repeat result after a clean release ABBA. Same-target work changes from `[D,C,0...]` to `[D,D,...]`; local/per-invocation/bulk/concurrent and resource claims remain withheld. Change 0095 adds semantic existing-comment and worksheet-visibility publishers: one/256-comment plans submit 109/27,904 bytes instead of an 80,946-byte Workbook, and one/64-visibility plans submit 1/64 instead of 18,166. Change 0100 adds a native PPT one-shape equal-length text consumer; change 0102 replaces its full metadata-stream selector materialization with bounded Current User, persist-chain, Document-header, SlideList and selected-slide reads. Change 0105 adds a correctness-only Word97+ DOC main-story one-paragraph Unicode-piece splice with positional bounded chunk selection and same-width `WordDocument` replacement; its complete fingerprint and CFB validation/publication scans remain. Balanced XLS ABBA accepts no latency speedup; the PPT and DOC tranches remain correctness/selector-counter only because source identity and publication still read the complete artifact. Failure/retry, ineligible-root, FAT, native semantic, complete resource accounting, allocation/RSS, DOC broader CRUD, broader PPT edits, unstable FAT tails and cold/high-latency evidence remain open |

The XLSB addition is deliberately a separate opt-in binary, so it does not
change the 344-case selector count or the 36-case default matrix. Over one
fixed real POI workbook it covers facade identification, worksheet catalog,
one selected cell, a complete stored-cell scan of the selected worksheet,
facade full text, an exact no-op worksheet patch plus save, one existing-scalar
edit, and `ceil(1%)` existing-scalar edits on that worksheet. Every emitted
case fails closed on its applicable projection/output-reopen/readback,
deterministic output, unrelated-Part semantic-digest, malformed-input,
package/cell read-limit, and sparse-iteration
gates. This is warm initial coverage only: workbook-wide scan/update,
generated shape breadth, cold/process-isolated I/O, allocation/RSS, and ABBA
evidence remain open.

See [change 0148](changes/0148-cfb-same-target-repeat-policy.md) for the
correctness/source-event CFB extension and
[change 0149](changes/0149-cfb-same-target-repeat-release-abba.md) for its
narrow configured-simulator aggregate repeat acceptance. Neither adds semantic
CRUD coverage; 0149 explicitly withholds local, per-invocation, bulk,
concurrent, and resource claims.

The source/output matrix is also incomplete. Owned bytes and instrumented
`ReadAt` exist for OPC/XLSX, and the deterministic range simulator covers
latency/range effects. The bounded source-backed OPC publisher accepts a
forward-only sink and records complete positional input plus bounded output,
but this is not semantic conversion or memory-bounded authoring. RTF and DOCX
final serialization also accept and test a forward-only non-seek sink. RTF
semantic body-text output independently has a selectable bounded forward-only
sink benchmark. Narrow XLSX scalar-row and RTF plain-run creation also have
selectable fixed-window forward-only evidence whose timed sinks retain zero
output bytes. Existing-document RTF logical-tail append now has separate
16 KiB windowed non-seek publication benchmark; its sink retains zero output,
while the API's candidate snapshot remains intentionally outside that sink
window claim. Change 0153 adds matched Commit/PublicationPlan append and exact
no-op controls whose `elapsed_ns` is the pre-staged publication-call interval
around the respective public write call; their planning/publication vectors
are per-sample, while reopen/lifecycle vectors are one-element preflight-only
gates run once outside the sample loop. Retained-byte and failure evidence is
untimed. The
correctness-only same-length OLE2 overlay substrate also
accepts a direct sequential sink and an atomic path destination. Native XLS
and PPT now consume that boundary for narrow semantic edits, but only XLS has
matched release evidence and neither establishes a broad end-to-end
performance claim.
Managed direct `SourceBackedPackage` publication now charges each sink write
against `Resource::OutputBytes` and commits only exact accepted bytes, including
exact/no-op and changed overlays; typed refusal and partial-output behavior are
covered without a performance claim. This excludes `OpcPackage` atomic saves,
`to_bytes`, and unmanaged compatibility sinks.
The new validation tranche uses borrowed generated bytes for RTF and generic ODF
and instrumented positional `ReadAt` for XLS, DOCX, and PPTX; source hashes and
validation topology are checked outside timing, and DOCX section descriptors are
checked exactly.
Borrowed-byte comparisons, filesystem positional cold reads, atomic-save
timing, broad structural PPTX streaming output, and non-seek semantic
conversion for other formats remain.

## Highest-return next cases

1. Change 0165 records a private native-DOC lazy/fused fingerprint proof and a
   bounded descriptive comparison on the existing owner/public phase selector. Each immutable snapshot keeps its
   diagnostic FNV-1a fingerprint lazy, and same-allocation patch replay may
   reuse the exact lineage before falling back to fingerprint plus exact-byte
   checks for an independently reopened source. A clean CPU-2 A1/B1/B2/A2
   release at clean post-rebase revisions `d6818e290` (control) and
   `5dd813b1e` (candidate), with 20 warmups and 500 samples per shape (6,000
   primary samples), records paired lifecycle p50 positive-faster deltas of
   +33.77%/+33.21% for tiny, +12.28%/+13.81% for large, and +17.33%/+17.82%
   for payload-heavy; the immediate fingerprint-demand workflow records
   +14.56%/+13.89%, +4.50%/+5.83%, and +6.55%/+7.08% respectively. Same-lineage
   apply/patch p50/mean/p95 deltas are approximately 99.6%-99.99%, while the deferred
   fingerprint scan remains explicit. DOC guard cases pass and DOC open is
   within the disclosed final guard deltas: noop +78.84%/+79.89% (tiny) and
   +71.08%/+70.40% (large), one-edit +37.23%/+40.81% and +20.45%/+19.79%,
   and open -3.52%/+0.13% and +0.55%/-1.80%. The existing
   `measured_total_ns` lifecycle boundary is unchanged. At change 0165 the
   selector count remained 311; change 0166 raised it to 315, change 0174
   raised it to 319, change 0175 raised it to 320, and change 0180 raises the
   current count to 322, change 0187 raises it to 324, and change 0188 raises
   it to 332; the current XLSX edit-composition tranche raises it to 336,
   while the default remains 36 cases / 198 records. The guard run retains
   24,000 samples.
   This accepts the named private lazy/fused boundary only. A larger shared
   physical/parsed substrate or fused proof across the independent owner and
   public-reader validation layers remains the next DOC mechanism, together
   with broader DOC CRUD, producer, physical-I/O, cold-cache, allocation,
   memory, and publication coverage. No speedup claim is made. Neighboring XLS one-edit/open p50
   regressions stay below 5%; the tiny no-op nanosecond cell is directionally
   noisy. Heaptrack's three-sample, preflight-inclusive whole-process probe
   reports 50,677 allocation calls and a 128.28 MiB peak heap for both sides;
   it is not operation-scoped, and RSS is descriptive only. No physical-I/O, cold-cache,
   real-producer, or generic total-memory claim is made. Change 0161 tested
   the narrow borrowed-input clone removal and rejected it; do not repeat that
   substitution.
   Continue final owner/public-reader work after the accepted native
   XLS editor reuse and fixed-width inventory carry-forward, DOC batched stream
   publication, PPT root-open reuse and
   direct PPT text-edit resolver reuse. Change 0050 removes the newly
   attributed repeated DOC PieceTable physical-range scan but retains the
   complete owner and public-reader validation layers. Change 0051 removes
   repeated adjacent paragraph-style inheritance resolution while retaining
   every direct PAPX, style switch and both readback layers. Change 0053
   removes repeated full CHPX-vector scans from paragraph queries while
   preserving exact run identity/order and both readback layers;
   preserve exact-source patch/inverse and the rejected DOC move as independent
   guardrails.
2. Extend native RTF beyond the now-measured paragraph batch, logical-tail
   append, remove/move and bounded sequential semantic-text cases, and the
   correctness-only direct paragraph-layout and standalone PNG/JPEG
   payload replace/remove closures. The logical-tail evidence is still a
   bounded generated plain corpus with no release ABBA comparison; rich
   formatting/media ownership, malformed/security cases, more producers and
   broader edits remain open.
3. Separate logical authoring/append time from final serialization and reopen
   for DOCX, PPTX and XLSX.
4. Extend the new matched XLSX scalar-cell publication controls to broader
   formula, table/filter, deletion and real-producer semantics. The current
   selectors cover one cell, `ceil(1%)`, and exact-256 set publication over
   deterministic media-rich multi-sheet corpora; change 0096 accepts the
   source-backed provenance-reuse result without an eager, I/O, allocation or
   RSS claim. Scalar-cell clear/physical-remove now have selectable
   phase/counter/correctness evidence (change 0163), while row visibility now
   has matched eager/source-backed correctness and phase evidence (change
   0166). Change 0167 removes the source-backed row publisher's redundant
   semantic reload through an existing lineage/version proof, but unstable
   release drift withholds an end-to-end performance claim. Change 0184 then
   removes the changed candidate's redundant complete scalar-cell parse through
   an exact-source, lifetime-bound validated-store handoff and accepts only the
   stable large-case and selected medium semantic-commit statistics; no
   durable-source-patch claim is made. Change 0185 then removes one
   format-to-OPC selected-Part ownership copy from eligible changed scalar,
   row, metadata, DOCX and PPTX same-topology publishers without adding a new
   CRUD closure. Its matched XLSX result is scenario-scoped; topology-changing
   publication and broad OOXML CRUD remain open.
   Conditional-formatting replacement is selectable
   without an ABBA claim and tab visibility remains unmeasured.
   Direct writer-local action regrouping was measured and rejected
   in change 0030; broader coverage must not present that immaterial prototype
   as a solution.
5. Broaden the accepted narrow XLSX calculation-metadata, defined-name,
   page-break, page-margin/print-options/page-setup, sheet-protection and
   data-validation, auto-filter and conditional-formatting source transactions
   beyond changes 0046/0061/0067/0070/0073/0076/0078/0079/0080/0082 only where
   a complete one-Part semantic closure can be proved; tab-state remains
   correctness-covered but unmeasured, while
   general cell/formula/chains still need a wider publication design.
   Broaden DOCX/PPTX beyond changes 0039/0044 with real producers, MCE-aware
   editing, dependency transfers and explicit topology/signature fallback or
   refusal matrices. PPTX direct standard transition set/replace/clear is now
   correctness-covered through one selected slide Part but remains unmeasured;
   inherited layout/master transitions, sounds and extension forms are refused.
6. Unknown OOXML extension and media preservation during a known semantic edit.
7. Broaden durable PPTX patch coverage beyond direct transitions and the
   narrow dependency-free slide-removal patch into common
   produce/encode/decode/apply/inverse/join/three-way flows.
8. Broaden the supported PPTX slide-copy dependency closure and
   dependency-free PPTX/ODP slide removals across real-producer charts, media,
   themes, notes and collision-name reconciliation; slide split remains
   missing.
9. Extend the new bounded CFB/OPC/generic-ODF reports and the opt-in
   RTF/XLS/DOCX/PPTX validation selectors into broader format-semantic
   DOC/XLS/PPT, OOXML and ODF-family validation/security matrices, signature
   verification and explicit repair planning; the current reports and DOCX
   section inventory are correctness/baseline evidence only.
10. Smart detection versus prepared-source reuse. OOXML smart results retain an
   adoptable parsed OPC package; ODF detection/handoff remains unmeasured.
11. Broaden ODF beyond generated text/grid/deck and accepted compact ODT/ODS/ODP
   unchanged-member publication: formally measure the correctness-covered ODS
   cell batch; broaden the now-selectable matched ODT embedded-resource and ODP
   text-box batches beyond generated fixed-name owners; add source-backed
   selectors/I/O, richer structural families, unknown extensions, real
   producers, richer media and security coverage. Generated ODT 1% paragraph
   updates are covered by change 0045, ODS `ceil(1%)` cell batches now have a
   selectable case without comparative timing evidence, changed-result byte
   finalization by change 0052, content-only line-break publication by change
   0071, existing-style inline-run publication by change 0072, inert hyperlink
   publication by change 0074, and plain paragraph insertion/removal by change
   0075. Plain ODT paragraph move, ODS worksheet move/dependency-free copy and
   plain-scalar sheet transfer, and ODP dependency-free blank-slide
   copy/removal and cross-presentation transfer are correctness-only; new style
   creation, general dependency closure, extension preservation and broader
   structural operations remain. [Change 0104](changes/0104-odt-mixed-model-publication-evidence.md)
   now adds matched medium/large
   mixed model-content ODT publication evidence: 80/320 logical operations
   preserve per-shape output and logical hashes while the measured publication
   count falls from 49/193 scalar publications to one. Its timing excludes
   preparation, reopen/lifecycle/security/limits, I/O, serialization,
   allocation/RSS, and physical cold behavior, so it does not close the wider
   ODF CRUD or resource-evidence gaps.
   iWork is deliberately deferred while the `iwa-*` crates change separately.

The former first item is complete for existing-document ODT transaction
snapshots: change 0014 removes one archive copy and two allocations per
snapshot while retaining the exact no-op, limits, envelope, patch, and readback
contracts. Direct byte ingress and changed publication remain covered by item
11 rather than being implied complete.

The former native DOC/XLS/PPT baseline item is complete in change 0015. Changes
0016 and 0017 accept the first XLS and DOC publication follow-ups without
removing final validation. Change 0018 accepts same-topology ODS row-local
publication and exact untouched opaque-row preservation. Change 0019 accepts
ordinary RTF parser-state work elimination and records a rejected, reverted
ODS package-adoption candidate. Change 0020 accepts RTF ASCII transport
batching and records a rejected, reverted ODT final-document adoption whose
common read guard regressed. None claims broad native-format or ODF CRUD,
real-producer, security, or preservation coverage.

Change 0021 removes the measured 4.19 MiB changed-Part handoff copy from owned
same-topology OPC publication. It does not complete source-backed editing,
unknown-extension semantic corpora, or the broader OOXML CRUD rows above.

Change 0022 removes the separate 4.20 MiB post-validation generated-local-span
copy. The complete generated archive is still validated before output, and the
required compression buffer remains; this is not broader OOXML CRUD coverage.

Change 0023 removes two intermediate strings per block only from ODT full-text
extraction. Structured queries retain their former ownership contract, and
source-backed ODF reads, repeated ODT/ODP scans, other-family/structural
publication, measured ODT/ODP batch cases, broader bulk families and broader
producer/security coverage remain open.

Change 0024 removes one duplicate CFB index open from public PPT root
slide-order capture. It retains independent live-document, slide-order,
review-history and public-reader validation; it does not imply completion of
PPT edit/publication, real-producer, security, or broader OLE2 CRUD coverage.

Change 0025 adds a distinct XLSX commit-plus-first-read attribution case and
reuses commit-time validation only for bounded changed worksheets with exact
part and style/shared-string identity. The dense-wide handoff was rejected on
peak memory and remains a cold-cache fallback; broader XLSX bulk, structural,
source-backed and preservation scenarios remain open.

Change 0026 adds a distinct direct PPT text-edit attribution case and removes
one repeated editor open from target resolution. The fresh commit editor,
exact source comparison, patch/inverse and complete readback remain; this does
not complete PPT anchors, broad edits, real-producer or security coverage.

Change 0027 adds an aggregation-free public ODS cell-sweep attribution case and
a bounded lazy facade locator for repeated queries. It does not add structural,
bulk, source-backed, real-producer, media, security or publication coverage;
those remain under item 11.

Change 0029 broadens the RTF input matrix without adding or renaming timed case
types: deterministic raw CP-1252, deterministic LZFu and a content-addressed
producer watermark join the plain corpus under explicit capability filters.
It also gates the checked `relsize` source/Litchi/LibreOffice semantic chain.
This is coverage evidence, not an optimization result; formatted/media-heavy,
malformed/security and broad real-producer edits remain open.

Change 0030 attributes the existing XLSX 1% update and save cases to the
writer's nested row/cell action regrouping. An owned forward-stream prototype
improved formal p50 by at most 1.61%, reduced process allocation calls only
0.0623%, and left peak heap flat; it and its prototype-only tests were fully
reverted. Distinct bulk clear/remove/hide, structural edits and any larger
semantic-planning/emission coalescing remain open.

Change 0031 adds a deterministic media-rich ODS one-cell edit/save case and
accepts compact `content.xml` publication that raw-copies exact unchanged ZIP
members. It proves exact opaque/media payloads, manifest media types, no-op and
logical fallback behavior. It did not itself add 1%/bulk edits; the later ODS
atomic cell batch and selectable `ceil(1%)` case remain correctness/timing
coverage without a formal before/after performance result. Other-family or
structural publication, real-producer, signature/encryption, source-backed or
broad media semantics remain open.

Change 0034 adds the matching media-rich ODP source-backed text-box edit/save
case and accepts the already shared checked-splice/raw-copy publication path
for content-only rich-object operations. Resource additions and unsupported or
security-sensitive layouts keep the complete rebuild. The case proves exact
media/metadata framing, complete slide and rich-content readback, deterministic
patch/inverse/stale behavior, and material end-to-end latency improvement. It
did not itself add structural/bulk edits; the later correctness-only ODP
text-box model batch remains unmeasured. Repeated queries, real-producer media
and source-backed positional I/O remain open.

Change 0035 adds the corresponding media-rich ODT paragraph edit/save case and
accepts the common checked-splice/raw-copy path only for content-only paragraph
replacement. It proves raw unchanged-member identity, full semantic/media
readback and patch/inverse/stale behavior. Regenerated `content.xml` above the
common 16 MiB optimization limit explicitly returns to the established ODT
rebuild. The later correctness-only ODT embedded-resource batch remains
unmeasured; broader structural edits, repeated queries, real-producer media and
source-backed positional I/O remain open.

Change 0071 extends that same proven content-only publication boundary to the
existing ODT `AppendLineBreak` operation. Its matched media-rich case proves
that only `content.xml` changes, all untouched core/media members remain raw
identical, and the line break survives complete reopen, patch replay, inverse,
stale refusal, and deterministic output. It adds no arbitrary inline-format,
structural, resource, real-producer, or positional-I/O capability.

Change 0072 extends the boundary to the existing ODT `AppendRun` operation.
The packaged regression covers both unstyled and existing-style runs and
proves that only `content.xml` changes while core/media records stay raw
identical. Its matched media-rich case retains complete reopen,
patch/inverse/stale checks and deterministic output. It also isolates exact
no-op commit dispatch from the changed-operation stack frame without adding a
new semantic operation, style-creation API, structural edit, resource edit,
real-producer fixture, or positional-I/O capability.

Change 0074 extends the same boundary to inert ODT hyperlink appends. Complete
text/URL reopen and raw-member proofs remain; no relationship, fetch, or new
security capability is introduced.

Change 0075 covers plain paragraph insertion and removal. It proves exact
paragraph order, raw preservation of every non-`content.xml` member, and exact
patch/inverse behavior. Removal deliberately performs no resource garbage
collection; richer structural/resource edits remain outside this closure.

Change 0036 separates the common OLE2 one-stream edit into editor-open,
candidate-publication, final-render and end-to-end attribution cases over four
unchanged 4 MiB streams. It does not add semantic CRUD coverage. An inline
recapture allocation-reuse prototype was reverted because the complete public
operation improved only 2.61% p50/2.30% mean; the cases remain to gate a
materially different final-publication design.

Change 0037 covers the low-level source-backed replacement of one existing OPC
Part without URI, content-type, relationship or topology changes. The matched
case materializes only the selected Part, raw-copies all other physical
members, preserves exact output identity, and refuses signed real changes or
unsupported layouts before output. It does not add DOCX/PPTX/XLSX semantic
transactions, topology changes, signature policy, real-producer/media matrices
or atomic filesystem publication.

Change 0039 integrates that publisher with exact-source DOCX main-document
transactions over a fixed 16 MiB media-rich package. It materializes only the
main Part, raw-copies every other physical member, fully reopens the result and
retains byte-exact no-op/signature/source-version behavior. MCE-normalized
documents and paragraph transfers are deliberately refused; PPTX/XLSX,
topology changes, real producers, encrypted input and atomic filesystem
publication remain open.

Change 0044 integrates that publisher with an exact-source PPTX selected-slide
transaction over a fixed 200-slide, eight-media package. It materializes the
mandatory presentation root and selected slide, raw-copies the other 227
logical Parts, fully reopens the result, preserves every untouched Part and
media payload, and retains byte-exact no-op/signature/source-version behavior.
MCE-normalized slides, more than one edit operation, topology changes,
real-producer matrices, encrypted input and atomic filesystem publication
remain deliberately outside the contract.

Change 0046 integrates the publisher with an exact-source XLSX
calculation-metadata transaction over a fixed 12-Part, eight-media package. It
materializes only `xl/workbook.xml`, raw-copies the other 11 Parts, fully
reopens the result, verifies the typed calculation state, and preserves every
untouched Part and media payload. MCE projection, changed signed sources,
stale/foreign workbook closures, topology changes and partial sinks retain
typed refusals. Cells, formulas, cached results, styles, shared strings,
relationships and calculation-chain ownership remain outside the capability.

Change 0067 applies the same exact workbook/relationship/worksheet closure to
direct typed page margins. The source-backed editor exposes only set/remove on
one existing normal worksheet, materializes the workbook catalog and selected
worksheet, and raw-copies the other ten Parts. Exact no-ops—including signed
zero lexical variants—preserve the complete archive; changed signed sources,
chartsheets, retargeting, MCE-projected margins, stale/foreign closures, limits
and partial sinks retain typed refusals. General worksheet contents and
topology remain outside the capability.

Change 0070 applies that closure to direct typed worksheet print options. It
materializes the same workbook/selected-worksheet pair, exposes only set/remove
for five boolean flags, and raw-copies the other ten Parts. Exact signed no-ops,
patch/inverse, retargeting, MCE, limits and partial-sink behavior remain checked;
printer-setting relationships and general worksheet contents remain outside the
capability.

Change 0078 applies the complete workbook/worksheet/outbound-relationship
closure to direct typed worksheet protection. It atomically replaces or clears
the complete core and Office 2010 protection metadata, uses the exact
byte-preserving rewriter, and materializes only workbook plus selected
worksheet on the media-rich save. Add/replace/clear/no-op, patch/inverse,
relationships, MCE, chartsheets, signed sources, limits, partial sinks and
strong/legacy verifier metadata remain checked. Password verification, cells,
topology changes and relationship editing remain outside the capability.

Change 0079 applies the same complete closure to typed core and Office 2010
data-validation collections. It preserves inert formulas, quoted lists, UIDs
and references as validated values, replaces only the selected worksheet XML,
and retains exact no-op, patch/inverse, MCE, signature, relationship, version,
media and partial-sink guarantees.

Change 0080 applies a workbook/worksheet/relationship/styles closure to direct
typed worksheet auto-filter and sort state. Add, replace, clear and exact no-op
are distinct; source patches apply and invert exactly; style DXF references
are checked; and protected, MCE-selected, stale/foreign, relationship-mutated
or changed signed sources refuse. Table-owned filters and general worksheet
contents remain outside the capability.

Change 0047 makes the existing ODT one-paragraph benchmark a direct public
indexed query. It retains only the requested structured paragraph while
scanning and validating the complete XML, including malformed trailing input
and all resource limits. It does not provide positional source I/O, early XML
termination, an ODP selector, or broader real-producer/media query coverage.

Change 0049 makes the existing ODP one-slide benchmark a direct public indexed
query. It retains completed semantic content only for the requested slide while
resolving all transition styles and validating content through EOF. It proves
public/full-list parity, out-of-range behavior, late semantic failure and style
inheritance failure. It does not provide positional ZIP/XML I/O, early
termination, repeated-query caching, or broader real-producer/media selection.

Change 0060 leaves ODP CRUD capabilities unchanged while reusing the immutable
editing snapshot's already validated complete slide projection when creating an
isolated transaction. Package/security reopening, settings, declarations, page
metadata, lossless source-page coverage, deterministic commit, exact no-op,
patch/inverse and complete final semantic readback remain. It adds no new edit,
structural/resource, repair, producer, streaming, cold-source or security
capability.

Change 0050 leaves native DOC CRUD and publication semantics unchanged while
indexing the private CLX PieceTable for repeated PAPX/CHPX physical-range
queries. Scalar differential tests cover overlapping fast-save intervals,
ANSI/UTF-16 boundaries and numeric limits; the complete final snapshot and
public DOC reopens remain. This does not add repair, security, real-producer,
streaming-output, or new edit coverage.

Change 0051 likewise leaves native DOC CRUD and publication semantics
unchanged while retaining one parse-local resolved paragraph-style baseline.
Fresh/cached differential coverage includes base/derived inheritance, direct
and piece properties, direct mid-run style switching and cache rekeying. Huge
PAPX, tables/revisions, malformed styles, the complete final snapshot, patch
and inverse verification, and the independent public DOC reopen remain. This
does not add repair, security, real-producer, streaming-output, or new edit
coverage.

Change 0053 leaves the same DOC CRUD and publication surface unchanged. It
uses the private parser-normalized CHPX ordering to binary-search the first
possible overlap and stops after the matching slice. Differential tests cover
empty, reversed, adjacent, gapped, all/no-match and numeric-boundary queries
and preserve exact run identity/order. Formatting cascading, fields/pictures,
comments, glossary parsing, patches, inverse, final snapshot and independent
public DOC readback remain. No new repair, security, producer, streaming or
edit capability is claimed.

Change 0054 leaves the ODS CRUD and publication surface unchanged. It shares
the already retained exact source and target package allocations with the
forward/reverse semantic blob bundles and reuses their BlobIds for existing
operation preconditions. Focused tests preserve deterministic reversible wire,
limit/error precedence, allocation identity and source/target direction; the
complete ODS package/media reopen remains. It adds no patch encode/apply
timing, structural/bulk edit, producer, security or source-backed I/O coverage.

Change 0055 leaves the RTF CRUD and publication surface unchanged. It changes
only the private capacity chosen for the root body style-block vector. The
existing full structural pass supplies the count; source/token/absolute bounds,
fallible allocation, table/deletion fallback, all parser limits, exact no-op,
patch/inverse, candidate parse/readback and transport/producer coverage remain.
It adds no formatting/media edit, repair, security, conversion, cold-source or
real-producer capability.

Change 0056 leaves the native DOC CRUD and publication surface unchanged. It
uses the parser-normalized ordering already present in private CLX pieces and
PAPX runs to replace two repeated linear containment scans with predecessor
binary searches. Scalar differential tests preserve empty/gap, half-open and
numeric-boundary behavior; table filtering, strict SPRM decoding, exact patch
and inverse, final owner validation and independent public DOC readback remain.
It adds no formatting/media edit, repair, security, producer, streaming or new
semantic capability.

Change 0057 leaves the ODS CRUD and publication surface unchanged. Eligible
same-topology worksheet commits retain their exact checked row-range
provenance through raw package emission instead of losing it before the
existing package-preservation gate. Tests preserve exact untouched row,
manifest and media bytes; foreign provenance and unexpected assembly refuse;
signed changed packages retain the established stale-signature stripping
fallback. Compact audit, bounds, complete package/sheet/media readback,
patch/inverse and structural fallback remain. It adds no structural/resource
edit, security capability, real-producer, streaming or cold-source coverage.

Change 0058 leaves the ODS CRUD surface unchanged. It short-circuits only an
exact unified worksheet no-op whose nested commit already proved identical
bytes, and constructs the same durable empty patch without rediscovering
package effects. Changed commits retain every audit, preservation and readback
path. It adds no edit, resource, structural, security, producer, streaming or
cold-source capability.

Change 0059 leaves the native XLS CRUD and publication surface unchanged. It
certifies exact same-family Number/RK/MulRK value ranges, shares untouched
worksheet inventories and carries source offsets forward while retaining the
complete public Workbook validation/readback. Nonnumeric, storage-converting,
structural and resource edits retain the full private parse. It adds no new
cell family, edit, repair, producer, streaming, security or low-level archive
capability.

Change 0040 removes repeated UTF-8 scalar decoding from ordinary RTF text
delimiter discovery without changing the existing CRUD surface. It measures
plain, raw CP-1252 and LZFu opens at medium and large, retains the 25-row
transport/producer smoke, and preserves exact no-op, opaque syntax, limits,
patch/inverse and complete readback contracts. Formatting/media-heavy,
malformed/security and broader real-producer edits remain open.

Change 0043 measures and fully reverts direct RTF decoded-body ownership. It
adds no CRUD surface and keeps the existing parser/model handoff. The retained
test covers a lossy incomplete Shift-JIS body byte plus byte-exact immutable
publication; the broad prototype's plain/LZFu regression and the owned-only
variants' sub-threshold instability are negative evidence, not support claims.

Change 0048 keeps the same generated RTF paragraph-edit CRUD surface but
removes a repeated source clone/lexer from eligible changed commits. The
retained range is proven during the initial full parser preflight; empty,
ambiguous, binary, non-ASCII and LZFu inputs retain the established fallback or
refusal, and every changed candidate still receives complete parse/readback.
The capability-bounded 63-record transport/producer smoke remains green. This
does not add formatted/media editing, conversion, security repair or broader
real-producer mutation coverage.

Change 0041 removes three archive-sized copies from the existing ODT
changed-operation compactness audit without adding a CRUD surface. The
media-rich paragraph edit/save still parses both packages, audits compact XML,
reopens the final document, reads back the paragraph and media, and verifies
patch/inverse and stale-source behavior. Ordinary open/edit/no-op guards remain;
the unchanged sub-microsecond exact no-op segment's +39 ns p50 movement is
disclosed. Structural/resource-adding ODF edits, broader real producers and
repeated ODT/ODP reads remain open.

Change 0042 removes the remaining archive-sized envelope-classification copy
from changed ODT commits without adding a CRUD surface. The media-rich
paragraph edit/save still validates the ZIP, parses the manifest, classifies
encryption/signature state, audits compact XML, reopens the final document,
reads back paragraph/media content, and verifies patch/inverse and stale-source
behavior. Open/no-op/edit guards remain; the untouched large exact no-op
segment's +152 ns p50 movement is disclosed. Broader structural/resource edits,
real producers and repeated ODT/ODP reads remain open.

Change 0052 removes the final archive-sized changed-result copy and one
redundant parse from the same ODT transaction surface. The snapshot remains
byte-only and one independent complete final reopen remains, so this does not
revive the rejected parsed-final-document retention. Exact no-op, compact
audit, package bounds, raw media, patch/inverse, stale-source and
signed/encrypted refusal coverage are unchanged. No new CRUD capability is
claimed; broader structural/resource edits, real producers and repeated
ODT/ODP reads remain open.

Change 0065 leaves the ODP CRUD and publication surface unchanged. An exact
slide-only commit retains its already mandatory parsed slide candidate through
the final validation pipeline and moves it into the published snapshot instead
of parsing the same immutable bytes again. The independent final package
reopen, raw/compact XML audit, embedded-media verification, patch/inverse and
source lineage remain. RDF, chart, design, annotation and rich-content
compound commits retain the established final snapshot parse. It adds no
structural/resource edit, security, producer, streaming, cold-source or
positional-I/O capability.

Change 0066 leaves the RTF CRUD and publication surface unchanged. Explicit
sparse paragraph selection scans the already validated boundary descriptors
without constructing each discarded prefix paragraph, then uses the existing
selected-view location and formatting path exactly once. Differential tests
preserve first/middle/last/out-of-range results, resumed iterator state, empty
paragraphs, line breaks, decoded nonstructural U+000A, fragmented formatting
and trailing text. Parsing, exact no-op/save, edits, patches, variants and
candidate readback are unchanged. It adds no index, cache, structured edit,
format/media capability, security policy, conversion or producer coverage.

Change 0068 leaves the ODS CRUD and publication surface unchanged. The
existing unified worksheet edit moves and shares its exact archive allocation
through the nested worksheet snapshot, patch, private package reader and
candidate validation. Tests prove allocation identity and exact byte/pointer
rollback after a staged closure failure; complete row-splice provenance,
compact/package/media readback, durable patch/inverse, limits and
signed/encrypted policy remain. It adds no structural/resource edit, new patch
wire, security capability, producer, streaming or positional-I/O coverage.

Change 0076 adds a narrow source-backed workbook defined-name transaction. It
replaces or clears the complete direct catalog, validates global and local
sheet scope, publishes only `xl/workbook.xml`, and proves exact patch/inverse,
no-op, media, relationship, topology and calculation-policy preservation.
Protected or MCE/unknown catalogs, invalid local scopes, changed signatures,
stale sources and partial sinks refuse. It does not authorize sheet topology,
formula/cell edits, relationships, name merging or general workbook mutation.

Change 0077 broadens the low-level source-backed OPC publisher to a checked
64-Part replacement set and adds a guarded PPTX batch over at most 32 existing
slides. The measured case edits eight shapes on each of eight slides,
materializes the presentation root plus those slides, raw-copies every
unselected ZIP member, and proves exact patch/inverse, signed all-noop,
stale/foreign/version, limit, topology, relationship, media and partial-sink
behavior. Slide topology, relationships, notes/layouts/charts/media, MCE
projection and changed signed sources remain outside the capability.

The current post-0080 implementation wave adds correctness-reviewed CRUD APIs
for bounded ODS cells, OLE2 stream deletion, RTF paragraph batch/lifecycle,
XLSX conditional formatting and tab state, native PPT shape-text batches, DOCX
external-hyperlink wrapper detachment, XLS existing-comment author/text edits,
ODT embedded resources, ODP text-box models, RTF sequential semantic text and
PPTX direct standard transition set/replace/clear on one selected existing
slide. The PPTX transition closure is correctness-only and unmeasured; it
refuses inherited or extension transitions, sounds and relationship changes.
ODS `ceil(1%)`, RTF `ceil(1%)`, RTF semantic text-to-sink, matched XLSX
conditional-formatting replacement, and matched ODP scalar/batch text-box
replacement are selectable performance cases. Change
0081 records a matched scalar-loop comparison for the RTF batch; change 0082
adds deterministic CF source/sink/materialization/hash evidence but makes no
latency claim before ABBA. No before/after latency claim is made for ODS,
text-to-sink, ODP text-box replacement, or the other API-only additions.

Change 0083 makes the existing native RTF middle-paragraph removal and
first-to-final reorder independently selectable on the identical generated
plain corpus. Both time edit construction, the exact lifecycle staging call,
commit, a constant-size diagnostics assertion, one shared snapshot-handle
clone, and bounded sequential serialization; complete reopen/full projection,
volatile and durable forward/inverse, stale-source refusal, move no-op identity,
sink counters and output hashes remain outside timing. Changed CP-1252, LZFu,
producer-watermark and opaque/formatted inputs retain fail-closed refusal and
exact source bytes; equal-position moves remain exact no-ops. This is harness
correctness and baseline coverage only: no comparative latency, allocation,
memory or materialization improvement is claimed.

Change 0084 makes existing ODP scalar and bounded cross-slide text-box model
replacement independently selectable over one fixed 12-slide/eight-media
corpus. Both update the same eight existing fixed names without renames in one
transaction/commit and produce the same complete semantic projection. The
batch raw-preserves the manifest while repeated scalar staging regenerates it,
so case-specific physical digests are retained. Complete presentation and
rich-content reopen, volatile/durable forward/inverse, stale refusal, raw
auxiliary/media identity and bounded sink counters remain untimed. Owned ODP
exposes no positional-source or logical-Part materialization diagnostics, so
none are fabricated. This is selectable evidence only, with no latency,
allocation, memory, or materialization claim before frozen CPU-pinned ABBA.

Change 0085 makes existing ODT scalar and bounded embedded-resource
replacement independently selectable over one fixed 200-paragraph/eight-media
corpus extended with 64 existing package-backed image owners. Both replace the
corresponding fixed same-length target paths without owner
insertion/removal/reorder in one transaction/commit
and must produce one identical complete semantic resource projection; physical
byte identity is not required. Complete document/image reopen, frame/path/media
types, all payload digests, untouched raw members, volatile/durable
forward/inverse, stale refusal, deterministic case hashes and exact sink
counters remain outside timing. Owned ODT exposes no positional-source or
logical-Part materialization diagnostics, so none are fabricated. This is
selectable evidence only, with no latency, allocation, memory, I/O, or
materialization claim before frozen CPU-pinned ABBA.

The current XLSX merge/split tranche adds two independent opt-in lifecycle cases over a
deterministic sparse A1:B2 fixture. The eager merge and eager unmerge paths
prepare their inputs and transaction edits outside timing, then measure only
semantic commit plus bounded sequential save. Reopen semantics, anchor and
covered/uncovered-cell behavior, unrelated-cell preservation, exact durable
apply/inverse restoration, and stale-source refusal remain untimed. This is
selectable correctness evidence only; no latency claim is made without
controlled ABBA evidence.

The current native XLS comment tranche makes the existing exact-owner API
selectable through four opt-in cases on one deterministic CFB artifact: eager
and source-backed publication for one middle comment and for the exact 256-edit
bound. The corpus retains an untouched worksheet/comment plus eight exact 2 MiB
incompressible streams and an opaque metadata stream. All replacements preserve
record length and compressed encoding width. The report records distinct
semantic staging/plan and publication samples, complete bounded sink counters,
per-case output hashes, and source-backed changed-comment/stream/span counts,
equal Workbook lengths, exact NOTE/TXO splice/replacement-byte diagnostics, and
source/target fingerprints. Untimed gates
reopen all comments, compare every untouched stream, prove eager patch replay,
inverse and stale refusal, preserve the explicit eager fallback for a
length-changing update, and refuse protected edits. Generic source counters are
limited and labeled to the explicit owned-source ingress because this public XLS
API does not accept a caller-provided `ReadAt`. Change 0095 accepts only the
exact replacement-byte reduction; its balanced release ABBA establishes no
speedup, and allocation, RSS and source-I/O remain open.

Change 0135 adds four opt-in native XLS fixed-width numeric CRUD selectors.
The Number pair edits `Untouched!E21` from 42 to 43; the RK/MulRK pair edits a
standalone RK and both cells of one MulRK record in one transaction, preserving
each source storage family. Eager and source-backed timing separates edit,
`set_number`/`set_numeric`, commit, and complete sequential publication. The
deterministic CFB corpora retain opaque siblings and metadata. Untimed gates
cover source ingress, complete Snapshot/Workbook reopen and numeric readback,
deterministic digests, untouched stream bytes/topology, source-backed equal
Workbook lengths, patch apply/inverse/stale, exact no-op/fingerprint,
signed/macro/protected/unsupported refusal, and an untimed 54016.xls
real-producer reopen/inverse gate. Source-backed evidence reports complete
target materialization on both paths and is explicitly not positional-I/O,
bounded-memory, allocation/RSS, speedup, or broad-producer coverage.

Change 0137 adds two opt-in plan-only native XLS numeric CRUD selectors over
the same Number and RK/MulRK corpora. Commit timing includes validated
overlay-plan construction and composed semantic validation; complete
publication remains a separate `write_to` interval. The forward-only plan
retains no complete target snapshot or target byte vector and records false
target-retention/materialization flags plus zero commit-boundary target bytes,
while sink bytes prove complete output. Exact source/target fingerprint
preflights, no-op, full forward reopen/readback, untouched topology/member
identity, security/unsupported, partial-sink and 54016.xls producer gates stay
outside timing; ordinary source-backed patch/inverse/stale semantics remain
separate. Composed validation may allocate/read a candidate Workbook model, so
the zero target-artifact field is not a bounded total-memory measurement. This
adds correctness/descriptive coverage only; no plan-only performance claim is
made before balanced release ABBA.

Change 0091 adds four opt-in native XLS worksheet-visibility cases over one
deterministic CFB corpus with 66 worksheets, eight 256 KiB incompressible opaque
streams, and opaque metadata: eager and source-backed one-owner edits plus
eager and source-backed exact-64-owner batches. The one-owner selectors target
worksheet position 1; the batch selectors hide positions 1 through 64 and
leave positions 0 and 65 visible. The timed interval separates
transaction/commit planning from sequential publication through a bounded
64 KiB sink. Untimed gates reopen every worksheet, compare the complete CFB
stream catalog and every opaque stream, verify the exact one-byte `hsState`
offset set, prove eager patch replay/inverse, source-backed fingerprints and
changed-span counts, exact no-op identity, the 64-owner cap refusal, and
protected-source refusal. Source counters are explicitly limited to owned
source ingress, as with the comment tranche. Change 0095 submits 1/64 exact
visibility bytes rather than one 18,166-byte Workbook replacement. The
complete candidate snapshot remains; balanced ABBA accepts no speedup and
allocation, peak-memory and I/O evidence remain open.

The 2026-08-12 non-iWork wave from `cb797b382` through `f6bbdf19c` adds
correctness-only bounded validation reports for CFB, OPC and generic ODF;
exact CFB stream/storage moves; reversible DOCX plain-paragraph copy/removal
and irreversible external-hyperlink
relationship/URL redaction; PPTX supported-closure slide copy and
dependency-free removal; ODP dependency-free blank-slide copy/removal; ODS
worksheet move and dependency-free copy; source-backed XLSX scalar-cell
set/clear/remove and row visibility; XLS worksheet visibility; ODT plain
paragraph move; and RTF paragraph layout plus standalone picture
replace/removal. These commits added no selectable case to the then-167-case harness
and make no latency, allocation, memory, I/O or materialization claim. Large
streaming authoring, general dependency closure, broader merge/split, repair,
format-semantic validation, broad real-producer/security evidence and matching
benchmarks remain open. The matrix also records the immediately preceding
`cbf581c91` same-length CFB stream-splice substrate under the same
correctness-only restraint.

Production commit `c4c52c7ec` separately adds correctness-only ODS
content-validation catalog CRUD. Its supported closure is clone-staged
add/set/update/same-name replace/remove/clear/rollback, exact semantic no-op,
source-checked reversible patch, unified `content.xml` publication, untouched
member preservation and full reopen. Duplicate names, unsafe rename,
referenced removal/clear, unrepaired dangling references on changed commit,
opaque/MCE/DTD owners, operation/output limits and changed signed packages
refuse atomically. This closure is unmeasured and is not evidence for general
ODS cell, formula, style, or structural CRUD performance.

Production commit `82dfcf26a` separately adds a validated same-length OLE2
overlay substrate. Existing logical streams are resolved through FAT or
MiniFAT chains, bounded physical spans are derived without caller offsets, the
composed artifact and selected streams are reopened before output, and exact
source/target fingerprints and versions remain checked. Direct sequential
sinks report typed partial progress; path output uses synced sibling staging
and atomic rename; the common owner retains signed/encrypted/DRM refusals.
Length/topology changes, duplicates, overlaps and bounds fail closed. No
DOC/XLS/PPT end-to-end integration or speed claim has been adopted, and this
generic CFB substrate is not semantic format coverage.

Production commit `626540e22` separately adds correctness-only, opt-in ODT
resource GC. It plans and applies source-exact reversible deletion only for
explicitly named detached package files or subdocuments. Ordinary owner
removal still preserves displaced payloads. Signed, encrypted, protected, and
unknown-reference packages refuse; this is not selectable performance evidence
and makes no latency, allocation, memory, or I/O claim.

Change 0090 adds two opt-in native RTF logical-tail cases to the harness:
`rtf_logical_tail_append` and `rtf_logical_tail_noop_save`. They operate on an
existing, exact-source, default-formatted plain RTF rather than the separate
streaming-creation writer. Tiny/medium/large corpora append 4/64/256 bounded
one-run paragraphs under explicit input, inserted-byte, output, and durable
patch limits. Timed work includes staging, candidate commit validation, and
sequential publication through a fixed 16 KiB non-seek hashing window; the
sink retains zero output bytes and exposes source/input/inserted/output,
paragraph/run, write-count, and largest-write counters. Untimed gates prove
exact empty-append snapshot identity, complete paragraph/text reopen, exact
sequential bytes, in-memory patch replay/inverse, durable JSON
encode/decode/apply/inverse, and foreign-source conflict refusal. This is
coverage/baseline evidence only: no release ABBA, allocation, RSS, or speedup
claim is made, and the sink window is not a claim that the append transaction's
validated candidate snapshot is memory-bounded.

Change 0093 adds six opt-in matched XLSX scalar-cell publication selectors:
eager and source-backed one-cell, deterministic `ceil(1%)`, and exact-256
existing-cell batches. Each shape is a fixed four-sheet medium or dense/sparse
scalar grid with eight untouched deterministic media Parts. The timed interval
covers the sum of open/selector-first staging/commit and sequential publication
segments through a bounded sink; source cache-diagnostic sampling is excluded
from that sum. Untimed gates reopen the result, compare semantic cells and
package topology/relationships, verify raw media identity and exact hashes; the
source-backed pair additionally checks raw local/central ZIP identity for every
unselected member, while eager checks retain semantic/topology/media-payload
identity. The harness also exercises source-backed exact no-op, clear, and
physical-remove behavior. Source read and successful materialization counters
are reported outside timing.

Change 0109 (`2f70f08dc`) adds an opt-in managed source-backed tranche for the same committed
XLSX value-only editor: one cell, deterministic `ceil(1%)`, exact-256, and a
bounded two-worksheet two-cell transaction, plus an unmanaged source-backed
two-worksheet control. These controls use matched finite cache policies and
report separate open/plan/commit/publication/reopen vectors, source read bytes,
payload materializations, cache/Budget diagnostics, output and canonical
semantic hashes, exact untouched raw-member fingerprints, and release-to-zero
gates. The managed Budget covers retained and in-flight OPC `PartData` payloads
only; parsed stores, metadata, staging, rewritten candidates and output buffers
remain outside that accounting. No controlled release ABBA has been run for
this tranche, so it makes no speedup, allocation, RSS, hardware, cold-I/O,
decompression or real-producer claim.

[Change 0150](changes/0150-xlsx-managed-cell-values-budget-evidence.md)
extends those same four managed controls with separate cumulative
`InputBytes`, accepted `OutputBytes`, declared cold-load `Work`, retained
`Objects`, and `Memory` diagnostics around publication. Its untimed
one-byte-under `OutputBytes` replay requires a typed refusal before any sink
byte is accepted and verifies final managed object and memory release. This is
resource-accounting evidence only; it does not turn the excluded parsed,
staging, rewritten, or output allocations into a total-memory claim.

The current HEAD correctness wave (`1803a6ca6`, `8366f9df5`, `823b4a39e`,
`19c4d08b9`, `d77870da1`, and `292dbb459`) adds no selectable performance case
and makes no latency, allocation, RSS, I/O, or throughput claim. Native XLS
adds `cell_values::Transaction::copy_scalar_cell_from` and
`copy_scalar_cells_from` (with `copy_cells_from` as an alias) for bounded
cross-workbook rectangles or single cells. The closure accepts only standalone
BIFF8 `Number` and formatting-only `Blank` owners with canonical default XF;
occupied target family/width mismatches, formulas/SST/strings, styles,
drawings/relationships, unknown or dependency-bearing owners, protected or
stale input, and finite-bound violations refuse atomically. ODS adds
`document::Edit::transfer_plain_scalar_sheet_from`, which transfers one
bounded dependency-free worksheet containing only plain string/number/
boolean/date/time cells, bounded repetitions, and direct text paragraphs while
retaining the source table fragment lexically apart from its destination name.
Formulas, styles, ranges, merges, validation, drawings/charts/images,
scripts/events, external references, unknown/MCE XML, security-sensitive
packages, and namespace ambiguity remain refusals. ODP adds
`Transaction::transfer_dependency_free_blank_slide_from`, which transfers one
dependency-free self-closing donor `draw:page` across immutable presentations,
deterministically remaps the page name, and retains source-bound durable
forward/inverse and untouched-member checks; page bodies, layouts/masters,
identifiers/navigation, links, animation, macros, protection, signatures,
encryption, unsafe package parts, namespace rebinding, and self-closing
destination presentation roots remain refused. DOCX adds a separate
package-wide story-hyperlink inventory/redaction over relationship-reachable
main, header, footer, footnote, endnote, comments, and glossary stories. It
selects exact target URLs, unwraps ordinary visible `w:hyperlink` owners,
removes their external relationships, publishes changed story XML/.rels Parts
while raw-copying untouched members, and is bounded forward-only (no inverse);
fields/DDE/mail-merge/templates, MCE, protection, macros/embedded content,
signatures, non-story external references, unknown owners, and encryption
remain fail-closed diagnostics/refusals. RTF adds a bounded inventory and
forward redaction for passive top-level `nextfile` and `template` references
using authenticated exact source spans. Active fields/forms/objects/
pictures/shapes, protection, filetable/hyperlink-base/mail-merge, opaque or
unknown syntax (including skipped active-story destinations), and compressed
inputs without source spans remain explicit strict refusals.
Change 0096 retains a CPU-pinned balanced release ABBA comparison for those six
source-backed selectors. The p50 geomean improves 21.66%/22.65% and p95 improves
21.38%/22.70% in the two directions, with exact output hashes. Physical source
read and materialization counters are unchanged, so the accepted result is the
removal of a redundant semantic worksheet reload/reparse, not an I/O,
allocation, RSS or cold-filesystem claim. The default 36 cases / 198 records
remain unchanged.

Change 0106 adds correctness-only RTF ordinary-body paragraph split and
adjacent merge. The closure is literal ASCII, root-level and source-proven:
split inserts one canonical `\\par ` at a checked UTF-8 boundary, merge removes
only the selected exact `\\par` boundary, and a final paragraph cannot be split
at its end without an existing exact boundary. Candidate reopen/readback,
sequential publication, durable forward/inverse replay, exact boundary-byte
restoration, result-artifact SHA-256 digest verification, foreign/stale refusal, finite
limits, protected-source failure atomicity and active/external/unknown syntax
refusals are covered by six focused tests. This is CRUD correctness evidence
only; no latency, I/O, allocation/RSS, cold/high-latency, stream-window,
producer-breadth or rich-RTF claim is made. Signed-document verification and
preservation remain outside the RTF operation's proof.

Change 0097 separately measures fresh RTF streaming creation after bounded
escape-free ASCII batching. The p50 geomean improves 76.41%/76.47% and p95
75.23%/75.76%; the large case reduces sink calls from 7,208,970 to 1,441,802
under a hard 32-byte request ceiling, with exact bytes and hashes. This does not
claim existing-document edit, allocation, RSS or cold-I/O improvement. iWork
remains deliberately deferred while the `iwa-*` crates change separately.

Change 0116 adds eight opt-in native PPT `Pictures` selectors to the then-current
216-name matrix: matched eager/source-backed presentation-open cases, cold
all-`images()` query cases, repeated-query cases, and directly timed fresh
open-plus-all-images cases. The deterministic corpus has
eight slides, 32 distinct PNG records, and 256 KiB per record. The corpus
manifest records the exact `Pictures` stream length and SHA-256; per-result
source evidence records the canonical ordered image semantic SHA-256. Both
implementations use the same finite `RecordLimits`. Untimed exact and
one-byte-under package/`Pictures` gates cover
both paths. Source evidence reports `ReadAt` calls/bytes and overlap with the
contiguous `Pictures` payload window in this generated fixture: open must show
zero overlap, first query must show the full stream, and the repeated
source-backed query must show zero additional overlap. These are fixture-scoped
source-read observations from independent untimed replays, not per-sample
timing counters, a general CFB sector map, or internal materialization counts.
Source-backed elapsed samples instead use uninstrumented immutable
`litchi_core::OwnedSource`. No latency, allocation, RSS, cold-I/O, or release
ABBA claim is made by change 0116.

Change 0119 adds three opt-in native PPT selected-shape selectors, bringing the
current matrix to 219 names while leaving the default 36 cases / 198 records
unchanged. The existing eager query-only control is paired with a positional
source-backed query, and two new cases pair fresh eager/source-backed
open-plus-query phases over the same generated target. Timed source-backed
samples use an uninstrumented immutable source; independent untimed replays
record deterministic logical `ReadAt` calls/bytes once per measured sample,
excluding warmups, and retain one canonical selected-text digest. The umbrella
`Presentation::open` path also adopts the positional native-PPT package on
Unix/Windows after content-derived OLE2 routing, retaining DOC-before-PPT
polyglot precedence and the existing non-PPT fallback. This is correctness and
logical-read evidence only: there is no latency, allocation/RSS, physical-I/O,
cold-filesystem, or release-ABBA claim.

Change 0120 adds eight opt-in PPTX ordinary-root filesystem selectors, bringing
the current matrix to 227 names while preserving the default 36 cases / 198
records. The matched eager/source-path controls cover open, full owned
slide-list materialization, slide-count, and selector-first slide 100 on the
fixed 200-slide/eight-text-box/eight-2 MiB-media corpus. Untimed source replays
classify compressed payload ranges exactly (catalog-only open/count,
target-slide-only selection, and all-slides/no-media listing), while full
metadata, size, names, text hashes and parity remain outside timing. This is
correctness/logical-read coverage only; no latency or resource claim is made.

Change 0121 adds two opt-in native PPT repeated selected-shape query controls,
bringing the current matrix to 229 names while preserving the default 36 cases
/ 198 records. The eager and source-backed controls retain one prepared owner
and issue eight identical queries; source timing is uninstrumented and a
separate replay records calls, bytes, prior-covered range bytes, and a
canonical semantic digest. The production two-query regression binds
74 calls / 8,310 bytes for legacy CFB reconstruction and 66 calls / 3,190
bytes with a retained parsed CFB index. These are logical-I/O/correctness
figures only, with no latency or resource claim.

Change 0122 adds four opt-in matched ODP media-rich read controls, bringing the
current matrix to 233 names while preserving the default 36 cases / 198
records: eager/source-backed open and eager/source-backed one-middle-slide
query over the existing 12-slide/eight-2 MiB `Pictures/` corpus. Source timing
uses an uninstrumented `OwnedSource`; one independent untimed replay records
exact source calls, bytes, coalesced prior-range overlap, and compressed
Pictures overlap (`pictures_read_compressed_range_bytes`) for the named phase,
separate from prior-read overlap (`source_read_range_overlap_bytes`). A second
replay reads one explicit selected media member and gates all-Pictures overlap
to that member's complete compressed range while recording bytes outside
Pictures separately. Compressed ZIP range totals and uncompressed payload
bytes/digests use distinct evidence fields. The open phase may expose bounded
ZIP-tail range overlap with the final Pictures member; this is physical request
evidence and not a payload-materialization claim. Full eager/source semantic
parity and deterministic slide/media digests—including the eager one-slide
post-timing media check—are checked outside timing. No latency, decompression,
allocation, RSS, physical-I/O, or release-ABBA claim is made.

Change 0123 adds four opt-in unified-root ODP filesystem controls, bringing
the current matrix to 237 names while preserving the default 36 cases / 198
records: eager byte-backed and filesystem source-backed open plus matched
middle-slide queries over the existing media-rich corpus. Corpus creation and
temporary-file publication stay outside timing; open controls time only root
owner construction and query controls time only an already-open root query.
Untimed gates compare complete root semantic/metadata projections, source-file
archive/member/hash identity, and selected media payloads. Source controls
also run independent direct `SourceBackedPresentation` instrumented replays
to prove catalog/query media laziness and exact selected compressed-range
coverage. Production routing tests bind the root filesystem handoff. No
latency, physical-I/O, decompression, allocation, RSS, or release-ABBA claim
is made.

Change 0124 adds six opt-in ODS filesystem/source controls, bringing the
selectable matrix to 243 names while preserving the default 36 cases / 198
records: eager/source-backed unified-root open, plus typed ODS eager/source
selected-cell and selected-media queries over the existing two-sheet,
eight-`Pictures/`-member media-rich corpus. Corpus and temporary-file
publication stay outside timing; eager clones stay outside open timing; typed
owner construction stays outside selected-query timing. Every sample checks
root worksheet names/count/text, complete cell parity, metadata, exact file
bytes/hash/archive/member/media identity and typed ODS parity. Independent
instrumented `SourceBackedSpreadsheet` replays report positional calls/ranges:
open reads catalog/content without unrelated media, selected-cell queries add
zero reads after retained content preparation, and selected-media evidence
pairs an all-Pictures replay with a selected-range-only replay; both cover
exactly the selected compressed member range, excluding other media ranges.
Compressed ranges and uncompressed payloads are separate evidence fields, and
eager source vectors are empty.
This is correctness/logical-range evidence only; it makes no latency,
physical-I/O, decompression, allocation, RSS or release-ABBA claim.

Change 0125 adds two matched CFB MiniFAT boundary selectors, bringing the
selectable matrix to 245 names while preserving the default 36 cases / 198
records. The deterministic 4095-byte target occupies 64 logical 64-byte
mini-sectors. Legacy full-stream and `SharedOleFile::read_stream_range`
controls record separate open/read/total timing, exact source calls/bytes/range
sizes, returned length, and payload hash; the focused gate expects legacy
root-mini-stream amplification and one exact 4095-byte positional request.
This is generic CFB correctness/request-amplification evidence only, not
semantic DOC/XLS/PPT CRUD or a latency/resource claim.

Change 0126 adds eight ordinary-root DOCX filesystem/source controls, bringing
the selectable matrix to 253 names while preserving the default 36 cases / 198
records: eager/source open, paragraph-count, full paragraph listing, and full
text. The fixed source-edit corpus remains the unchanged 200-paragraph,
eight-incompressible-2 MiB-media archive. Open timing compares `fs::read` plus
`Document::from_bytes` with `Document::open(path)`; query roots are prepared
outside timing and only the exact root query is timed. Untimed eager/source
parity covers paragraph/full-text/table/element/metadata projections. Exact
source SHA plus logical OPC part/relationship/content-type/blob-hash gates
cover package preservation, including all media hashes and source
immutability. Independent typed source replays record calls, bytes,
request sizes, compressed-range coverage and materializations, classifying
zero payload overlap at open, complete main-document range coverage during
query-selector preparation, and zero main/media/unselected/core overlap during
the query.
This is correctness/logical-range evidence only, with no latency,
physical-I/O, decompression, allocation, RSS, cold-cache, ABBA, broad-security,
or Markdown-performance claim.

Change 0127 adds matched eager/source-backed ODS repeated-cell sweep selectors
over the existing two-sheet 32 by 32 media-rich corpus. Four identical sweeps
cross the adaptive locator threshold; source preparation counters and zero-read
post-preparation replay evidence are retained per measured sample. Semantic
digest/count and complete source/member/media preservation checks remain
untimed. The selectors bring the selectable matrix to 255 names while the
default 36 cases / 198 records remain unchanged; no latency or resource claim
is made pending a clean release ABBA run.

Change 0134 adds matched eager/source-backed ODS ordered cell-batch sweep
selectors over the same two-sheet 32 by 32 corpus. The timed scope prepares
owners and 2,048 borrowed selectors first, then performs four bounded
`cell_batch` calls (8,192 result slots) with black-boxed results. Independent
source replay records exactly eight version observations and zero
post-preparation payload reads per sweep, with ordered parity and digest/count
gates outside timing. At that stage the selectable matrix was 257 while the
default 36-case / 198-record tranche was unchanged; this is correctness/logical-read
evidence only, not a release speed or resource claim. See
[`0134`](changes/0134-ods-source-cell-batch-sweep-evidence.md).

Change 0135 brings the current selectable matrix to 261 names with four
matched native XLS Number and RK/MulRK publication selectors, preserving the
default 36 cases / 198 records. The Number control edits `Untouched!E21`
42 -> 43; the packed control edits one standalone RK and both cells in one
MulRK record transactionally. Complete target materialization, source ingress,
full Snapshot/Workbook reopen, family/value readback, untouched CFB topology and
member bytes, equal source-backed Workbook lengths, patch/inverse/stale,
no-op/fingerprint, security/unsupported refusals, deterministic sink evidence,
and the untimed 54016.xls real-producer reopen/inverse gate are explicit
outside-timer checks. This adds correctness coverage only and claims no
positional I/O, allocation/RSS, bounded artifact memory, speedup, or broad
producer support.

Change 0137 brings the current selectable matrix to 263 names with two
additional plan-only native XLS Number and RK/MulRK publication selectors. The
timed commit vector includes plan construction and composed semantic
validation; complete `write_to` publication remains a separate timed vector.
The production plan retains no target snapshot or complete target byte vector,
so evidence explicitly records false target-retention/materialization flags,
zero commit-boundary target materialization, and unsupported patch/inverse
semantics while sink bytes prove complete publication. Forward reopen,
topology/member identity, no-op and exact source/target fingerprint
preflights, partial-sink, security/unsupported and 54016.xls producer gates
remain outside timing. Composed semantic validation may allocate/read a
candidate Workbook model, so zero target-artifact bytes at commit is not a
bounded total-memory claim. This
is correctness/descriptive evidence only and makes no latency, memory,
allocation, RSS, I/O, or speedup claim before balanced release ABBA. See
[`0137`](changes/0137-xls-numeric-plan-only-publication.md).

Change 0138 supplies the balanced release evidence for the two plan-only
selectors. Number and RK/MulRK each run strict one-process `A1, B1, B2, A2`
legs on CPU 2 with 20 warmups and 200 samples; complete-operation p50, p95,
p99 and mean all improve in both paired directions. Number p50 improves
27.57%/28.58% and RK/MulRK 24.90%/24.56%. Matched process-level VmHWM
captures (three warmups/30 samples) agree for Number (-10.73%/-10.66%) but
disagree for RK/MulRK. Valid heaptrack profiles report whole-process
allocation totals and unchanged peak heap in each sampled A/B pair. Exact
output digests, complete sink bytes, zero plan-only target-artifact fields,
forward reopen/readback, topology, security, no-op, partial-sink and
real-producer gates remain green. This is an accepted latency result for
these deterministic fixed-width families only; it does not claim operation-
only allocation, bounded total memory, physical I/O, cold-cache behavior or
broad producer coverage. See
[`0138`](changes/0138-xls-numeric-plan-only-release-abba.md).

Change 0139 adds two opt-in source-backed ODP repeated-text selectors, taking
the current selectable matrix from 263 to 265 while preserving the default 36
cases / 198 records. Both use the same 12-slide, eight-picture corpus and time
four full-text projections on a prepared owner. The control reproduces the
historical uncached public sequence; the candidate exercises
`SourceBackedPresentation::text()` and its threshold-two cache. Untimed gates
bind text parity, archive/media identity, zero post-preparation source reads,
and exact freshness vectors. This is correctness and logical replay evidence
only, with no latency, physical-I/O, decompression, allocation, RSS,
cold-cache, ABBA, or release claim. See
[`0139`](changes/0139-odp-repeated-text-cache-evidence.md).

Change 0140 accepts the corresponding clean-revision CPU-2 matched release
result for this exact four-call source-backed projection shape. Paired p50
improves 45.80%/46.32% and p95 improves 45.25%/45.83%; p99 and mean agree as
well. Whole-process Heaptrack allocation-call and temporary-allocation counts
fall 14.31% and 17.25%, while peak heap and process VmHWM remain neutral. The
result does not broaden CRUD coverage and makes no single-call, open,
physical-I/O, decompression, cold-cache, peak-memory, operation-local
allocated-byte, or generic ODF claim. See
[`0140`](changes/0140-odp-repeated-text-cache-release-abba.md).

Change 0144 adds six opt-in CFB simulated-range selective-read selectors,
taking the current selectable matrix from 265 to 271 while preserving the
default 36 cases / 198 records. The clean CPU-2 release ABBA accepts only the
configured simulator's MiniFAT request/byte/service-floor reductions and
matching p50/p95 direction; the 4 MiB FAT exact-work control stays near neutral.
This is generic OLE2 substrate evidence and does not add DOC/XLS/PPT semantic
CRUD coverage or claim real cold/network/device, allocation, or RSS behavior.
See [`0144`](changes/0144-cfb-simulated-range-source-evidence.md).

Change 0145 adds two opt-in PPTX cross-presentation slide-copy selectors,
bringing the selectable matrix from 271 to 273 at that point while preserving the
default 36 cases / 198 records. Plain and media-rich deterministic source and
destination packages report plan, commit, and sequential OPC publication
phases separately, with complete semantic/package/closure/collision,
source-immutability, durable-patch, stale/foreign, and refusal gates outside
timing. This is correctness and sink-counter evidence only; it makes no
speedup, allocation, RSS, release-ABBA, or physical-I/O claim at the 0145
revision. Change 0158 later accepts the canonical generated owned-source
prepared-operation result: total p50 improves 29.643%/26.196% for plain and
43.294%/43.604% for media-rich, while media-rich publication p50 improves
49.321%/49.680% in the two ABBA directions. Source-backed, physical-I/O,
real-producer, broader dependency, and generic PPTX claims remain open. See
[`0145`](changes/0145-pptx-cross-slide-copy-evidence.md) and
[`0158`](changes/0158-pptx-additive-topology-release-abba.md).

Change 0146 adds twelve opt-in generic CFB MiniFAT `open_stream` selectors,
bringing the selectable matrix from 273 to 285 while preserving the default
36 cases / 198 records. They cover 36-byte and 4,095-byte targets, one-shot,
repeat-3, and sequential repeat-8 operations on the two bounded sibling shapes,
with local and deterministic-range-model source evidence. This is
correctness/counter coverage only; no release latency, allocation, RSS,
physical-I/O, native DOC/XLS/PPT, cross-format, or iWork claim is made. See
[`0146`](changes/0146-cfb-open-stream-evidence.md).

Change 0147 adds no CRUD surface or selector. Its clean release ABBA accepts
only the configured simulator's one-shot result and exact proportional source
work. The repeated many-small tradeoff remains visible and no generic repeat,
local wall-clock, resource, physical-I/O, native semantic, cross-format, or
iWork claim is added. See
[`0147`](changes/0147-cfb-open-stream-release-abba.md).

Change 0149 also adds no CRUD surface or selector. Its clean 28,800-sample
release ABBA accepts only aggregate repeat-3/repeat-8 totals under the named
configured range simulator after same-target work changes from `[D,C,0...]` to
`[D,D,...]`. One-shot controls are neutral, and noisy local/per-invocation/
bulk/concurrent distributions remain withheld together with allocation/RSS,
physical-I/O, cold/network/device, native-format, cross-format, and iWork
claims. See
[`0149`](changes/0149-cfb-same-target-repeat-release-abba.md).

Change 0152 adds no CRUD surface or runtime selector. The final same-target
MiniFAT single-flight revision (`c270c8f3b` plus `f46381c6f`) was compared with
clean control `e486e4b1` on CPU 2 using 20 warmups, 500 samples, and 24 records
per leg (48,000 retained samples). All correctness/source-event invariants
passed; existing concurrent scenarios recorded 6,473 candidate versus 8,000
control logical source calls (19.09% fewer). At that revision the 291-name
matrix was unchanged; change 0153 adds four RTF selectors measured at the
pre-staged publication-call interval,
making that matrix 295. Change 0154 adds six ODF content-COW publication
selectors, making that matrix 301; change 0159 made it 302, change 0160 made
it 303, change 0162 made it 305, change 0163 made it 309, change 0164 made it
311, change 0166 made it 315, change 0174 made it 319, and change 0175 made the then-current matrix 320. Only `cfg(test)` source-event acceptance and
tests changed in 0152. This is
source-event/correctness evidence only. Root MiniStream cache and
resource-accounting boundaries and broader performance gaps remain; local or
generic latency, allocation/RSS/peak memory, physical I/O/syscalls,
cold-cache/device/network, decompression, native semantic, OOXML, ODF, RTF,
and iWork claims are withheld. See the
[`0152` release record](changes/0152-cfb-same-target-singleflight-release-abba.md)
and [summary](results/cfb-singleflight-abba-0152-summary.json).

Change 0154 adds no new production CRUD capability; it measures the committed
generic source-positional ODF `content.xml` publisher against the owned rebuild
for matched semantic ODT/ODS/ODP edits. The clean CPU-2 A/B/B/A result accepts
96.35%-96.63% p50 improvement in both pair directions at the prepared
in-memory publication boundary. Exact content, semantic reopen, member
inventory, positional untouched-member raw identity plus physical/central
order, no-op, limits, cancellation, source immutability, and logical `ReadAt`
replay remain untimed gates. End-to-end edit/save, allocation/RSS,
physical-I/O, decompression, cold-cache, filesystem, real-producer, structural
or resource-adding ODF CRUD, and iWork claims remain open. See
[`0154`](changes/0154-odf-content-cow-publication-evidence.md) and the
[summary](results/odf-content-cow-abba-0154-summary.json).

Change 0159 adds one opt-in source-backed PPTX plain cross-presentation
slide-copy selector, taking the selectable matrix from 301 to 302 at that
revision
while preserving the default 36 cases / 198 records. It reuses the exact
plain source/destination corpus from the owned-source selector and calls the
public source-backed plan/publication API. Only `plan_ns` and
`publication_ns` are reported; setup, reopen, raw ZIP member/order/comment,
semantic/topology, source-version, and typed stale/foreign refusal gates are
untimed. Separate source/destination logical `ReadAt` call/byte counters are
recorded, but cache, eager/source speedup, allocation/RSS, physical-I/O,
media-rich, real-producer, and release-ABBA claims remain explicitly open.
See [`0159`](changes/0159-pptx-source-backed-cross-copy-evidence.md).

Change 0160 adds the opt-in `doc_owner_public_phases` attribution selector over
the exact tiny, large, and payload-heavy native DOC writer corpora. It emits
per-sample strict-owner, complete public-reader, exact-source retention,
authoring, in-memory owner-render, patch, outer-operation, output-materializing,
and checked unattributed intervals. Exact semantic readback, no-op identity,
patch/inverse/stale behavior, malformed/typed refusal, hashes, and untouched
CFB streams are untimed gates. Successful event order/cardinality is checked
after each named outer interval but before the lifecycle timer stops, so that
recorder validation remains visible in checked unattributed time. Separate
format tests bind balanced error events. The production feature emits ordered,
content-free events and owns no clock. A clean CPU-2 release run at revision
`ab333008d3` retained 800 samples per shape across four fresh processes:
lifecycle p50 was 0.081/1.157/44.227 ms and combined complete public-reader
validation p50 was 0.016/0.598/20.721 ms for tiny/large/payload-heavy. This
accepts only the exact attribution distribution; synchronous observer overhead
and the two tiny subphase mean spread triggers are disclosed, and no
speedup, physical-I/O, allocation/RSS, cold-cache, real-producer, or
optimization claim is accepted. See
[`0160`](changes/0160-doc-owner-public-phase-attribution.md) and the
[summary](results/doc-owner-public-phases-0160-summary.json).

Change 0162 adds selectable phase/correctness evidence for the already
committed RTF standalone-picture payload replacement and group-removal APIs.
The two opt-in selectors use one deterministic ASCII/uncompressed root-level
PNG/JPEG corpus per tiny/medium/large shape, an independent exact source splice,
and complete semantic/no-op/volatile-patch/durable-patch/stale/foreign/refusal/
partial-sink gates. Replacement preserves mixed-case hexadecimal transport,
whitespace and at least one unselected group; removal preserves every retained
group and surrounding byte exactly. At change 0162 the harness had 305
selectable names;
the historical default remains 36 cases / 198 records. This is not insertion,
image rendering, compressed/binary/nested/shape picture support, real-producer
coverage or a release performance result. See
[`0162`](changes/0162-rtf-picture-crud-evidence.md).

Change 0163 adds four opt-in XLSX eager/source-backed scalar-cell clear/remove
selectors over the existing medium and dense/sparse numeric four-sheet corpus.
They target one existing `Sheet1!A1` owner: clear retains an empty `<c>` owner,
while remove deletes it. Open, planning/staging, commit, sequential
publication, lifecycle, generic logical source/materialization counters, and
source raw-unselected preservation are recorded in their narrow scopes. The
default remains 36 cases / 198 records; at change 0163 the matrix had 309
selectable names.
This is debug correctness/phase/counter evidence only, with no latency,
allocation/RSS, physical-I/O, cold-cache, decompression, durable-source-patch,
or real-producer claim. See
[`0163`](changes/0163-xlsx-cell-clear-remove-evidence.md).

Change 0164 adds two opt-in RTF ordinary-paragraph selectors across the exact
plain lifecycle corpus at tiny, medium, and large shapes. Split inserts one
canonical five-byte `\\par ` boundary and merge removes one authenticated
adjacent boundary, yielding independent raw splice deltas of +5 and -5 bytes.
The `SourceSummary.rtf_paragraph_split_merge` record carries selected and
adjacent positions, split offset, source/expected/publication hashes, separate
open/stage/commit/publication/lifecycle vectors, and the semantic, raw,
no-op, volatile/durable, stale/foreign, forged-result, bounded-refusal,
partial/zero-sink and hash gates. The public publication sink is a fixed
16-KiB window retaining zero output bytes; this is correctness/phase evidence
only, with no latency, resource, physical-I/O, cold-cache, source-backed,
real-producer, or rich-RTF claim. The focused harness gate is
`rtf_paragraph_split_merge_selectors_are_opt_in_bounded_and_gate_complete`;
the historical default remains 36 cases / 198 records. That revision had 311
selectable names; change 0166 made it 315, change 0174 made it 319, and change
0175 made the then-current matrix 320. See
[`0164`](changes/0164-rtf-paragraph-split-merge-evidence.md).

Change 0165 records a private native-DOC lazy/fused fingerprint implementation,
a bounded descriptive comparison, and extends the existing `doc_owner_public_phases` record without adding a
selector. Each immutable snapshot defers its diagnostic FNV-1a fingerprint
until demanded; same-allocation patch replay is an explicit post-lifecycle
workflow extension, while an independently reopened source still takes the
fingerprint plus exact-byte authority path. The record adds independent
expected source/target fingerprints, per-sample source/target fingerprint
vectors, `same_lineage_apply_ns`, `deferred_fingerprint_ns`,
`workflow_no_diagnostic_ns`, `workflow_with_fingerprint_demand_ns`, and four
gates for same-lineage apply, reopened-source apply, independent fingerprints,
and workflow arithmetic. The historical `measured_total_ns` lifecycle
boundary is unchanged. At that revision the selectable matrix remained 311
names; change 0166 made it 315, change 0174 made it 319, and change 0175 made the then-current matrix 320,
while the default remains 36
cases / 198 records.

The final clean post-rebase CPU-2 A1/B1/B2/A2 release uses control revision
`d6818e290` and candidate revision `5dd813b1e`, with 20 warmups and 500 samples
per shape (6,000 primary samples) plus 24,000 guard samples. Paired lifecycle
p50 deltas are +33.77%/+33.21% (tiny), +12.28%/+13.81% (large), and
+17.33%/+17.82% (payload-heavy); the immediate fingerprint-demand workflow
is +14.56%/+13.89%, +4.50%/+5.83%, and +6.55%/+7.08% respectively.
Same-lineage apply/patch p50/mean/p95 deltas are approximately 99.6%-99.99%, and the
deferred scan is reported explicitly. Final DOC guard p50 deltas are noop
+78.84%/+79.89% (tiny) and +71.08%/+70.40% (large), one-edit
+37.23%/+40.81% and +20.45%/+19.79%, while DOC open is -3.52%/+0.13% and
+0.55%/-1.80%. Neighboring XLS one-edit/open p50 results are mostly neutral or
improved; its tiny no-op nanosecond cell remains directionally noisy.
Heaptrack's three-sample, preflight-inclusive whole-process probe reports 50,677
allocation calls and a 128.28 MiB peak heap for both sides; it is not
operation-scoped, and RSS is descriptive only. No speedup, physical-I/O,
cold-cache, real-producer,
or generic total-memory claim is made, and the non-`const` fingerprint accessor
is a capability change, not a deprecation. See
[`0165`](changes/0165-doc-lazy-fingerprint.md), the
[summary](results/doc-lazy-fingerprint-0165-summary.json), and the
[release manifest](results/doc-lazy-fingerprint-0165-manifest.json).

Change 0166 adds four opt-in XLSX existing-row visibility lifecycle selectors:
matched eager/source-backed one-row hide and exact-256-row unhide controls.
They use one-sheet media-rich `medium` (512 × 16) and `large` (2,048 × 32)
corpora with eight untouched 512-KiB media members. Open, stage/plan, commit,
publication, and lifecycle vectors are separate; source-backed records expose
only logical owned-source `ReadAt` and pre-publication cache diagnostics.
Measured publication is bound by exact length/SHA-256 to an untimed
semantically reopened expected artifact. Raw untouched-member identity is
common; exact no-op, foreign/stale,
signed/protected/formula/MCE/macro/relationship, and partial/zero-sink refusal
fields are source-backed-only and omitted from eager records. This raises the selectable matrix
from 311 to 315 while preserving the default 36 cases / 198 records, and adds
correctness/phase evidence only—no speedup, latency, allocation/RSS,
physical-I/O, cold-cache, decompression, or real-producer claim. See
[`0166`](changes/0166-xlsx-row-visibility-evidence.md).

Change 0167 keeps the same four selectors and matrix counts while removing one
redundant publication-time worksheet semantic reload, cell parse and row scan
from the source-backed row publisher. A >8 MiB read-trap regression proves the
mandatory OPC selected-member read remains and a second semantic read does not.
Clean 20-warmup/500-sample CPU-2 A/B/B/A records observe 50.42%-68.23%
publication reductions across all paired p50/mean/p95/p99 values, but the 5%
same-implementation gate fails: maximum absolute drift is 34.80% for control
large/unhide publication p99 and 10.23% for candidate medium/hide complete-
workflow p50; first-pair medium hide/unhide complete-workflow p99 regresses
6.95%/2.69%. The implementation and raw evidence are
retained without an acceptance-grade end-to-end latency, allocation/RSS,
physical-I/O, cold-cache, decompression, or producer claim. See
[`0167`](changes/0167-xlsx-row-visibility-provenance-reuse.md).

Change 0184 keeps the same four row-visibility selectors and matrix counts
while replacing the changed candidate's complete scalar-cell reparse with a
private, lifetime-bound rewrite token that may reuse only the exact snapshot
source's validated cell store. Row markup is still freshly scanned, worksheet
XML is still fully validated, and generic scalar-cell rewrites retain their
full parse. Clean 20-warmup/500-sample CPU-2 A/B/B/A records accept both large
semantic-commit directions and all large/unhide complete-workflow statistics;
medium/unhide semantic-commit p50/p99 and large/hide complete-workflow
mean/p95/p99 also pass the paired drift gates. Other statistics are withheld
where same-implementation drift exceeds the preregistered threshold. This is
correctness-equivalent work elimination only: it does not extend CRUD coverage
beyond existing explicit row owners and makes no allocation/RSS, physical-I/O,
cold-cache, decompression, independent-publication, or producer claim. See
[`0184`](changes/0184-xlsx-row-visibility-store-reuse.md).

Change 0185 likewise adds no selector or CRUD closure and leaves the matrix at
322. It introduces additive shared-Arc OPC overlay methods and migrates
eligible DOCX, PPTX, and XLSX source publishers, removing one complete
selected-Part `Arc -> Vec -> Arc` ownership copy on changed same-topology
publication. Exact no-ops continue byte-copying the source through an empty
overlay plan; managed Arc escape, selected-member validation, signatures,
limits, source fences, compression, reopen and partial-sink behavior remain.
Clean matched XLSX evidence accepts only the statistics listed in the change
record; dense multi-update and most row cases are withheld. No new formula,
structural, topology-changing, producer, allocation/RSS, physical-I/O,
decompression, cold-cache, broad OOXML, or iWork coverage follows. See
[`0185`](changes/0185-opc-shared-source-overlay.md).

Change 0168 keeps the same native XLS Number/RK/MulRK plan-only selectors and
matrix counts while fusing BIFF semantic target validation into the common
CFB planner's existing composed-view fingerprint bracket. Two redundant
post-plan complete source scans are removed per effective edit; no-op plans
still skip semantic target validation, and CFB reopen/range checks, source
preconditions, final source/target fingerprints, format security policy,
numeric readback, partial-output typing, and publication remain. Clean CPU-2
20-warmup/500-sample A/B/B/A records observe lower complete-workflow and
semantic-commit p50/mean/p95/p99 values in both paired directions, but the 5%
same-implementation gate fails at 10.56% control and 9.81% candidate maximum
drift. The correctness-equivalent work elimination is retained without an
acceptance-grade latency, allocation/RSS, physical-I/O, cold-cache,
decompression, or producer claim. This does not extend CRUD coverage beyond
existing fixed-width existing-cell numeric replacement. See
[`0168`](changes/0168-xls-numeric-validation-fusion.md).

Change 0172 keeps the same 315-case matrix and the same two native XLS
plan-only fixed-width numeric selectors. An explicit owned-byte CFB ingress
now preserves the immutable snapshot proof into direct sequential publication,
which removes two redundant complete fingerprint scans while retaining the
64 KiB emission pass and source/target hashes. Generic positional sources,
composed views and atomic saves remain fully fenced. Clean CPU-2
20-warmup/500-sample A/B/B/A records accept 37.54%-39.00% lower complete-
workflow statistics and 64.44%-66.76% lower direct-publication statistics
through p95 for Number and RK/MulRK; Number p99 also passes. RK/MulRK
publication p99 and all resource/I/O/producer claims are withheld. This is a
performance result for the existing fixed-width existing-cell replacement
closure, not broader formula/string/structural CRUD coverage. See
[`0172`](changes/0172-cfb-owned-numeric-publication.md).

Change 0173 keeps the same 315-case matrix and the existing four native XLS
comment selectors. Effective source-backed comment transactions now reuse the
planner's exact composed candidate for semantic readback and preserve immutable
snapshot provenance into direct sequential publication. This removes three
complete artifact scans while retaining emission hashing, exact no-op,
fixed-length/encoding-width refusals, semantic reopen, fingerprints,
partial-sink typing, and atomic-save fences. Clean CPU-2
20-warmup/500-sample A/B/B/A records accept scalar complete p50/mean/p99,
scalar semantic/publication, and bounded-batch semantic improvements. Scalar
p95 and batch complete/publication are withheld by the drift policy. This is
performance evidence for existing NOTE/TXO updates only; add/remove, shape
topology, and length-changing comment lifecycle CRUD remain uncovered. See
[`0173`](changes/0173-cfb-comment-publication-fusion.md).

Change 0169 keeps the 315-case matrix and the existing
`xlsx_streaming_create` selector unchanged. It removes transient hierarchical-
budget ancestor-vector allocations and retains four reservation nodes inline,
with deeper caller-defined hierarchies spilling. Clean CPU-2 A/B/B/A records
accept medium/large p50/mean/p95/p99 and tiny p50/mean/p95 reductions in both
paired directions; tiny p99 regresses 1.81%/2.75% and is withheld. Matched
whole-process Heaptrack allocation calls fall 48.81% and temporary allocations
69.77%, peak heap is unchanged, and RSS directions disagree. Exact one-sheet
archive/worksheet hashes, scalar-cell counts, zero retained output, and the
4 KiB authoring window remain fixed. This adds performance evidence, not CRUD
coverage: multi-sheet, shared-string/style/formula/date, physical/cold I/O,
total-memory, and producer claims remain open. See
[`0169`](changes/0169-xlsx-streaming-budget-charge.md).

Change 0170 also keeps the 315-case matrix and existing selector unchanged.
It batches ordinary UTF-8 between XML entity boundaries, skips redundant scalar
counting when byte length proves the character bound, and reuses one formatted
row number. Clean CPU-2 A/B/B/A accepts large p50/mean/p95/p99, medium
p50/mean/p95, and tiny p50 improvements while preserving exact archive/
worksheet hashes, sink topology, zero retained output, and the 4 KiB window.
Tiny mean/tails and medium p99 are withheld. This remains performance evidence,
not new CRUD or richer-authoring coverage. See
[`0170`](changes/0170-xlsx-streaming-escape-runs.md).

Change 0174 raises the selectable matrix from 315 to 319 while preserving the
default 36 cases / 198 records. It adds a bounded source-backed ODS
existing-cell transaction and matched owned/source-backed one-cell and 21-cell
selectors over the fixed media-rich corpus. The source-backed path materializes
no complete candidate artifact, rewrites only eligible physical rows, raw-copies
untouched ZIP members, and publishes through a 16 KiB zero-retention sink.
Exact no-op, semantic patch/inverse, full reopen/digest, media payload, raw
member, source/output hash, limit, stale, protection/signature, lossy-row and
partial-sink gates are covered. Standard table metadata outside row spans is
preserved; formulas, merges, repeated physical rows, style retargeting,
unknown values in rewritten rows, nested/multiple paragraph markup,
encryption and changed signatures refuse. This is correctness and matched
timing coverage only; release latency, allocation/RSS, physical-I/O,
cold-cache, real-producer, durable-ZIP-patch and atomic-save claims remain
open. See [`0174`](changes/0174-ods-source-backed-cell-edit-publication.md).

Change 0175 adds one opt-in immutable-owned CFB atomic-save selector, raising
the matrix from 319 to 320 while preserving the default 36 cases / 198
records. It exercises the existing low-level same-length OLE2 overlay scenario
and does not add format-semantic CRUD coverage. The retained production path
removes two redundant complete fingerprints only for sealed `Arc<[u8]>`
sources; generic positional sources remain fully fenced. Exact output,
semantic reopen, untouched-stream bytes, source/output hashes, one changed
span and atomic destination publication are gated. The exact logical-work
reduction is accepted; latency
is withheld because control drift exceeded 5%.

Change 0176 records two reverted experiments and adds no coverage. ODS
authenticated content reuse regressed both source-backed cell workflows, and
XLSX conditional-formatting readback reuse was directionally inconsistent.
Their existing CRUD closures remain as documented before the experiments.

Change 0177 added no selector name and left the then-current matrix at 320. It
hardens the existing four ODS source-cell records with aligned lifecycle and
phase vectors, backend-applicable gates, and a separately untimed logical
source replay. Clean release A/B/B/A accepts the fixed one-existing-cell
source-backed lifecycle at 75.03%/74.27% lower p50 than eager ownership. The
21-cell deterministic 1% result is retained as correctness/phase evidence only
because mean and tail stability gates fail. Structural cells/rows, formulas,
merges, insert/delete, real-producer, allocation/RSS and physical-I/O coverage
remain open.

Change 0183 also adds no selector or CRUD closure and leaves the current matrix
at 322. It repeats only the previously withheld deterministic ODS 21-existing-
cell workload on clean current HEAD. Complete source-backed lifecycle p50 is
72.07%/72.61% lower than eager ownership and p50/mean/p95/p99 stability gates
all pass, so the fixed generated 1% warm-latency claim is now accepted. The
result does not cover formulas, merges, structural rows, insert/delete,
allocation/RSS, physical I/O, cold cache, real producers, durable ZIP patch,
atomic save, or broad ODS CRUD.

Change 0178 also left the then-current matrix at 320 and adds no CRUD closure. It removes
one redundant final complete fingerprint only from plans rooted in sealed
owned CFB bytes, after candidate reopen and optional format-owner validation.
Generic positional sources retain their final mutation fence. The existing XLS
comment and fixed-width numeric selectors prove exact output and paired lower
p50 directions, but workload stability gates withhold latency acceptance. The
accepted scope is one deterministic logical scan/digest-pair reduction per
effective owned plan; comment add/remove, length-changing edits, formulas,
structural BIFF changes, resource/physical-I/O/cold-cache and producer coverage
remain open. See
[`0178`](changes/0178-cfb-owned-planning-fingerprint.md).

Change 0179 likewise adds no selector or CRUD closure and leaves the matrix at
320. It reuses the source-backed PPTX editor catalog already validated at open,
removing two complete 200-slide graph builds from one/same-slide edits and nine
from the eight-slide batch. Payload materializations, logical source reads,
output hashes, topology, signatures, MCE and stale/foreign refusal remain
unchanged. Clean release paired latency directions disagree and stability
gates fail, so only the exact `3 -> 1` / `10 -> 1` catalog-build reduction is
accepted. General topology edits, relationships, real producers, allocation,
RSS, physical I/O, cold-cache and broader PPTX coverage remain open. See
[`0179`](changes/0179-pptx-source-catalog-reuse.md).

Change 0180 raises the matrix from 320 to 322 with matched uncached-control and
public cached source-backed ODT repeated-text selectors. It adds no CRUD
closure: the full-text row was already covered. The exact four-call generated
workload reduces complete XML/block-model projection phases from four to two,
retains at most one 16 MiB string, returns four fresh owned strings, and proves
semantic/archive/media/range/freshness parity plus zero post-preparation source
reads. Two clean balanced cycles accept p50/mean only; tails, allocation/RSS,
physical I/O, single-call/open, producer, generic ODF, non-text projection, and
broad CRUD coverage remain open. See
[`0180`](changes/0180-odt-source-text-cache.md).

Change 0181 adds no selector or CRUD closure and leaves the matrix at 322. The
existing native-XLS plan-only fixed-width numeric path now reuses three private
source policy facts captured from its immutable complete Workbook validation,
removing one repeated source semantic reopen while retaining independent target
validation and every CFB/publication fence. Clean Number total and commit
p50/mean/p95/p99 pass; RK/MulRK latency, publication, resource/I/O, cold-cache,
atomic-save, real-producer performance and broader formula/string/structural
XLS coverage remain open. See
[`0181`](changes/0181-xls-source-policy-reuse.md).

Change 0192 adds no selector or CRUD closure and leaves the matrix unchanged.
It repeats only the withheld change 0191 ODT open-only workload on clean
current HEAD with a bit-identical release binary and accepts warm open-only
p50 (51.68%/52.62% lower than eager byte ownership) and p99 (51.83%/59.56%
lower); mean and p95 remain withheld on eager drift. See
[`0192`](changes/0192-odt-open-only-rerun-evidence.md).

Change 0193 adds one opt-in selector (`ods_source_backed_repeated_edit`) and
raises the matrix to 341 names while leaving the default 36 cases / 198
records unchanged. It computes the ODS source-backed edit protection parse at
most once per owner through a private success-only `OnceLock`, removing three
of four identical protection-domain parses in the new four-transaction
repeated-edit workload. Frozen cross-binary CPU-2 A/B/B/A accepts the
four-transaction total at 9.31%-10.68% lower and the stage phase at
67.87%-71.61% lower across p50/mean/p95/p99; commit/publication phases are
neutral-withheld, single-transaction lifecycles are unchanged by design, and
no allocation/RSS, physical-I/O, cold-cache, producer, or broad ODF claim is
made. See [`0193`](changes/0193-ods-source-edit-protection-cache.md).

Change 0194 adds no selector or CRUD closure and leaves the matrix unchanged.
It rewrites the litchi-ods worksheet `validate_text` forbidden-character check
from a per-`char` scan to a per-byte scan (exactly equivalent: every forbidden
character is ASCII and encodes as one identical byte). Frozen cross-binary
CPU-2 A/B/B/A over the existing three ODS source-backed edit selectors accepts
the four-transaction repeated-edit total p50/mean (2.08%/0.63% and
2.77%/0.19% lower) and commit-phase p50/mean/p99 (up to 13.11% lower), the
one-cell guardrail commit p50 (0.96%/2.60% lower), and the one-percent
guardrail lifecycle p50/mean plus commit p50/mean/p95 (up to 3.45% lower);
all other statistics are withheld. See
[`0194`](changes/0194-ods-validate-text-byte-scan.md).

Change 0195 adds no selector or CRUD closure and leaves the matrix
unchanged. It caches the row-local edit layout scan of `content.xml` at most
once per source-backed ODS owner through a private success-only `OnceLock`,
removing three of four full-document scans in the four-transaction
repeated-edit workload. Frozen cross-binary CPU-2 A/B/B/A accepts the
four-transaction total at 5.55%-12.49% lower and the commit phase at
25.13%-36.98% lower across p50/mean/p95/p99, plus the one-cell guardrail
commit p50/mean/p95/p99 and the one-percent guardrail lifecycle p99 and
commit p50/mean/p99; all other statistics are withheld as neutral. See
[`0195`](changes/0195-ods-source-content-layout-cache.md).

Change 0196 adds no selector or CRUD closure and leaves the matrix
unchanged. It replaces the Aho-Corasick automata behind
`litchi_core::xml::{escape_xml, unescape_xml}` with exactly equivalent byte
scans (fuzz-verified byte-identical over 2M randomized cases) and drops the
dependency from litchi-core. Frozen cross-binary CPU-2 A/B/B/A accepts the
one-percent ODS lifecycle p50/mean (0.44%-3.02% lower) and commit
p50/mean/p99 (1.53%-20.15% lower); the repeated-edit and one-edit selectors
are neutral-withheld throughout. See
[`0196`](changes/0196-xml-escape-byte-scan.md).

Change 0197 adds no selector or CRUD closure and leaves the matrix
unchanged. It explored batching the ODS per-row synthetic-document reparse
into one reparse per changed window; the one-percent commit phase accepted
p50/mean/p95 in two independent runs, but the lifecycle scope measured a
consistent adverse pattern in both runs and the change was withheld and
reverted. Two multi-row window regression tests added under it remain. See
[`0197`](changes/0197-ods-batched-row-window-reparse.md).

Change 0198 adds no selector or CRUD closure and leaves the matrix
unchanged. It extends the 0195 ODS content-layout cache with the derived
table/row topology so cached-layout commits skip the per-commit span-vector
re-scans. Frozen cross-binary CPU-2 A/B/B/A accepts the four-transaction
repeated-edit total p50/mean/p95/p99 (0.19%-5.77% lower) and commit
p50/mean/p95/p99 (2.20%-7.72% lower), plus guardrail subsets listed in the
change doc; all other statistics are withheld as neutral. See
[`0198`](changes/0198-ods-content-layout-topology-cache.md).

Change 0199 adds no selector or CRUD closure and leaves the matrix
unchanged. It removes the per-event `Event::into_owned()` deep copies from
the two full-document `NsReader` parse loops in litchi-ods
(`worksheet::codec::parse` and `settings::codec::locate`), a byte-exact
borrow-lifetime elision. Frozen cross-binary CPU-2 A/B/B/A accepts the
source-backed and eager opens on all four statistics (6.42%-9.70% and
9.46%-15.17% lower), both guardrails' lifecycle and commit p50/mean/p95,
and the repeated-edit total, commit, and publication p50/mean/p95/p99 plus
stage p50/mean/p95; the remaining tails are withheld as neutral. See
[`0199`](changes/0199-ods-parse-event-copy-elision.md).

Change 0200 adds no selector or CRUD closure and leaves the matrix
unchanged. It fuses the three litchi-ods-owned open passes over
`content.xml` into one shared tokenization with pass-ordered error
selection, while the standalone shells keep the original inline loops.
Frozen cross-binary CPU-2 A/B/B/A accepts the source-backed open
p50/mean/p95/p99 (19.72%-24.55% lower), the one-percent lifecycle all-four
(3.26%-5.72% lower), the one-edit lifecycle p50/mean (2.81%-5.38% lower),
the eager-open p99, and repeated-edit total mean/p95/p99, stage p99, and
publication mean/p95/p99; a documented sub-1% repeated-edit
total-p50/publication-p50 reading on source-identical phases is recorded as
code-layout wobble. See
[`0200`](changes/0200-ods-fused-open-parse.md).

Change 0201 adds no selector or CRUD closure and leaves the matrix
unchanged. It folds the ODS `content.xml` structural validation pass into
the 0200 fused open tokenization (two-phase driver; standalone validator
unchanged for its other call sites). Frozen cross-binary CPU-2 A/B/B/A
accepts the source-backed open p50/mean/p95/p99 (9.32%-17.88% lower), the
one-edit lifecycle all-four plus commit p50/mean, the one-percent lifecycle
p95, and the repeated-edit total p95/p99 plus publication all-four; the
eager-open primary adverse reading did not reproduce in the single permitted
rerun, and a sub-1.5% repeated-edit commit p50/mean reading on
source-identical phases is documented as code-layout wobble. See
[`0201`](changes/0201-ods-fused-open-validate-fold.md).

Change 0202 adds no selector or CRUD closure and leaves the matrix
unchanged. It folds the pass-2a calculation-settings parse into the fused
open tokenization (fifth handler; standalone calculation parse unchanged
for its other call sites), so the source-backed open tokenizes
`content.xml` exactly once. Frozen cross-binary CPU-2 A/B/B/A accepts the
source-backed open p50/mean/p95/p99 (17.41%-23.21% lower), the one-edit
lifecycle all-four plus commit p50/mean, the one-percent lifecycle
p50/mean/p95 in primary and rerun (rerun commit all-four), the
repeated-edit total p50/mean, stage p50/mean/p95, and commit all-four, and
the eager-open mean; the one-percent lifecycle p99 primary adverse reading
did not reproduce in the single permitted rerun, and a sub-0.5%
repeated-edit publication p50/mean reading on source-identical phases is
documented as code-layout wobble. See
[`0202`](changes/0202-ods-single-tokenization-open.md).

Change 0203 adds no selector or CRUD closure and leaves the matrix
unchanged. It explored memoizing the fused open driver's namespace
classifications (HashMap v1, then a direct-mapped slot cache v2); v1
measured a mechanism-confirmed regression on all workloads and v2
measured the targeted open neutral with broadly adverse guardrails, so
both were withheld and reverted byte-exact to the 0202 state. See
[`0203`](changes/0203-ods-open-namespace-classify-memo.md).

Change 0204 adds no selector or CRUD closure and leaves the matrix
unchanged. It fuses the ODS protection double-parse of `content.xml` into
one tokenization; the repeated-edit stage phase accepted all-four at
20.01%-25.86% lower. Originally withheld for a reproduced both-directions
adverse pattern on the source-open phase (executing none of the changed
code), it was re-verdicted **banked** after change 0205 calibrated the
per-binary-pair layout noise floor and showed every blocking reading sits
within it. Claim scope is the repeated-edit stage statistics only. See
[`0204`](changes/0204-ods-protection-fused-parse.md).

Change 0205 is a measurement-methodology calibration (no code change, no
selector or CRUD closure; matrix unchanged). Probe binaries carrying
never-executed parser-shaped padding measured the per-binary-pair layout
noise floor per phase, and the banking rule was refined: within-floor
adverse readings on phases executing no changed code are layout readings
and do not block, adverse readings on executed phases still block unless
rerun-cleared, and accepts are claimed only above the floor. See
[`0205`](changes/0205-layout-noise-floor-calibration.md).

Change 0206 adds no selector or CRUD closure and leaves the matrix
unchanged. It removes dead per-element qualified-name materialization in
the ODS settings location scan (shell and fused handler identically);
banked with a deterministic allocation claim (source-open allocations
-27.96%, allocated bytes -2.33% on the measurement corpus) and neutral
latency under the 0205 floor rule. See
[`0206`](changes/0206-ods-lazy-settings-qname.md).

Change 0207 adds no selector or CRUD closure and leaves the matrix
unchanged. It byte-matches resolved attribute names in the ODS worksheet
codec and allocates owned strings only for consumed values, preserving
the exact per-attribute error order; banked with a deterministic
allocation claim (source-open allocations -19.22%, bytes -5.13%;
cumulative with 0206 -41.81%) and neutral latency under the 0205 floor
rule. See
[`0207`](changes/0207-ods-byte-matched-attributes.md).

Change 0208 adds no selector or CRUD closure and leaves the matrix
unchanged. It tried borrowed decoding in the ODS commit validate-reparse
with a deterministic -43.0% allocation win per commit transaction, but
the adverse both-directions latency patterns on the executed commit
phases reproduced in both rule-2 reruns — withheld and reverted per
banking rule 2; the tree is bit-exact back at the 0207 state. See
[`0208`](changes/0208-ods-borrowed-validate-decode.md).

Change 0209 adds no selector or CRUD closure and leaves the matrix
unchanged. It fuses the ODT source-backed open's two complete
`content.xml` tokenizations (validation + content style-registry scan)
into one pass with exact error precedence; banked with
`odt_file_source_open` p50/mean/p95 14.63%-18.99% lower in both
directions (pre-floor acceptance; the eager-open adverse reading did
not reproduce in the single permitted rerun). See
[`0209`](changes/0209-odt-fused-open-parse.md).

Change 0210 adds no selector or CRUD closure and leaves the matrix
unchanged. It replaces the ODP slide parser's per-lookup attribute
re-scans with a lazy per-element cache (exact per-lookup error
semantics); banked with all-four-statistic accepts on every executed
workload (`odp_semantic_list_slides` 16.77%-47.07%, `one_slide`
7.45%-19.70%, `full_text` 13.11%-21.71% lower, pre-floor acceptance). See
[`0210`](changes/0210-odp-lazy-element-attrs.md).

Change 0211 adds no selector or CRUD closure and leaves the matrix
unchanged. It fuses the ODP per-query double scan of `content.xml`
(transition-style collection + slide parse) into one tokenization with
exact historical error precedence; banked with
`odp_semantic_list_slides` p50/mean 15.56%-17.84%, `one_slide` p50/mean
18.88%-19.50%, and `full_text` p50 15.85%-16.43% lower (pre-floor
acceptance; the byte-identical open workload's adverse primary reading
cleared in the single permitted rerun). See
[`0211`](changes/0211-odp-fused-query-parse.md).

Change 0212 adds no selector or CRUD closure and leaves the matrix
unchanged. It caches attribute namespace resolution in the ODP slide
parser (resolve each attribute key once at scan; lookup replay compares a
cached snapshot with no resolver calls); originally withheld under the
pre-floor rule for a reproduced adverse open p50 reading on a
byte-identical phase, re-verdicted **banked** after the 0213 floor
reclassified it as a within-floor layout reading, and re-applied
bit-exact. Claim scope: `full_text` p50/mean/p95/p99 20.82%-29.50%,
`list_slides` p50/mean 19.15%-25.51%, `one_slide` p50/mean/p99
17.48%-31.16% lower. See
[`0212`](changes/0212-odp-cached-attr-resolution.md).

Change 0213 is a measurement-methodology calibration (no code change, no
selector or CRUD closure; matrix unchanged) — the litchi-odp analog of
0205. Probe binaries carrying never-executed parser-shaped padding
measured the litchi-odp per-binary-pair layout noise floor per phase
(open p50/mean 3.1%/2.5%, list-slides 2.0%/3.6%, one-slide 2.5%/3.2%,
full-text 0.1%/0.5%), and the 0205 banking rule extends unchanged to
litchi-odp. See
[`0213`](changes/0213-odp-layout-floor-calibration.md).

Change 0214 adds no selector or CRUD closure and leaves the matrix
unchanged. It is a profiling/target-selection analysis (no code change):
post-0212 `perf record` profiles of the four ODP workloads located the
remaining hotspots and selected the single-scan shape-attribute harvest
as the next target. See
[`0214`](changes/0214-odp-post-0212-reprofile.md).

Change 0215 adds no selector or CRUD closure and leaves the matrix
unchanged. It folds the ODP `drawing_attributes` fresh attribute re-scan
into the shared `ElementAttrs` incremental scan (exact document order,
error-message identity by first reach, duplicate detection, and decode
positions; 7 new pinning tests); banked under the 0213 floor rule with
`odp_semantic_list_slides` p50/mean 6.53%-9.97%, `one_slide` p50/mean
8.80%-15.10%, and `full_text` p50/mean 5.44%-11.32% lower; the
byte-identical open workload's adverse p50 reading is a within-floor
layout reading. See
[`0215`](changes/0215-odp-shape-attr-harvest.md).

Change 0216 adds no selector or CRUD closure and leaves the matrix
unchanged. It is a profiling/target-selection analysis (no code change):
post-0215 profiles of the ODT semantic workloads refuted the
double-tokenization hypothesis and selected the discard-but-validate text
extraction implemented as 0217. See
[`0216`](changes/0216-odt-query-reprofile.md).

Change 0217 adds no selector or CRUD closure and leaves the matrix
unchanged. It gives the ODT `extract_text` path a discard-but-validate
parse mode (no retained `Element` materialization; exact text, error
precedence/messages, and limits pinned against the pre-change path as a
live oracle). All three executed workloads accepted all four statistics
(`full_text` 42.12%-52.07%, repeated-text cached 51.91%-57.18%, uncached
40.31%-53.85% lower); the reproduced guardrail layout readings (open p50
max 3.26%, list-paragraphs mean max 6.67%) were reclassified as
within-floor layout readings by the 0218 calibration, and the change was
re-verdicted **banked** and re-applied bit-exact. See
[`0217`](changes/0217-odt-discard-validate-text.md).

Change 0218 is a measurement-methodology calibration (no code change, no
selector or CRUD closure; matrix unchanged) — the litchi-odt analog of
0205/0213. Probe binaries carrying never-executed parser-shaped padding
measured the litchi-odt per-binary-pair layout noise floor per phase
(open p50/mean 3.3%/7.2%, list-paragraphs 5.2%/6.7%, one-paragraph
4.8%/8.3%, full-text p50 4.1%, repeated-text-cached 7.1%/7.4%,
repeated-text-uncached 4.8%/4.1%; tails considerably wider), and the 0205
banking rule extends unchanged to litchi-odt. See
[`0218`](changes/0218-odt-layout-floor-calibration.md).

Change 0219 is profiling and target-selection analysis only (no code
change, no selector or CRUD closure; matrix unchanged): dwarf call-graph
profiles of the three ODF family open workloads attributed the shared
content validator `validate_content_document_part` (~69% of the timed
`odt_semantic_open` call) to per-event buffer copies and per-event
namespace resolution consumed only at depth ≤ 2, and selected the
borrowing-reads + depth-gated-resolution target implemented as 0220. See
[`0219`](changes/0219-odf-validator-reprofile.md).

Change 0220 adds no selector or CRUD closure and leaves the matrix
unchanged. It rewrites the shared ODF content validator with borrowing
`read_event()` reads and depth-gated namespace resolution, with the
pre-change body retained as a cfg(test) oracle (accept/reject parity and
byte-identical error messages across synthetic edge cases and the full
ODF corpus). The sole production caller is the ODT owned open path.
**Banked**: `odt_semantic_open` p50 5.99%-7.30% lower (over the 0218
floor) and `odt_file_eager_open` p95 19.12%-21.14% lower claimed; the
`odt_semantic_full_text` p95 primary adverse was cleared by its rerun,
and `ods_file_source_open` p95 is recorded as a flagged above-floor
layout reading (max 4.90% vs the 4.5% floor) on a zero-changed-code
phase, watch-listed for the next ODF change. See
[`0220`](changes/0220-odf-validator-borrowing-reads.md).

Change 0221 adds no selector or CRUD closure and leaves the matrix
unchanged. It replays the 0220 borrowing-reads + depth-gated-resolution
transformation on the fused source-backed ODT open (`OpenParse::run`),
with the pre-change loop retained as a cfg(test) oracle (parity across 26
synthetic edge cases and 69 ODT fixtures, byte-identical errors).
**Banked**: `odt_file_source_open` p50/mean/p95 43.28%-46.30% lower and
`odt_file_source_open_full_text_lifecycle` p50/mean/p95 20.55%-23.22%
lower; all guardrails clean or within-floor, and the 0220 watch-listed
`ods_file_source_open` p95 read within floor under this layout — flag
cleared. See
[`0221`](changes/0221-odt-openparse-borrowing-reads.md).

Change 0222 adds no selector or CRUD closure and leaves the matrix
unchanged. It promotes the fused `OpenParse` to the owned ODT open path
(`from_owned_package`): one borrowing, depth-gated pass replaces the
standalone validator scan plus the content-styles rescan, with
stage-by-stage error precedence preserved byte-exactly (pinned by
cross-stage parity tests against the cfg(test) sequential oracle).
Provisionally withheld under the pre-floor rule (lifecycle p50 adverse
reproduced at max 1.76%), re-verdicted **banked** under the 0223 floors:
`odt_semantic_open` p50/mean/p95 6.33%-6.45% / 10.02%-11.55% /
31.02%-33.20% lower and `odt_file_eager_open` p50/mean/p95/p99
12.05%-17.05% / 12.29%-16.82% / 11.46%-18.86% / 13.40%-17.27% lower. See
[`0222`](changes/0222-odt-owned-open-fused-parse.md).

Change 0223 is a measurement-methodology calibration (no code change, no
selector or CRUD closure; matrix unchanged) — the 0218 analog extended to
the ODT source-path and eager-open phases. Probe binaries carrying
never-executed parser-shaped padding (.text +6.1KB to +14.6KB) measured
pure layout noise; effective floors: file-source-open p95 2.5% / p99
28.0%, file-source-lifecycle 3.8%/2.5%/4.0%/6.5% (p50/mean/p95/p99),
file-eager-open 5.6%/5.7%/9.3%/9.2%. The 0205 banking rule extends to
these phases unchanged. See
[`0223`](changes/0223-odt-source-path-floor-calibration.md).

Change 0224 adds no selector or CRUD closure and leaves the matrix
unchanged. It replaces `NsReader` in the fused ODT open parse with a
plain `Reader` plus a hand-rolled `BindingTracker` whose error stream and
resolutions are byte-exact by construction (differential oracles at every
depth; 898 litchi-odt tests). A v1 tiny-shape p50 regression was
diagnosed as a real fixed per-open overhead (crossover ~70 paragraphs)
and eliminated in v2. **Banked**: `odt_file_eager_open` p50/mean/p99
9.41%-13.07% / 9.86%-14.48% / 14.01%-25.39% lower, `odt_file_source_open`
p50 21.85%-23.05% lower, `odt_file_source_open_full_text_lifecycle` mean
9.39%-11.77% lower; all guardrails clean or within-floor. See
[`0224`](changes/0224-odt-openparse-binding-tracker.md).

Change 0225 adds no selector or CRUD closure and leaves the matrix
unchanged. It adds a last-prefix namespace resolution memo
(`TextNamespaceMemo`) to the discard-but-validate ODT text path
(`parse_text_block_texts`), replacing per-event `resolve_event` reverse
scans with a content-versioned memo whose invalidation is provably exact
(differential rebinding battery, per-event classification replay,
corpus-wide parity; 900 litchi-odt tests). Provisionally withheld
pending the 0226 floor calibration, then **banked**:
`odt_semantic_full_text` p50/mean 15.79%-17.44% / 19.03%-20.70% lower,
`odt_repeated_text_cached` p50/mean/p95 20.10%-20.24% / 19.85%-20.68% /
19.02%-21.93% lower, `odt_repeated_text_uncached` p50/mean/p95/p99
21.79%-23.51% / 22.56%-23.42% / 23.58%-24.96% / 26.73%-26.96% lower,
`odt_file_source_open_full_text_lifecycle` p50 15.96%-16.50% lower;
guardrails clean, within-floor, or cleared by rerun. See
[`0225`](changes/0225-odt-text-resolution-memo.md).

Change 0226 is a measurement-methodology calibration (no code change, no
selector or CRUD closure; matrix unchanged) — the 0223 analog completing
the `odt_file_source_open` floor set after 0225's residual guardrail
reading. Probe binaries carrying never-executed parser-shaped padding
(.text +3,872/+5,984/+7,744 B) measured pure layout noise; folding in
the 0225 byte-identical-phase history, the file-source-open effective
floor is now p50 6.7% / mean 6.1% / p95 45.0% / p99 38.7% (p95/p99
supersede 0223's 2.5%/28.0%; lifecycle/eager floors unchanged). The 0205
banking rule applies unchanged. See
[`0226`](changes/0226-odt-source-open-floor-calibration.md).

Change 0227 adds no selector or CRUD closure and leaves the matrix
unchanged. It lifts the 0224 `BindingTracker` into the shared
crate-private `litchi-odt::binding_tracker` module and rewires the
discard-but-validate ODT text path (`parse_text_block_texts`) from
`NsReader` to a plain borrowing `Reader` with hand-maintained binding
push/pop, removing the per-event `process_event` cost while keeping the
tokenization, namespace error stream, and the 0225 resolution memo
byte-exact (lockstep tracker-vs-`NsReader` differential replay and an
adversarial reserved-prefix/declaration-limit battery; 901 litchi-odt
tests). **Banked**: `odt_semantic_full_text` p50 9.44%-13.32% lower,
`odt_repeated_text_cached` p50/mean/p95 17.57%-18.52% / 17.66%-18.11% /
13.60%-17.56% lower, `odt_repeated_text_uncached` p50/mean/p95/p99
18.79%-20.32% / 18.45%-20.37% / 17.43%-21.68% / 8.65%-18.10% lower
(rerun 0227r), `odt_file_source_open_full_text_lifecycle` p50/mean
10.37%-13.34% / 9.10%-13.15% lower; guardrails clean, within-floor, or
cleared by rerun. See
[`0227`](changes/0227-odt-text-binding-tracker.md).

Each new case should use deterministic object positions and digests, separate
semantic work from publication, reopen outputs, verify untouched content, and
record source/sink, allocation and peak-memory behavior in addition to time.
