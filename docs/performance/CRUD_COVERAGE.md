# Performance CRUD coverage

Date: 2026-08-14

This is a coverage map, not a completion claim. It compares the 200 selectable
benchmark cases and the explicitly labeled correctness-only APIs with
`docs/CRUD_Scenario_Checklist.md`. Generic ZIP/OPC/CFB substrate measurements
do not certify format-semantic CRUD behavior, and API-only coverage is not a
performance claim.

| Required scenario | Current status | Coverage evidence |
|---|---|---|
| Open and identify format | Partial | ZIP/OPC/CFB plus owned DOC/XLS/PPT/RTF/XLSX and source-backed XLSX open; opt-in bounded RTF/XLS/DOCX/PPTX/generic-ODF reports now exercise format validation, while RTF still covers plain, raw CP-1252, LZFu and a real-producer watermark input; no smart-detection handoff case |
| List semantic children without payloads | Partial | XLS/XLSX/ODS sheets, DOC/RTF/DOCX/ODT paragraphs and PPT/PPTX/ODP slides; opt-in `docx_section_inventory` now lists the exact source-backed DOCX section descriptors, but broader section/edit matrices remain missing |
| Query one property or named object | Partial | XLS/XLSX/ODS cells, one DOC/RTF/DOCX paragraph, an indexed fully validated ODT paragraph, one PPT shape and one PPTX/ODP slide; the already-open RTF paragraph query reuses its parser-derived exact story length and cardinality, while explicit sparse `nth` skips discarded-view construction; broader properties/images remain missing |
| Read one cell/paragraph/slide/image/Part | Partial | XLS/XLSX/ODS cells, DOC/RTF/DOCX paragraphs, indexed ODT paragraphs, PPT/PPTX/ODP text objects and generic OPC Part. Three opt-in OPC source-cache selectors add exact managed-Budget boundary evidence plus finite-control/managed same-Part and fixed-work disjoint-Part contention across `1/2x`, `1x`, and `2x` capacities and capped worker widths; their deterministic source delay is correctness evidence, not a latency claim. Semantic image selection remains missing |
| Scan all cells/paragraphs/slides | Covered for generated native/OOXML/RTF/ODF text corpora | XLS/XLSX/ODS cell scans, including the isolated ODS public-cell sweep; DOC/RTF/DOCX/ODT paragraph enumeration and PPT/PPTX/ODP slide/text enumeration |
| Full text extraction | Covered for generated DOC/PPT/RTF/DOCX/PPTX/ODT/ODS/ODP | Complete deterministic text or row-major cell text is checked; RTF additionally verifies raw CP-1252/LZFu text and the body plus public header-shape projection of a real-producer watermark; ODT consuming block ownership is measured in change 0023; broader real-producer/media-heavy corpora remain missing |
| Semantic conversion to sequential sink | Partial | Measured: RTF writes semantic body text to a bounded forward-only, non-seek UTF-8 sink with configurable paragraph separators; output and sink counters are fully verified outside timing. Production correctness tests, rather than the timing sink, cover partial/interrupted/write-zero behavior. No other-format semantic sink benchmark exists |
| Create a small document | Partial | Fresh DOC/XLS/PPT plus DOCX/PPTX/ODT/ODS/ODP public authoring, and narrow XLSX/RTF forward-only creation with selectable timing evidence; broader streaming authoring remains missing |
| Create or append a very large stream | Partial | XLSX one-sheet scalar rows and plain-run RTF paragraphs have opt-in small/medium/large fixed-window creation cases through non-seek hashing sinks, deterministic untimed reopen, exact counters, and output hashes. Existing-document RTF logical tail append is now a separate opt-in small/medium/large bounded case with a 16 KiB non-seek sink window, exact byte counters, and untimed patch/reopen/source-conflict gates; it does not claim that the retained candidate snapshot is window-bounded. Large fresh legacy writers still accumulate before final output |
| Exact no-op edit and commit | Covered for generated DOC/XLS/PPT/RTF/XLSX/DOCX/PPTX/ODT/ODS/ODP | Public semantic transaction plus save/reopen; RTF also proves exact raw CP-1252, LZFu and producer-watermark publication, and the dedicated logical-tail no-op case proves shared snapshot identity plus exact sequential bytes; signed/extension corpora remain missing |
| One semantic edit and save | Covered for generated DOC/XLS/PPT/RTF/XLSX/DOCX/PPTX/ODT/ODS/ODP | Cell/paragraph/shape edit or supported ODP slide append, then save/reopen; XLSX now also has six matched opt-in eager/source-backed scalar-cell publication controls over deterministic medium and dense/sparse four-sheet media-rich corpora (one cell, `ceil(1%)`, and exact 256-cell batch). Their timed interval covers open/stage/commit/sequential publication, while lifecycle, reopen, topology, media identity, hashes and source counters are untimed; change 0096 accepts the source-backed provenance-reuse result without making an eager, physical-I/O, allocation or RSS claim. Native XLS also has matched eager/source-backed existing-comment and worksheet-visibility controls over deterministic opaque-heavy CFB corpora. ODT separately measures paragraph replacement, inline line-break/run/hyperlink insertion, structural paragraph insertion/removal, and matched existing-image replacement while ODT, ODS and ODP verify exact retained resources and manifest media types; source-backed DOCX/PPTX and narrow XLSX calculation-metadata/defined-name/sheet-protection/data-validation/auto-filter/conditional-formatting publication likewise verify eight exact 2 MiB media Parts after one semantic edit. RTF now separately measures bounded existing-document logical-tail append; its sequential sink, full reopen, durable apply/inverse and foreign-source refusal are untimed gates, not a speedup claim. Correctness-only additions include source-backed XLSX existing scalar-cell set/clear/remove and row visibility, XLS worksheet visibility, RTF direct paragraph-layout updates and same-length standalone PNG/JPEG payload replacement, and ODT plain-paragraph moves. ODS content-validation catalog CRUD and direct PPTX transition edits retain their previously documented narrow closures. A separate RTF correctness gate proves real-producer root-shape edit plus checked LibreOffice resave/readback without presenting it as a paragraph benchmark |
| About 1% semantic update and save | Covered for XLSX/DOCX/PPTX/RTF/ODT/ODS generated corpora | Deterministic evenly spaced `ceil(1%)` cell, paragraph and shape changes; the matched XLSX scalar-cell control pairs eager and source-backed selector-first multi-sheet publication over both fixed media-rich shapes, and change 0096 accepts the source-backed provenance-reuse result. DOCX and RTF use one canonical atomic paragraph batch, ODT coalesces ordinary scalar durable replacements internally, and ODS partitions flat cell positions by worksheet into bounded atomic `set_cells` calls; each commits once and reopens the package. ODS still has no comparative latency claim pending release ABBA; the RTF batch case has a matched scalar-loop comparison recorded in change 0081 |
| Bulk update matching objects | Partial | Selectable: XLSX has a matched eager/source-backed exact-256 existing scalar-cell batch over all four sheets; PPTX has a bounded atomic batch across up to 32 existing slides, with up to 256 unique nonoverlapping shape-text selectors per slide; RTF measures one bounded ordered paragraph batch; native XLS measures the exact 256-existing-comment and 64-worksheet-visibility limits through matched eager/source-backed controls; XLSX also replaces a complete three-owner core conditional-formatting collection through matched eager/source-backed controls; ODP measures eight fixed-name existing text boxes across eight slides; ODT replaces 64 fixed existing image owners through matched scalar/bounded-batch controls. Change 0096 accepts only the XLSX source-backed scalar-cell provenance-reuse result; the other evidence additions still make no latency claim pending frozen ABBA. API-only correctness coverage: ODS updates up to 4,096 cells on one selected sheet; native PPT updates up to 256 persisted shape-text targets; XLSX also clears the complete conditional-formatting collection; ODT correctness extends to the exact 256-change embedded object/image bound; ODP correctness extends to the exact 256-owner batch bound. These remaining APIs are not selectable timing evidence |
| Clear/remove/hide/detach/GC distinctions | Partial | Measured: ODT exact paragraph removal intentionally preserves orphaned resources, and RTF removes one exact middle ordinary paragraph on its narrow generated plain-source closure. API-only: OLE2 deletes one or a bounded stream set while retaining storages and unrelated streams; XLS shows/hides/very-hides worksheets, while XLSX additionally shows/hides/very-hides/activates tabs, hides/unhides existing rows, clears scalar cell values without deleting owners, or physically removes supported existing scalar `<c>` owners; ODS clears only unbound validation definitions; DOCX removes direct plain main-document paragraphs; PPTX and ODP remove only dependency-free supported slides; RTF removes standalone PNG/JPEG picture groups; ODT removes selected resource owners while retaining payloads. DOCX reversible hyperlink-wrapper detachment remains distinct from irreversible relationship/URL redaction. Opt-in ODT GC remains exact-source and explicit-name only. General cascading delete, orphan collection and dependency-aware removal remain missing |
| Sanitization and irreversible redaction | Partial | API-only: DOCX now has two deliberately separate exact-source flows. Reversible detachment unwraps selected main-document external `w:hyperlink` owners but retains relationship records and target URLs. Irreversible redaction inventories and selects exact target URLs, unwraps their visible owners, removes the corresponding external hyperlink relationships, exposes effects before publication, and intentionally provides no inverse. It is bounded, fail-closed and main-document-only; it is not general personal-data cleanup, field/DDE sanitization, embedded-object cleanup or package-wide external-reference removal |
| Copy object with dependency closure | Partial (correctness only) | PPTX atomically applies a bounded in-package slide-copy plan for its supported finite owned dependency graph while reusing shared layout dependencies and refusing unsupported/shared ownership; ODP copies only dependency-free blank slides; ODS copies only dependency-free worksheets; DOCX copies only direct plain main-document paragraphs. These APIs have reopen and preservation tests but no selectable timing evidence. General cross-document copy, arbitrary charts/media/themes/notes/formulas/styles, name-collision reconciliation and complete dependency closure remain missing |
| Merge and split | Partial (selectable correctness evidence) | XLSX now has two opt-in eager commit-plus-bounded-save cases over one deterministic sparse `Sheet1` A1:B2 fixture: merge and unmerge prepare the transaction outside timing and time only commit plus `Workbook::write_to`; untimed checks reopen merge membership, anchor retention, covered/uncovered and unrelated-cell semantics, exact durable patch apply/inverse restoration, and stale-source refusal. These cases make no latency claim without controlled ABBA evidence; broader format merge/split coverage remains missing |
| Patch encode/apply/invert/merge | Partial | Durable encode/decode/apply/inverse coverage now also includes exact DOCX plain-paragraph copy/removal, ODT plain-paragraph move, PPTX dependency-free slide removal, ODS worksheet move/copy, XLS worksheet visibility, RTF picture-payload/removal edits, and the new logical-tail append's source-checked durable replay/inverse/reopen/foreign-source gates. Source-backed XLSX scalar-cell and row-visibility edits and RTF paragraph-layout edits have source-bound apply/inverse but no durable wire claim. The selectable RTF paragraph remove/move and logical-tail cases keep durable work outside timing. DOCX irreversible hyperlink redaction intentionally has no inverse. Patch merge/join/three-way conflict resolution remains incomplete and unmeasured |
| Validate without mutation | Partial (bounded opt-in evidence) | Opt-in `rtf_validation_report`, `xls_validation_report`, `docx_validation_report`, `pptx_validation_report`, and `odf_validation_report` now retain deterministic report hashes, check IDs/statuses, issue codes/counts, source hashes, and bounded source-read counters for ReadAt validators; `docx_section_inventory` additionally retains the exact source-backed section topology. Existing CFB/OPC/generic-ODF reports remain correctness-only APIs. These selectors are not default cases and carry no speedup claim; they do not decrypt content, verify signatures, execute macros, fetch links, or provide general repair, and broader format-semantic/security matrices remain open |
| Explicit repair plan | Partial (correctness only) | Generic ODF exposes one bounded typed `RepairPlan<NonDestructive>` for an otherwise-valid first/stored `mimetype` with one recognized Extended Timestamp local-header extra. It binds source length and SHA-256 evidence, rejects stale/foreign reports and unsafe ZIP/security/semantic cases before output, previews deterministic changed-member/archive effects without source bytes, removes only that local extra, reopens the candidate, proves member digest equality and raw preservation, and reports sequential-sink progress. Explicit `apply` yields a source-checked reversible patch with an exact inverse; destructive plans and unsupported structural/XML/encryption/signature/macro repairs remain unconstructible. No latency claim is made. |
| Preserve unknown extension during understood edit | Partial | Targeted OPC raw-copy framing/unknown-member tests, exact untouched opaque ODS rows/members, and exact raw ODT/ODS/ODP auxiliary/media members during neighboring edits; ODT resource replacement retains frame attributes and unknown children while removal retains payload files; ODP model batches retain unselected producer content; the new XLSX closures raw-copy unselected Parts and refuse selected MCE/unknown owners. This is narrow fail-closed preservation, not general extension-aware editing |
| Replace one or a bounded set of low-level Parts, preserve the rest | Covered for owned OPC, bounded source-backed OPC, guarded DOCX main-document semantics, guarded PPTX selected/multi-slide semantics, and guarded XLSX worksheet/workbook semantics | Changes 0008/0021/0022 test owned raw framing, fallback and payload ownership; changes 0037/0077 add consuming one-Part/bounded multi-Part publishers; changes 0039, 0044/0063/0077, 0046, 0061, 0067, 0070, 0073, 0076, 0078, 0079 and 0080 integrate accepted measured guarded semantic transactions while refusing unsafe MCE, signatures, stale closure, topology, relationship, style-reference and printer-reference cases before output. Change 0082 adds matched selectable XLSX conditional-formatting publication through one selected worksheet after workbook/relationship/styles validation, with no latency claim before ABBA. Correctness-only guarded closures add direct standard PPTX transition set/replace/clear through one selected slide Part, XLSX tab-state publication through the workbook alone or the workbook plus old/new active worksheets, and DOCX main-document-only hyperlink-wrapper detachment. General XLSX cell/formula/table-filter, printer graph and structural/inherited/extension-transition PPTX editing remain outside the closure |
| Same-length OLE2 stream edits and metadata moves | Substrate plus bounded native-XLS consumers | `litchi-cfb` resolves existing logical streams through validated FAT/MiniFAT chains for bounded same-length whole-stream overlays and byte-range splices, and can move existing stream directory entries without copying payload sectors. Its writer also moves storage subtrees atomically while retaining descendant payload allocations and order. Exact source/version and target checks, overlap/duplicate/path/topology limits, complete reopen, direct sequential publication and atomic path output are covered where exposed. `litchi-ole-common` retains signing/encryption/DRM refusals. Change 0094 proves proportional MiniFAT exact-range reads at the generic substrate. Change 0095 adds semantic existing-comment and worksheet-visibility publishers: one/256-comment plans submit 109/27,904 bytes instead of an 80,946-byte Workbook, and one/64-visibility plans submit 1/64 instead of 18,166. Balanced ABBA accepts no latency speedup; allocation/RSS/source-I/O, DOC/PPT consumers, unstable FAT tails and cold/high-latency evidence remain open |

The source/output matrix is also incomplete. Owned bytes and instrumented
`ReadAt` exist for OPC/XLSX, and the deterministic range simulator covers
latency/range effects. The bounded source-backed OPC publisher accepts a
forward-only sink and records complete positional input plus bounded output,
but this is not semantic conversion or memory-bounded authoring. RTF and DOCX
final serialization also accept and test a forward-only non-seek sink. RTF
semantic body-text output independently has a selectable bounded forward-only
sink benchmark. Narrow XLSX scalar-row and RTF plain-run creation also have
selectable fixed-window forward-only evidence whose timed sinks retain zero
output bytes. Existing-document RTF logical-tail append now has a separate
16 KiB windowed non-seek publication benchmark; its sink retains zero output,
while the API's candidate snapshot remains intentionally outside that sink
window claim. The correctness-only same-length OLE2 overlay substrate also
accepts a direct sequential sink and an atomic path destination, but no format
owner consumes it in an adopted end-to-end performance path yet.
The new validation tranche uses borrowed generated bytes for RTF and generic ODF
and instrumented positional `ReadAt` for XLS, DOCX, and PPTX; source hashes and
validation topology are checked outside timing, and DOCX section descriptors are
checked exactly.
Borrowed-byte comparisons, filesystem positional cold reads, atomic-save
timing, broad structural PPTX streaming output, and non-seek semantic
conversion for other formats remain.

## Highest-return next cases

1. Attribute remaining final owner/public-reader work after the accepted native
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
   RSS claim. Scalar-cell clear/physical-remove and row visibility remain
   source-backed correctness coverage, while conditional-formatting replacement
   is selectable without an ABBA claim and tab visibility remains unmeasured.
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
   0075. Plain ODT paragraph move, ODS worksheet move/dependency-free copy, and
   ODP dependency-free blank-slide copy/removal are correctness-only; new style
   creation, general dependency closure, extension preservation and broader
   structural operations remain.
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
identity. The harness also
exercise source-backed exact no-op, clear, and physical-remove behavior. Source
read and successful materialization counters are reported outside timing.
Change 0096 retains a CPU-pinned balanced release ABBA comparison for those six
source-backed selectors. The p50 geomean improves 21.66%/22.65% and p95 improves
21.38%/22.70% in the two directions, with exact output hashes. Physical source
read and materialization counters are unchanged, so the accepted result is the
removal of a redundant semantic worksheet reload/reparse, not an I/O,
allocation, RSS or cold-filesystem claim. The default 36 cases / 198 records
remain unchanged.

Change 0097 separately measures fresh RTF streaming creation after bounded
escape-free ASCII batching. The p50 geomean improves 76.41%/76.47% and p95
75.23%/75.76%; the large case reduces sink calls from 7,208,970 to 1,441,802
under a hard 32-byte request ceiling, with exact bytes and hashes. This does not
claim existing-document edit, allocation, RSS or cold-I/O improvement. iWork
remains deliberately deferred while the `iwa-*` crates change separately.

Each new case should use deterministic object positions and digests, separate
semantic work from publication, reopen outputs, verify untouched content, and
record source/sink, allocation and peak-memory behavior in addition to time.
