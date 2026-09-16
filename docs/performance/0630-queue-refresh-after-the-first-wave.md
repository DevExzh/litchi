# 0630: the 0587 queue, worked: what landed, what fell, and what the survey got wrong

Status: retained, coordination record. `performance_claim: none` — every number
below is quoted from the change record that measured it, with that record's own
tier and floor; nothing is re-measured or aggregated here.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## What this record is

Change [0587](0587-remaining-opportunity-survey.md) ranked 36 opportunities
across the OLE2 and OOXML path and authorized none of them. This record closes
the wave of changes that worked that queue — 41 numbered records, 0588 to
0631, each written by one implementing agent in its own worktree and branch,
each with its own paired measurement, evidence packet and four log sections, and
each merged onto the main branch by the coordinator in the order it finished.
It states the outcome per queue item, quotes the numbers, lists the survey's
own errors that the wave found, and refreshes the ranked queue in
`HOTSPOTS.md` so that the next wave starts from what is left rather than from
what was believed.

The method of the wave is the method of the program: before measurements,
the smallest coherent change, correctness and preservation oracles over the
fixture corpus, after measurements with floors measured in the same window, a
frozen design where a contract would move, and a retained rejection where the
paired result lost. Two rules were added for the wave and held: no agent
edited the program logs (each shipped its four paragraphs in its packet, and
the coordinator inserted them), and no agent built in the shared working copy.

## Outcome per queue item

Rank and identifier as in 0587's table; "landed" means the production change is
on the main branch with its record; "frozen" means a design record exists and
nothing was implemented; "declined" and "falsified" mean the record measured the
item and closed it.

| # | id | record | outcome | the number that decides it |
| ---: | --- | --- | --- | --- |
| 1 | XML-1 | [0588](0588-mce-codec-namespace-emission.md) | **landed in part**; the namespace-emission rewrite **withdrawn** | codec −57.8% Ir, real-file eager open plus cell **−52.1% Ir, −43.9% p50** (floor −0.9%), output byte-identical; the rewrite worked (a further −80% of the codec) but the per-element re-declaration is load-bearing: `litchi-docx` parses sliced ranges standalone and two writers publish the processed stream |
| 2 | DOC-1 | [0589](0589-ole2-snapshot-fingerprint-passes.md) | **landed in part** | hash passes per DOC snapshot open 12 → 6, PPT 4 → 2; native cycles **−36.6% median over all 38 admitted fixtures**, p50 −43% to −48% on the largest four (floor ±0.1%); the third identity pass and the duplicate index parse are load-bearing and kept |
| 3 | PPTX-1 | [0590](0590-pptx-opened-transaction-revision-reuse.md) | **landed** (a, b); (c) designed; (d) rejected | hashes per lifecycle 6 → 4, captures 5 → 4, **−15.8% Ir**; p50 −7.5% batch edit, −38.6% slide move (commit phase −45%), slide remove +2.1% inside its 2.7% floor |
| 4 | DOCX-1 | [0591](0591-docx-edit-single-scan.md) | **landed** (a, b); (c) waits on the whitespace question | main-part scans per edit 4 → 2; p50 **−37% / −32% / −19%** one edit at 10,000 / 200 / 24 paragraphs, no-op −42% to −50% (floors ≤ 2.1%); nine scenarios, none regressed |
| 5 | PPTX-2 | [0598](0598-pptx-cross-copy-revision-cache.md) | **landed** | hashes per copy 12 → 8, serializations 9 → 4; media-rich copy **−24.6% Ir**, cycles −8.4% (floor −0.4%); two of four wall-clock cells inside a noisy floor |
| 6 | XML-2 | [0597](0597-xlsx-selected-cell-ineligibility-gate.md) | gate **frozen** (moves error identity); a defect **fixed** | the gate is worth −63.2% Ir of an ineligible source-backed read but moves ten of sixty first-error classes; the ineligible scan was refusing correctly ordered `mergeCells` on 24 of 180 fixtures (434 rows), now zero mismatches against the materialized parser |
| 7 | SAVE-1 | [0593](0593-opc-publication-pristine-members.md) | **landed** | audits per unchanged 132-member save 42 → 0, reserializations 78 → 0; publish cycles **−63% to −71%**; 1,344 corpus rows byte-identical; one disclosed behaviour change |
| 8 | ZIP-1 | [0594](0594-zip-session-reuse-per-open.md) | **landed, above a measured threshold** | open allocation bytes **−59% / −73%** on the 132-member workbook and the 48-member deck; the first version's DOCX regression was eight extra page faults from holding the 80 KB workspace and is removed by a 16-relationship-member threshold; 15 of 179 fixtures qualify |
| 9 | XLS-1 | [0595](0595-xls-frame-loop-and-sst-walk.md) | **landed** with XLS-2b | `54016.xls` open **−23.8% p50, −25.4% cycles**, one cell −18.1% p50; `WithCustomViews.xls` −6% to −13%; flagship inside the floor; error drops per query 75,858 → 0 |
| 10 | DOCX-2 | [0592](0592-docx-lazy-paragraph-index.md) | **landed** | `document()` **−69.3% Ir**, allocations 51 → 26; full-text selectors −40.2% / −32.4% p50 (floors ≤ 1.5%); `docx_file_eager_paragraph_count` +5.49% pooled, above the trigger, contradicted by the native probe and left open |
| 11 | DOC-2/3/4 | [0596](0596-doc-eager-open-terms.md) | **landed** | eager DOC open **−20.1% / −11.3% / −18.4% Ir** on the three fixtures; `parse_sprms` per PAPX entry 3.1 → 1.1; 57-fixture differential byte-identical; a +2.0% cycle rise on one fixture unexplained |
| 12 | XLSX-2 | [0602](0602-xlsx-real-producer-admission-design.md) | **frozen**, mechanism **corrected** | the value editor admits **0 of 95** real fixtures (93 refuse at the package-root relationship allow-list, before the shared-string gate can fire), and the original-bytes compactness audit refuses **94 of 95** at publication |
| 13 | XML-3 | [0603](0603-xlsx-fused-traversal-marker-admission.md) | **landed in part**, consequence **corrected** | declaration-only compatibility worksheets admitted under an equivalence proof, plan-and-commit −27% to −61% on 0602's projections; on real files the fused traversal completes on **0 of 207 parts before and after**, because the value-only vocabulary refuses them first |
| 14 | ZIP-2 | [0611](0611-zip-single-read-per-member.md) | **landed** | a member's first read is one bounded request instead of two or three: opens **87 → 45, 45 → 24, 12 → 6** requests, exactly as modelled; on the 1 ms-per-request transport p50 **−44% to −53%** (floor ±0.06%); bytes +0.5% to +23%; 0582's differential extended to the changed path and byte-identical over 22,875 archives; local-file timing inside its floor and unclaimed |
| 15 | ZIP-5 | [0623](0623-zip-structural-span-accessor-and-prefetch.md) | **landed** (0577's candidate (c)) | a read-side span accessor and a run-coalesced structural prefetch: the 132-member workbook open **45 → 10 requests for identical bytes**, `shapes.pptx` 24 → 6; over 533 containers 4,160 → 2,569 requests with none costing more; range-source open **−76.6%** (floor ≤ 0.05%); a +12%/+7% local review trigger chased to binary layout (the mechanism itself −0.8%) on a selector whose own floor is wider; gated to the ordinary unmanaged exact-policy open |
| 16 | SAVE-2 | [0593](0593-opc-publication-pristine-members.md) | **landed**; the latency model **falsified** | `write(2)` per 654 KB save **531 → 14**; `save(path)` p50 −2.3%, inside the floor, because the save is fsync-bound (7.5 of 8.3 ms) |
| 17 | XLS-2b | [0595](0595-xls-frame-loop-and-sst-walk.md) | **landed** | `scan_shared_string_records` 1.16 M → 0.24 M Ir on the `54016` query; 0576's differential unchanged over 17,434 entries |
| 18 | XLS-3 | [0605](0605-xls-retained-sheet-index.md) | walk **landed**; retained index **frozen** | one validated scan per sheet: `54016.xls` **32 bytes per cell against 617,079** for per-cell reading, zero disagreements over 90,543 cells; the index has no evictable cache to live in |
| 19 | XLSB-1 | [0599](0599-xlsb-commit-single-parse.md) | **landed** | no-op commit and save **−43.9% Ir, −46.4% p50**; one edit −30.8% Ir, −31.4% p50; a synthetic XLSB generator lifts the 22.7 KB corpus ceiling |
| 20 | PPT-1 | [0606](0606-ppt-record-tree-borrowed-payloads.md) | **landed** with PPT-2 | PPT open **−23.9% Ir**, allocations 613 → 304, retained bytes 2.87× → 1.24× the stream; `ppt_semantic_open` −14.9% p50; a text-only selector +4.7% and a glibc heap-trim loop artifact reported |
| 21 | SAVE-5 | [0610](0610-opc-lazy-part-decode-design.md) | proposed **ADR 0030** drafted; design frozen | an XLSX open-edit-save reads **1 part of 90** it inflates; 259 in-scope `iter_parts()` sites, 27 already on the lazy type |
| 22 | CFB-1 | [0604](0604-cfb-append-reads-design.md) | **declined** | the zero-fill is **0.57% to 2.94% of open cycles** natively, not 9-32% of instructions: `rep stosb` is counted 35× by callgrind; zero on `FileSource` |
| 23 | XLS-2 | [0608](0608-xls-lazy-sst-index-design.md) | **declined** | after 0595 the walk is 3.2% to 35.7% of an open, but a deferred index is a prefix index and **94 of 94** string-bearing fixtures reference the last entry |
| 24 | XLS-6 | [0612](0612-xls-skip-uninterpreted-globals-design.md) | **falsified** | the gate fires on 19 of 123 fixtures and cuts read bytes 85%, and the open is **+5.5% slower owned, +17.2% on a file source** (floor under 1%), each extra read re-resolving the directory entry by path |
| 25 | XLSX-1 | [0622](0622-xlsx-compact-source-facts.md) | **landed** (0551's design) | the commit's second whole-sheet scan **−99%** of its instructions; planning-plus-commit **−34% to −41% Ir**; all fifteen timing scenarios faster at every statistic, +12% to +32% at p50 against floors of 3-16%; 32 of 32 published packages byte-identical; peak RSS +4.2% median inside the before leg's 6.5% spread; unreachable on real producer files (0602) |
| 26 | XLSX-3 | [0613](0613-opc-original-audit-memo.md) | memo **declined**, unreachable | the original-bytes audit is **27.6%** of source-backed publication instructions (0587 modelled at most 28%), but every publication door consumes its package, so a second publication from one source never exists: the memo was implemented, measured at zero hits, and retained as a patch |
| 27 | SAVE-3 | [0607](0607-pptx-authored-slide-regeneration-design.md) | **unreachable**; design frozen | `materialize_presentation` is reachable only from `Package::new()`; all 78 opened decks refuse `presentation_mut`; on the authored path regeneration is 2.3% of instructions and the real cost was the per-member compressor (see 0618) |
| 28 | ZIP-3 | [0600](0600-opc-cold-read-observations-and-name-lookup.md) | **landed**, 5 → 4 not 5 → 2 | observations per cold part read 5 → 4, after-stream reads 102 → 32; the four survivors each prove something no later one does |
| 29 | ZIP-4 | [0600](0600-opc-cold-read-observations-and-name-lookup.md) | **landed** | allocations per open −6%; open plus one part −3.5% Ir; 4,215 members byte-identical |
| 30 | CORE-1 | [0609](0609-facade-doc-source-route-design.md) | **falsified** | the two readers' refusal sets are not nested (the snapshot admits four corrupt-stylesheet files the eager reader refuses), the snapshot serves one query, and costs **6.3× to 9.9×** the latency; only allocator peak favours it |
| 31 | XLS-9 | [0620](0620-xls-edit-save-attribution.md) | attributed; **landed** a duplicate-parse removal; owner swap frozen | four complete parses were 52.1% of an XLS save; reusing the recorded source-policy facts: commit p50 **−32% / −27%**, cycles −25%, peak live bytes −34%; the readback-owner swap is inadmissible because the eager coverage check is not reproducible on the source owner |
| 32 | CFB-2 | [0617](0617-cfb-copy-through-writer-design.md) | XLS **declined**, DOC **blocked on an ADR clarification** | container rebuild is 1.07% of an XLS save and 13.3% to 30.0% of a DOC save in cycles; unchanged-stream share 4.8% on XLS, 81% on one DOC |
| 33 | XLS-4 | [0621](0621-xls-open-fence-count.md) | fence half **landed** | full-text observations **89,789 → 36,503** on `54016.xls`, p50 **−27.6% / −28.3%** (floor −3.7%); the survey's modelled count was within 0.04% |
| 34 | CORE-2 | [0621](0621-xls-open-fence-count.md) | **falsified** on its own terms | the open keeps 22 observations, 18 of them structural (the trailing fence of each globals fill ADR 0005 requires) |
| 35 | CORE-3 | [0615](0615-execution-context-completeness-design.md) | proposed **ADR 0031** drafted | none of the three parallel sessions is reachable from any format crate or the facade; a `workers = 4` session runs 2W + min(W, tasks) threads |
| 36 | CORE-4 | [0624](0624-parallel-changed-member-deflate-design.md) | **frozen**; the threshold **corrected** | deflate of the changed set is **24% to 62%** of every save's cycles, but no real fixture regenerates two members of 256 KiB and the one that nearly does gains 0.997×, while forty 871-byte `.rels` gain 3.67× at width four; a measured balance rule replaces the threshold; implementation waits on ADR 0031 |

Beyond the queue: [0618](0618-zip-writer-deflate-state-reuse.md) took the
finding 0607 surfaced and drives **one Deflate compressor per authored save
instead of one per member** (161 → 1 on a 50-slide deck; authored PPTX save
cycles **−12.2%**, p50 −12.1% against a −1.9% floor; the preservation writer's
cross-member pool was tried and rejected at +14% cycles from page faults);
[0601](0601-perf-harness-real-producer-shape.md) gave the harness the shape real
producers write and priced it — the producer signature costs **7.0×** on a
selected cell and **18.6×** on value-only planning against a marker-free
control; [0627](0627-ole2-range-source-selectors.md) closed the OLE2 range-source
evidence gap for XLS and PPT with fourteen opt-in selectors and a first baseline
(a PPT snapshot open reads 2.02× the file over a range source; 0605's
whole-sheet walk is request-bound at 16,145 requests, 16,061 of them under 512
bytes; and XLS-6's "+46 requests" is now priced at a 1.87× open); [0626](0626-perf-ci-smoke-baseline-fetch.md) made the CI smoke job compare
against the last successful full run's artifact, bound to policy identity, corpus
digest, binary profile and runner labels, advisory rather than a merge gate, and
found that the self-comparison it replaced had been failing closed on every CI
run for 273 commits because the allocator policy omitted a tool-identity field
the harness emits.

## Correctness work the wave produced

Reported as defects and fixed: `OleWriter` serialized documents with two or
more storages in hash-seed order, 8 of 214 fixtures non-deterministically
([0625](0625-cfb-writer-deterministic-storage-order.md), ADR 0006); the XLSX
selected-cell path refused correctly ordered `mergeCells` after any
ineligibility mark ([0597](0597-xlsx-selected-cell-ineligibility-gate.md));
the harness's XLS lifecycle test had asserted the pre-0565 read contract for 54
records and poisoned six further tests ([0619](0619-harness-xls-lifecycle-assertion.md),
which also cleared 0595 and 0605 of a suspected counter change); relationship reuse in `litchi-opc` picked whichever duplicate relationship hash
order met first, four real packages affected, now the smallest rId
([0628](0628-opc-relationship-iteration-order.md), which also audited 497
consumer-crate sites and found three where order still reaches a verdict, fixed
by [0631](0631-ooxml-relationship-order-verdict-sites.md): an XLSX exact no-op
patch refused as `PatchConflict`, a PPTX public snapshot revision with two values,
and the DOCX signature staleness token, each proved unstable over 128 hasher
seeds and then stable, 312 of 312 republished packages byte-identical);
the managed DOCX facade budget test that change 0621's gate run surfaced bisects
to change 0495, 94 commits before the survey base — a stale expectation against a
deliberate 131,072-byte parser workspace fence, invisible because the facade's
default features compile almost none of its tests — and now asserts the contract
([0629](0629-facade-docx-budget-test-bisect.md)); none of 0588-0594 is implicated.

Reported and left for their owners: the ordinary eager XLSX save refuses
sheet-view edits on Excel-produced files over the CRLF after the XML
declaration (0587, confirmed at 94 of 95 packages by 0602 as the publication
audit of original bytes); the `office_crud_demo` example's PPTX update cannot
succeed (0607); the DOC facade's field-table refusals on ordinary Word files
(0587); no `.ppt` fixture reaches a length-changing shape-text edit (0617); MCE
output limits are enforced on the codec's self-expanded stream (0588,
unchanged).

## What the survey got wrong, as the wave found it

The survey's method note said that a profile ranks work and only the record
set says whether it is available. The wave adds a second: **a count of removed
work is not a saving until the cost it trades into is measured natively.**
- Callgrind's per-byte pricing of `rep stosb` overstated the zero-fill 35×
  (0604), and 0574's 53.4 ns per KiB coefficient was about 15× too large
  (0612); both items ranked on those figures fell.
- XLS-6 removed 85% of an open's read bytes and made the open slower, because
  each read it added re-resolved a directory entry (0612); 0574's opportunity
  3 regresses on its own for the same reason.
- XML-1's redundancy is load-bearing (0588), XLSX-2's gate cannot fire (0602),
  XML-3's consequence was one layer off (0603), SAVE-3's scenario does not
  exist (0607), CORE-1's retained-`Vec` claim was wrong and its route is
  slower (0609), CORE-2's fences are mostly structural (0621), and the
  "0586's zero was masked by hashing" reading was wrong: it was structural
  (0589).
- The survey under-counted what real producers cost: the harness now shows
  the producer signature at 7× to 19× (0601), and the source-backed value
  editor admits none of them (0602, 0601).

The survey's own annotations carry each correction beside the item it
corrects, so the 0587 record reads true at every row.

## The queue after this wave

The queue is now a set of prerequisites rather than a set of optimizations:
almost every remaining item waits on a decision a human has to make, a
contract that has to be written, or an infrastructure piece that does not
exist. Ordered by the size of what each unblocks, with the record that priced
it.

| # | item | what it needs first | priced by |
| ---: | --- | --- | --- |
| 1 | the publication audit of original part bytes refuses 94 of 95 real OOXML packages; the same contract is 27.6% of source-backed publication instructions | a human decision on the compactness contract for original bytes (ADR 0006, record 0528); it gates every real-producer edit, XLSX-2, XLSX-3 and the XML-2 gate | 0602, 0613 |
| 2 | lazy OPC part decode (C2′): an eager XLSX open-edit-save inflates 90 parts to read one | acceptance of proposed ADR 0030, then a migration of 259 sites | 0610, 0581 |
| 3 | parallel deflate of changed members: 24-62% of every save's cycles | acceptance of proposed ADR 0031, then the measured balance rule, not the survey's threshold | 0624, 0615 |
| 4 | the MCE codec's namespace-emission rewrite: a further −80% of the codec, −85% of a real-file open plus cell | the slice consumers (`Paragraph::extensions`, `Shape::xml`, five XLSX raw accessors, two publishing writers) must stop depending on per-element re-declaration; 133 sites classified | 0588 |
| 5 | admission of `mc:Ignorable`/`dyDescent` worksheets to the fused traversal, and the value-only vocabulary that refuses every real worksheet | an admission-surface widening designed with 0602's D1-D4 | 0603, 0602, 0601 |
| 6 | the XLSX selected-cell ineligibility gate: −63% of an ineligible read | the limit question and an ADR 0005 clarification on which reader owns a lazily loaded payload's refusal, or an observer-detach signal in `litchi-ooxml-common` | 0597 |
| 7 | PPTX revision proof format: about 23% of what remains in an opened transaction | a frozen design with a magic bump for the `LPRM0001`/`LPCP0002` headers | 0590 |
| 8 | the cross-package copy still builds and deflates the candidate twice (5.96 G Ir) | retaining the planned archive in the plan, an ADR 0005 memory question | 0598 |
| 9 | fragment-only DOCX compaction and the main-part copy in `document_snapshot` | the owner's answer on whitespace compaction across untouched paragraphs (defect 5 of 0587) | 0591 |
| 10 | the DOC and PPT snapshot's remaining hashing: the third identity pass and the duplicate index parse, 2 of the 6 surviving complete reads | an ADR 0006 fence design; until then the facade's `.doc` route stays eager by measurement | 0589, 0609 |
| 11 | a copy-through OLE2 writer for DOC saves (13-30% of a DOC save) | an ADR clarification of physical sector layout policy | 0617 |
| 12 | the retained XLS sheet index (repeat queries) | an evictable weighted cache, which no crate has | 0605 |
| 13 | the XLS commit's second candidate parse (137 M Ir) and `Snapshot::from_bytes`'s double framing | a measured fusion that keeps the coverage proof's order; the readback-owner swap is inadmissible | 0620 |
| 14 | framing the XLS globals once (1.5% of the flagship open) | nothing; the one survivor of 0574's list | 0612 |
| 15 | XLSB: `insert_candidate_cell`/`transfer_cell` reparse twice; `apply_sparklines`/`apply_cell_watches` never refresh derived fields | a selector for `apply_workbook_structure`; a correctness answer for the latter | 0599 |
| 16 | PPT: `SlideFactory`/`NotesIndex` still re-parse each slide through the copying entry point (the +4.7% text-only selector) | a follow-on change in `litchi-ppt` | 0606 |
| 17 | XLSX facts: the builder is 15-18% of planning; snapshot chains drop facts after the first commit; `<f>` worksheets declined | measurement, then a follow-on | 0622 |
| 18 | range sources: the whole-sheet XLS walk is request-bound (16,145 requests, 16,061 under 512 bytes); a PPT snapshot open reads 2.02× the file | a cursor read batch for range sources, designed against the fence records | 0627, 0605 |
| 19 | `SequentialTextWriter`'s two writes per row; `query_cell`'s leading fence | a follow-on in `litchi-xls` | 0621 |
| 20 | correctness: 86 order-sensitive sites in the consumer crates (three verdict-changing ones taken by 0631), `get_or_add` invalidating the source capture on a no-op reuse, the DOC facade's field-table refusals, the `office_crud_demo` PPTX update, `refine_workbook_format`, MCE output limits on the expanded stream | owners' decisions | 0628, 0587, 0607, 0588 |

Evidence gaps that remain open after this wave: cold-cache and physical-device
distributions; a real range source (only the simulated transport, now with
OLE2 coverage); peak RSS for read paths; concurrency scaling (no parallel path
is reachable from any format crate); cross-platform confirmation;
coverage-guided fuzzing on this host; DOC and PPT facade selectors and an
ordinary-save selector; a large real-producer XLSX and any real-producer XLSB.
Two gate-list gaps were found by bisects and are cheap to close: the harness's
own test suite (0619) and a feature-bearing `cargo test -p litchi` (0629) are
reachable from no standing gate.

## Limitations

Every figure is its record's, with that record's floor; the wave ran with up
to eight agents building and measuring on one 32-core host at once, so every
timing floor is the record's own and several records rest on counts where the
floor was wide. Two session usage caps terminated agents mid-work (seven agents at the first, eight and one sub-agent at the second);
each was resumed from its transcript and worktree, and no evidence was lost,
but the interruptions are the reason several records were finished in two
sittings. The coordinator merged each branch by cherry-pick in completion
order, inserted the four log paragraphs from each packet, appended its own
attribution trailer where the agent's commit carried only its model's, and ran
the fourteen in-scope crates' fmt, clippy and test gates on the merged head
(fmt clean, clippy clean, 436 test binaries all passing, exit 0, at 8b07b45dc); it did not re-run any agent's measurement.

## Retained evidence

Each change's packet under `results/change-NNNN/`; this record's
`results/change-0630/` holds the merge log, the integration-gate transcript and
`cleanup.json`.
