# 0651: the 0630 queue, worked: every measurable row landed, every contract row froze with a price, and a real deck re-priced the codec

Status: retained, coordination record. `performance_claim: none` — every number
below is quoted from the change record that measured it, with that record's own
tier and floor; nothing is re-measured or aggregated here.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## What this record is

Change [0630](0630-queue-refresh-after-the-first-wave.md) closed the first wave
and left a twenty-row queue in `HOTSPOTS.md` in which almost every row waited on
a decision, a contract or an infrastructure piece. This record closes the second
wave: the records 0632 to 0650, each written by one implementing agent in its
own worktree and branch from the base `c7326f680` (0648 from `9f28ea621`, after
0636 landed), each with its own measurement, evidence packet and four log
sections, each merged onto the main branch by the coordinator in the order it
finished. It states the outcome per queue row, quotes the numbers, records what
the wave found that no row predicted, and refreshes the ranked queue so that the
third wave starts from what is left.

The wave's rules are 0630's, with one addition: rows that wait on a human
decision or an ADR acceptance (1 to 6, 9, 11, 12 and 15 of 0630's table) were
**not dispatched**, because an agent cannot take the decision they wait on. The
wave worked the rows that needed a measurement, a follow-on change or a frozen
design, the survey's unranked low-risk items, the two gate gaps 0630 named, and,
in the two slots that freed mid-wave, the two findings of change 0638 that no row
had predicted.

## Outcome per queue row

Row numbers are 0630's; "landed" means the production change is on the main
branch with its record; "frozen" means a design record exists with its price
and admission gates and nothing was implemented; "rejected" means the record
measured the item and closed it; "not dispatched" means the row's prerequisite
is a decision this wave could not take.

| # | row | record | outcome | the number that decides it |
| ---: | --- | --- | --- | --- |
| 1 | publication audit of original part bytes (compactness contract) | — | not dispatched | waits on the ADR 0006 / 0528 decision |
| 2 | lazy OPC part decode (proposed ADR 0030) | — | not dispatched | ADR 0030 is still Proposed |
| 3 | parallel deflate of changed members (proposed ADR 0031) | — | not dispatched; re-priced by 0638 | ADR 0031 is still Proposed; on the documented ordinary save, compression is at most 0.2% to 8.7% of a lifecycle |
| 4 | MCE namespace-emission rewrite | — | not dispatched; re-priced by 0649 | the slice consumers' migration is an owner decision (0588 withdrew the rewrite); 0649 shows the frozen emission amplifies a real deck's whole-slide rewrites 16× |
| 5 | admission of `mc:Ignorable`/`dyDescent` worksheets | — | not dispatched | needs the admission-surface widening designed with 0602's D1-D4 |
| 6 | XLSX selected-cell ineligibility gate | — | not dispatched | waits on the ADR 0005 clarification |
| 7 | PPTX revision proof format | [0645](0645-pptx-memoized-revision-proof-design.md) | **frozen** with a sizing measurement (scratch implementation retained as patches) | **−26.15% Ir per opened lifecycle** with the facade carrying the memo (−8.63% without), cycles −4.9% and p50 −7.3% against ±0.55% floors; costs an `LPRM0002`/`LPCP0003` bump with a typed refusal for every previously serialized patch; `pptx_slide_remove_boundary_save` +2.45% Ir, caused by the design and disclosed; 0590's modelled 23% confirmed at a measured 26% |
| 8 | cross-copy candidate retention | [0646](0646-pptx-cross-copy-candidate-retention-design.md) | **frozen** with a sizing measurement; retaining the candidate `Snapshot` **rejected** (0598's open question answered) | deflate calls **1,112 → 556**, lifecycle Ir −11.4% media-rich; apply phase **−72.1%** and media-rich lifecycle **−35.6% p50** (floors ≤3.4%), cycles −23.1%; the price is one whole serialized package (up to 128 MiB) held in a public plan value, which needs a budget the path does not have (a breaking `opened::Limits` change) |
| 9 | fragment-only DOCX compaction | — | not dispatched | waits on the owner's whitespace answer (defect 5 of 0587) |
| 10 | DOC/PPT snapshot hashing: the ADR 0006 fence design | [0644](0644-ole2-snapshot-fence-design.md) | **frozen**; the "duplicate index parse" half **rejected**, the "third identity pass" half admitted; the `SourceVersion` substitution refused again with a runnable witness | four of the six complete reads are load-bearing; the adopted variant (drop the open's third identity call, single-scan identity inside `ensure_current`) takes the DOC open from 6 to 4 scans, **modelled −13.2%** of the DOC open (per-scan constant 2.16 cycles per byte by OLS, DOC and PPT agreeing to 0.5%), PPT unchanged; A/A floor −0.15% median; four admission gates written |
| 11 | copy-through OLE2 writer | — | not dispatched | waits on the sector-layout policy clarification |
| 12 | retained XLS sheet index | — | not dispatched | no evictable weighted cache exists in any crate; 0641 froze the snapshot-scoped chain hint behind the same gate |
| 13 | the XLS commit's second candidate parse and `Snapshot::from_bytes`'s double framing | [0633](0633-xls-commit-single-framing.md) | **landed** (the commit's second complete parse deleted); the framing passes **frozen** | commit p50 **−38.5% forward / −38.3% reverse** on `54016.xls`, floor −0.15%; allocation calls −24.8%; each framing pass is 1.6% of the open |
| 14 | framing the XLS globals once | [0633](0633-xls-commit-single-framing.md) | **frozen** with row 13's framing residue | 1.6% per pass, measured; a fusion must keep the coverage proof's order |
| 15 | XLSB reparse and derived-field refresh | — | not dispatched | needs the `apply_workbook_structure` selector and a correctness answer |
| 16 | PPT `SlideFactory`/`NotesIndex` copying entry point | [0634](0634-ppt-slide-factory-borrowed-reparse.md) | **landed** | Ir −6.1% (open + list slides), −4.9% (open + full text), −5.3% (open + slides + notes); the glibc heap-trim artifact 0606 left open disappears; paired p50 +0.61% against a 1.04% floor under glibc defaults |
| 17 | XLSX facts builder and snapshot chains | [0635](0635-xlsx-facts-builder-and-chains.md) | **landed in part**; the chain capture **rejected** and kept as a patch | fact builder **−35% to −36% Ir** on every one-edit scenario; producer-shaped edits +2.4% to +3.4% at p50 against 0.45% to 0.59% floors (below the trigger, reported) |
| 18 | range sources: cursor read batch | [0636](0636-cfb-cursor-bounded-window.md), [0648](0648-xls-shared-string-resolver-window.md) | **landed** twice (a bounded window over the CFB stream cursor, re-aimed at the validation walk; then the shared-string resolver reading the table once) | validation reads **82,727 → 60 (−99.93%)** on `54016.xls`, bytes unchanged; the whole-sheet walk **16,105 → 37 positional reads**, range-source requests 16,145 → 77, `all-cells` over a file source **−80.1% p50** (floor ≤0.87%) |
| 19 | `SequentialTextWriter`'s two writes per row; `query_cell`'s leading fence | [0641](0641-xls-scan-measure-only-cells-and-sheet-hint.md) | taken with XLS-5 and XLS-8 below; the snapshot-scoped hint **frozen** (a contract change) | see the XLS-5/XLS-8 row |
| 20a | the `office_crud_demo` PPTX update; `refine_workbook_format` | [0639](0639-gate-list-gaps-and-two-dead-paths.md) | **landed** (correctness) | the demo's PPTX update takes the route that works; `refine_workbook_format`, a public dead whole-input slurp with zero callers, deleted |
| 20b | the DOC facade's field-table refusals | [0640](0640-doc-field-table-flag-refusals.md) | **settled** (two refusals removed as defects, five confirmed correct) | `grffldEnd.fNested`/`fHasSep` restate what the `Plcfld` grammar fixes and the reader consults neither (98 end markers: `fNested` set on none of the five nested fields); the five `style names must be unique` refusals are right and `Leniency::TolerateStylesheetDefects` already admits all five; **no fixture is newly admitted** (each witness reaches a further refusal, decoded and left open); 57-fixture differential: 42 admitted on both legs, byte-identical |
| 20c | `get_or_add` invalidating the source capture on a no-op reuse | [0647](0647-opc-get-or-add-noop-reuse-design.md) | **landed** (the design measured its way into an implementation: only the vacant arm of `add_relationship` clears the capture) | the byte question 0628 froze does not exist: the compare route already copies the source member, so **6,281 of 6,281** reuse publications over 334 fixtures are byte-identical on both legs; one canonical `.rels` build per save removed; the ordinary DOCX/XLSX/PPTX/XLSB routes gain nothing today (audited) |
| 20d | the remaining order-sensitive sites | — | not dispatched | 0628's audit found no further verdict-changing site after 0631 |

Unranked survey items, gaps and mid-wave findings worked by the wave:

| item | record | outcome | the number that decides it |
| --- | --- | --- | --- |
| ZIP-6 (tail-window locate) | [0632](0632-zip-directory-prefill-locate.md) | **landed** (the central directory is read once into a buffer sized to it) | every one of 533 containers costs exactly one request and 46 bytes less (open requests 2,569 → 2,036); range-source opens **−17% to −25% at p50**, floors ≤0.3%; the 0611 differential over 22,875 inputs is byte-identical; `docx_file_source_full_text` +8.55% reported against a −7.03% B/B floor, unresolved |
| PPTX-3, PPTX-4 (eager slide catalog) | [0637](0637-pptx-eager-slide-catalog-memo.md) | **landed** | 200-slide by-index walk **−95.3% Ir**; eight `slide_count()` calls −86.9%; paired p50 not usable on that path (A/A floor +22.7%), counts decide |
| XLSX-7 (`visit_cells` streaming) | [0642](0642-xlsx-visit-cells-streaming.md) | **landed** (the range visitor no longer builds the range it visits); the scanner deliberately not made to yield (an ADR 0006 refusal-order change, left for a design record) | warm whole-sheet visit **65,551 → 0 allocations**, **−81.7% Ir**, p50 **6.1×** faster (floors −3.9% to +0.2%); cold reads +0.19% Ir and unchanged peak live bytes, reported; 397-worksheet four-way differential byte-identical |
| XLS-5, XLS-8 (measure-only cells, sheet hint) | [0641](0641-xls-scan-measure-only-cells-and-sheet-hint.md) | **landed** (measure-only validation of a query's non-target cells; the per-sheet cursor's chain walk resumed instead of restarted); the snapshot-scoped hint **frozen** | queries **−28% to −41% cycles** (floors within ±2.2%), p50 −11% to −40% in a quiet window (floor ±1.6%); multi-sheet text **−6.14% chain links** over 96 fixtures; `15228.xls` full text **+0.84% cycles**, chased to outlined drop glue and reported; a census of 127,072 cell values corrects the survey's `drop_in_place` reading (`54016.xls` has no `Formula` or `Label`); `MulRk`/`MulBlank` (47.7% of values) not covered |
| XLS-10 and 0636's follow-on (shared-string resolver window) | [0648](0648-xls-shared-string-resolver-window.md) | **landed** (the resolver retains the whole string table for one scan after eight resolves); the brief's sliding window implemented first and **rejected** on bytes | `54016.xls` whole-sheet walk **16,105 → 37 positional reads**, cycles −67.9%; corpus-wide `all-cells` reads −96.4% with no fixture reading more times; open unchanged (+0.10% cycles) |
| 0592's open question and the DOCX sink text (XML-6) | [0643](0643-docx-paragraph-count-and-sink-text.md) | **landed** (the sink text parser borrows its events); 0592's regression **attributed**, not fixed | sink read allocations **−88.8%**, p50 **−15.6%** (eager, floors ≤0.6%), −12.4% on a 16.79 MB real file; `docx_file_eager_paragraph_count` +9.81% reproduces with **−0.53% instructions** and decomposes into +0.85% code layout plus about 5.7% of warm-up the harness's untimed preparation used to supply |
| evidence gap 5 (facade `.doc`/`.ppt` and ordinary-save selectors) | [0638](0638-facade-and-ordinary-save-selectors.md) | **landed** (harness only, thirty opt-in selectors) | the documented save's atomic publication is **7.7× to 61.7×** its serialization; a one-element edit regenerates 0.45% of a real XLSX's payload bytes; A/A p50 2.09% over 39 rows |
| gate-list gaps (harness suite, feature-bearing facade tests) | [0639](0639-gate-list-gaps-and-two-dead-paths.md) | **landed** | two new `non_iwork_gate.py` modes CI runs |
| 0638's unexplained 133.61 ms real-deck edit (dispatched mid-wave) | [0649](0649-pptx-opened-transaction-real-deck-edit.md) | **attributed**, design frozen, nothing implemented | **six whole-slide MCE rewrites per edit** (three per capture, two captures), 1.73 MB in → 28.06 MB out (**16.25×**, 259× the archive), 373,086 namespace re-declarations; a marker-stripped control removes **93.9%** of the edit; the revision proof is **0.51%**; the generated corpus is cheaper by 77× because none of its members carries the MCE namespace |
| 0638's DOCX editor refusal of 0593's fixture (dispatched mid-wave) | [0650](0650-docx-editor-byte-order-mark-admission.md) | **fixed** (a defect: the editor's preserved body ranges ran three bytes low on a byte-order-marked main part); four follow-ups **frozen** with witnesses | 63-fixture census: `open` admits 55, `document_mut()` refused 2 of them; the two-leg census differs in **3 lines**, all refusal text on the one marked fixture, every published digest identical; the fixture itself is still refused, now for the body-final `w:sectPr` placement the managed route already reported; the same offset defect exists in `litchi-opc`'s publication audit and the managed DOCX transaction (reported, not touched) |

## Correctness work the wave produced

Reported as defects and fixed: the `office_crud_demo` example's PPTX update
took a route that cannot succeed, and `refine_workbook_format` was a public,
dead, whole-input slurp with no caller ([0639](0639-gate-list-gaps-and-two-dead-paths.md));
two of the seven DOC refusals 0587 called ordinary-file refusals read a
redundant `grffldEnd` bit as structure, and are removed, while the five
stylesheet refusals are confirmed correct ([0640](0640-doc-field-table-flag-refusals.md),
which also corrects 0587: `watermark.doc` is a LibreOffice export and the other
witness's producer is unidentifiable); the DOCX editor refused a byte-order-marked
main document because quick-xml strips the mark from its slice without counting
it in `buffer_position`, so every preserved body range ran three bytes low
([0650](0650-docx-editor-byte-order-mark-admission.md)); a relationship call that
establishes nothing dropped 0593's pristine proof for its `.rels` member, which
0628 had left frozen on a byte question that does not exist
([0647](0647-opc-get-or-add-noop-reuse-design.md)). Three of 0623's absolute
request-count assertions moved by exactly one when 0632 removed a request
(4/10/14 to 3/9/13), and 0642 made the scanner's refusal-before-visit guarantee
structural rather than an argument about reachable arms.

Reported and left for their owners, each with a witness in its packet: the same
byte-order-mark offset defect in `litchi-opc`'s publication audit (a regenerated
marked part is refused at publication) and in the managed DOCX transaction, the
body-final `w:sectPr` placement and ISO Strict `pt` page sizes that keep two
reader-admitted DOCX fixtures out of the editor (0650); the two DOC refusals
0640 unmasked (`FBKF.ibkl` uniqueness, a `PapxInFkp` pad byte), each
provisionally stricter than the format, and the `litchi` facade's `.doc` route
that cannot reach the lenient stylesheet path (0640); two XLSB resource guards
that test the part rather than the relationship (0647); seven more
`read_event_into` sites in `litchi-docx` with the shape 0643 fixed, one on a
save path (0643); the scanner's own record vector that sets a cold XLSX read's
peak and cannot yield without moving a refusal (0642).

## What the wave found that no row predicted

- **The documented ordinary save is publication-bound, not serialization-bound**
  (0638): a compression or layout change on that route is worth at most 0.2% to
  8.7% of a lifecycle, which re-prices row 3 (parallel deflate) for the ordinary
  route before ADR 0031 is decided.
- **The MCE codec's frozen namespace emission is 94% of a real PPTX edit** (0649).
  0638 measured one shape-text edit on a real 108 KB deck at 133.61 ms, 490× the
  deck's serialization; 0649 attributed it to six whole-slide rewrites per edit,
  each amplified 16× by the per-element re-declaration 0588 measured and
  withdrew, with the revision proof at 0.51%. Every prior PPTX record measured on
  generated corpora whose members never carry the MCE namespace, so the codec's
  cheap branch was the only one ever priced on an edit. Row 4 is therefore the
  top item of the refreshed queue, and the harness needs a marker-bearing PPTX
  corpus (0601 built one for XLSX only).
- **A count of removed reads is still not a saving until bytes are counted**
  (0648): the sliding window the brief proposed cut the whole-sheet walk from
  16,105 to 5,539 reads while taking 1,057× the bytes, because the access order
  has almost no spatial locality; reading the table once was better on both axes.
- **The survey's `drop_in_place` reading was a type-size artifact** (0641):
  `54016.xls`, the fixture it priced the 2.05% on, has no `Formula` or `Label`
  record at all; the cost is the 88-byte enum's drop glue, which is why the fix
  covers all seven single-cell kinds and why `15228.xls` joined the measured set.
- **Two frozen designs cost more than they save without an owner's decision**
  (0645, 0646): both are worth a quarter to a third of their lifecycles, both
  are measured with retained scratch implementations, and both wait on an ADR
  0005 answer about state they would keep (a facade-carried memo; a retained
  serialized package) rather than on any measurement.
- **The `release` profile's `lto = true` makes the measured binaries
  non-reproducible byte for byte** (0635), so a packet's binary digest names the
  binary that was timed, not a rebuild of the committed source.
- **`litchi-perf-baseline-alloc` emits no allocation metrics for 0638's
  ordinary-save family** (0649): a harness gap for the next wave.

## The three stale facade tests the rebase brings in

Before the rebase this wave ends with, the coordinator ran the facade's
`docx,odt` feature combination, which no standing gate compiles (the
`facade-format-tests` mode 0639 added runs `docx,xlsx,pptx,xls`). Three tests
are stale on this branch and fail with the behaviour the code already has:
`owned_odt_bytes_keep_odt_owner_with_malformed_ooxml_catalog` and
`filesystem_odt_keeps_source_owner_with_malformed_ooxml_catalog` expect a
fallback to ODT that the catalog reader refuses (`Invalid [Content_Types].xml
manifest: root must be Types in the OPC content-types namespace`), and
`ordinary_odt_uses_native_policy_even_with_ooxml_suffix_but_polyglot_honors_docx_input_limit`
expects a DOCX-suffixed ODT to bypass the caller's input limit, which it no
longer does. The upstream branch's commit `cb2a1a2d4` rewrites all three to the
current contract; the rebase takes those rewrites, the coordinator's log of the
failing run is retained in this record's packet (`odt-polyglot-local.log`), and
the gate gap is row 19 of the refreshed queue. The same upstream commit repairs
the managed exact-no-op budget tests in `litchi-xlsx` with a 64 KiB headroom
this branch had already added under another name; the rebase keeps this
branch's spelling, since the values are identical.

## The queue after this wave

The queue is now, even more than after 0630, a list of decisions: the first ten
rows each wait on an owner, an ADR acceptance or an ADR clarification, and the
records named have already priced what each decision unblocks. Rows 11 to 20 are
the small follow-ons and gaps the wave itself named, none of which needs a
decision. The same table replaces the "Ranked work queue" in `HOTSPOTS.md`.

| # | item | what it needs first | priced by |
| ---: | --- | --- | --- |
| 1 | the MCE codec's namespace re-declaration, now known to amplify every whole-slide rewrite of a real deck 16×: a marker-stripped control removes 93.9% of a 133.61 ms opened-transaction shape-text edit, and every prior PPTX record measured on corpora whose members never carry the namespace | an owner decision that the slice consumers (`Paragraph::extensions`, `Shape::xml`, five XLSX raw accessors, two publishing writers) may stop seeing per-element re-declarations (0588 withdrew the rewrite on that contract), then the migration design | 0649, 0588, 0630 row 4 |
| 2 | the publication audit of original part bytes refuses 94 of 95 real OOXML packages; the same contract is 27.6% of source-backed publication instructions | a human decision on the compactness contract for original bytes (ADR 0006, record 0528); it gates every real-producer source-backed save | 0602, 0616, 0528 |
| 3 | lazy OPC part decode (C2′): an eager XLSX open-edit-save inflates 90 parts to read one | acceptance of proposed ADR 0030, then a migration of 259 sites | 0610, 0581 |
| 4 | the PPTX memoized revision proof: −26% of an opened lifecycle's instructions, −7.3% at p50, at the price of a durable `LPRM0002`/`LPCP0003` bump with a typed refusal for every serialized patch, and a memo the facade must carry | an owner decision on the format bump and on the facade-carried memo under ADR 0005; the nine admission gates are written | 0645, 0590 |
| 5 | the cross-package copy's second candidate serialization: deflate calls halve, the apply phase −72%, the media-rich lifecycle −36% at p50, at the price of one whole serialized package retained in a public plan value | a budget for the retained candidate (a breaking `opened::Limits` field, or ADR 0005's hierarchical budget which this path lacks); fourteen gates are written | 0646, 0598 |
| 6 | parallel deflate of changed members: 24-62% of a save's cycles on the in-memory routes, but at most 0.2-8.7% of the documented ordinary save, which is publication-bound | acceptance of proposed ADR 0031, then the measured balance rule; and an ADR clarification of what durability the documented save promises, since its atomic publication costs 7.7× to 61.7× its serialization | 0624, 0615, 0638 |
| 7 | admission of `mc:Ignorable`/`dyDescent` worksheets to the fused traversal, and the value-only vocabulary that refuses every real worksheet | an admission-surface widening designed with 0602's D1-D4 | 0603, 0602, 0601 |
| 8 | the XLSX selected-cell ineligibility gate: −63% of an ineligible read | the limit question and an ADR 0005 clarification on which reader owns a lazily loaded payload's refusal, or an observer-detach signal in `litchi-ooxml-common` | 0597 |
| 9 | the DOC snapshot's source-identity fence: four of its six complete reads are load-bearing, a `SourceVersion` replaces none of them, and the adopted variant (the open's third identity call dropped, one scan per identity call inside `ensure_current`) is modelled at −13.2% of the DOC open, PPT unchanged | an implementing change under 0644's four admission gates (the before halves of two are in its packet); the cheaper variants cost a second public `litchi-cfb` entry point to keep one typed error's name | 0644, 0589, 0609 |
| 10 | fragment-only DOCX compaction and the main-part copy in `document_snapshot`; a copy-through OLE2 writer for DOC saves (13-30% of a DOC save) | the owner's answer on whitespace compaction across untouched paragraphs (defect 5 of 0587); an ADR clarification of physical sector layout policy | 0591, 0617 |
| 11 | XLS cross-query retention: the retained sheet index, and the snapshot-scoped chain hint 0641 froze (a few thousand dependent loads per repeated query) | an evictable weighted cache under ADR 0005's clean-value-cache gate, which no crate has; `StreamChainHint` borrows its reader, so a snapshot-resident position needs a reader-independent value or a lock | 0605, 0641 |
| 12 | XLS residues: `MulRk`/`MulBlank` cells (47.7% of the corpus's values) still build an 88-byte record each on an unwanted query; the open still scans the whole string table to build its locators; the two framing passes at 1.6% each | follow-ons in `litchi-xls`; the framing fusion must keep the coverage proof's order | 0641, 0648, 0633 |
| 13 | XLSB: `insert_candidate_cell`/`transfer_cell` reparse twice; `apply_sparklines`/`apply_cell_watches` never refresh derived fields; two resource guards test the part rather than the relationship | a selector for `apply_workbook_structure`; a correctness answer for the latter two | 0599, 0647 |
| 14 | DOCX residues: seven more `read_event_into` sites with the shape 0643 fixed (one on the `tail_append` save path); one bounded `String` per paragraph on the sink path; a memoized paragraph index across `document()` views (a `DocumentIndexAdmission` charge on the managed route) | pricing, then a follow-on; the memo needs a frozen design | 0643, 0592 |
| 15 | XLSX residues: the scanner's own `Vec<SelectedRecord>` sets a cold whole-sheet read's peak and cannot yield without moving a refusal past a partial result; `cells` over-allocates 42% on the stored route | an ADR 0006 refusal-order design for the first; a follow-on for the second | 0642 |
| 16 | correctness follow-ups: the byte-order-mark offset defect 0650 fixed in the DOCX editor also lives in `litchi-opc`'s publication audit (a regenerated marked part is refused at publication) and in the managed DOCX transaction; body-final `w:sectPr` placement and ISO Strict `pt` page sizes keep two reader-admitted fixtures out of the editor; the two DOC refusals 0640 unmasked (`FBKF.ibkl` uniqueness, a `PapxInFkp` pad byte), each provisionally stricter than the format; the `litchi` facade's `.doc` route cannot reach `Leniency::TolerateStylesheetDefects` | two offset fixes with 0650's witnesses; an ECMA-376 `CT_Body` decision; a spec reading for the two DOC refusals; a facade option | 0650, 0640 |
| 17 | harness gaps: no marker-bearing real-producer PPTX or DOCX corpus (every prior PPTX record measured on the codec's cheap branch); `litchi-perf-baseline-alloc` emits no allocation metrics for the ordinary-save family; no DOCX text-sink selector; no PPTX opened-transaction phase timer | harness work in `tools/perf-baseline`, modelled on 0601 and 0638 | 0649, 0643, 0638 |
| 18 | the ZIP locator's own 64 KiB scratch (a two-stage locate costs one extra request on the 1 of 533 containers that misses the `len − 22` fast path) | a measurement | 0632 |
| 19 | gates: the facade's `docx,odt` feature combination compiles three polyglot-detection tests no standing gate runs (stale on this branch until the upstream rewrite the wave's rebase took); `--mode structural` of `check_perf_claims.py` fails on landed `claim-0251`; two flaky allocator tests in the harness; a stale `tools/native-resave/Cargo.lock` | small gate and hygiene changes | 0651, 0641, 0642 |
| 20 | method: `lto = true` makes the release binaries non-reproducible byte for byte, so a packet's binary digest names the binary timed, not a rebuild of the source | a note in the harness README, or a reproducible measurement profile | 0635 |

## Limitations

Every figure is its record's, with that record's floor; the wave ran with up
to eight agents building and measuring on one 32-core host at once, so every
timing floor is the record's own and several records rest on counts where the
floor was wide (0637 and 0646 say so explicitly). One session usage cap
terminated 0644 (after its commit, during its own verification pass), 0650
(mid-census) and one reviewer sub-agent; both were resumed from their
transcripts and worktrees after the reset, no evidence was lost, and 0648's
report had already been delivered. The coordinator merged each branch by
cherry-pick in completion order (the merge log is in the packet), inserted the
four log paragraphs from each packet, appended its own attribution trailer, and
resolved one conflict: 0648 and 0641 each added names to the same import list
in `crates/litchi-xls/src/workbook/source.rs`, resolved by union, rustfmt clean,
and the crate's suite re-run on the merged head (`xls-merged-0641-0648.log`, exit
0). The integration gate on the merged head at `f0fc20178` ran the fourteen
in-scope crates' fmt, clippy, test and doc gates plus 0639's two CI modes and
the four documentation checkers: fmt clean, clippy clean, 439 test binaries all passing, rustdoc clean, both CI modes green, 10 claims validated in strict mode, 167 report rows classified, the coverage index and the non-iWork gate verified, every step exit 0 (08:23 to 08:43 UTC). The coordinator did
not re-run any agent's measurement.

Pre-existing findings agents reproduced on the untouched base, none fixed here:
`check_perf_claims.py --mode structural` fails on landed
`claim-0251-xlsx-xml-borrowed` (CI's strict mode passes; 0643, 0646, 0641);
`tools/test_perf_claims.py`, `test_check_crate_boundaries.py` and
`test_native_odf_resave.py` fail as 0627 recorded (0638); two flaky
process-global allocator tests in the harness's `docx_bounded_tail_append_compare`
(0641, six runs); six clippy warnings in the `litchi` facade under
`--features docx,xlsx,pptx,xls` (0640, 0642); a stale `tools/native-resave/Cargo.lock`
that any cargo run there regenerates (0642, 0650).

This record is committed before the branch is rebased onto the upstream
`feat/office-format-completeness`, so the commit hashes the records 0632 to 0650
cite (their base `c7326f680`, their branch commits and their merged commits) are
pre-rebase hashes. A follow-up commit after the rebase adds
`results/change-0651/rebase-commit-map.txt`, mapping every pre-rebase hash of
this branch to its post-rebase hash, so the citations stay resolvable.

## Retained evidence

Each change's packet under `results/change-NNNN/`; this record's
[`results/change-0651/`](results/change-0651/README.md) holds the merge log,
the integration-gate transcript, the facade `docx,odt` test log, the merged
`litchi-xls` suite log, `decision.json`, `cleanup.json` and, after the rebase,
the commit map.
