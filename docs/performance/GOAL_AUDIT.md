# Non-iWork `docs/GOAL.md` audit

## 0636 — a range-source entry point priced and fixed, and a queue item re-aimed

`docs/GOAL.md` names caller-supplied remote and range sources as a benchmarked
dimension, and change 0627 opened that axis for OLE2. This change is the first
production change measured on it, and it moves a P1 row rather than closing one:
`litchi_xls::validation::validate_source` takes an `Arc<dyn ReadAt>` and was
costing 82,727 positional requests on a 984,576-byte workbook — modelled at 82.7
seconds of fixed service at change 0627's transport — and now costs 60,
modelled at 69 ms. It also corrects the audit's working picture of where the XLS
range-source cost sits: the whole-sheet walk 0627 flagged is dominated by
per-shared-string cursor constructions at random offsets, not by the framing
loop, so the "finish source-backed CRUD adoption across formats" row still owns
that work and it is a different mechanism. What this change does **not** supply
is an observed range-source measurement for validation: no selector exists for
it, and building one before the change would have cost 82.7 s per sample, so the
counts carry the argument and the 0627 model only prices it. The read path is
unchanged on every axis measured — identical reads, bytes and observations over
113 fixtures, and identical request-sequence digests over 0627's five
range-source selectors — so no existing coverage row changes status. OLE2 and
OOXML remain the active priority; ODF stays deferred and iWork excluded.
[Change and limitations](0636-cfb-cursor-bounded-window.md);
[evidence](results/change-0636/README.md).

## 0635 — queue item 17 and survey item XLSX-5 closed, one by code and one by evidence

Record: [0635](0635-xlsx-facts-builder-and-chains.md).

Change 0635 closes queue item 17 of change 0630 and item XLSX-5 of the 0587 survey.
Item 17 had three parts and they end differently. The builder's cost is **reduced**:
the builder falls **35.0-36.2%** on every one of eight scenarios (`FactsBuilder::element` −41.8 to −43.4%, `FactsBuilder::cell` −36.9 to −37.7%, `raw_attribute` −24.6 to −27.4%), planning falls **4.29-5.26%**, and the whole harness iteration falls **0.13-2.19%** — the largest on the 0601 producer-shaped edit selectors, which are not dominated by the cell-CRUD corpora's 4 MB of media. The builder's share of planning goes from 13.2-14.3% to 8.9-9.7%. No scenario got worse, and every admitted and declined worksheet is unchanged. The
`<f>` decline is **unchanged** and still a design narrowing, not a measured
rejection. The snapshot chain is **rejected on evidence**: the capture was
implemented, proved with change 0622's oracle extended to two-commit chains, and then
measured at **+0.20% to +4.06% of a whole save**, the largest on `producer-dense` and within sight of the 5% review trigger — and no public API seeds a value-only edit
from a post-commit snapshot, so the facts it would carry are read by nobody.
`SourceBackedEditor::edit` and `edit_sheets` always load a fresh snapshot from the
immutable package; `MultiSourceEdit::new` is private; `publish_multi_commit_to_stream`
consumes the editor; the only way to chain two cumulative value-only commits is
publish-then-reopen, which re-plans. The implementation and its four oracle tests are
retained as a patch so a future change that exposes a chained edit can land them.
XLSX-5 is closed twice over: the traversal's per-event copies are removed
(−80 allocation calls, −6,457 allocated bytes per planning, exactly), and the
alternative the survey offered — substituting the existing `styles::stream_count` —
is **declined and frozen as a design**, because `process_ooxml` short-circuits on
MCE-free input while the stream path does not, so the two disagree on trailing text,
DTDs, processing instructions, custom entities, a late declaration, unbound prefixes,
depth, event count and several stream-only bounds, and the error type and message
would move as well. The optimization order in `docs/GOAL.md` is respected throughout:
this removes unnecessary work, parsing and allocation, ahead of layout, algorithms
and parallelism; nothing was vectorized and no parallelism was introduced. **The
audit gap this change does not close** is the one change 0622 recorded: the route it
optimizes is unreachable on real producer files, and this change neither widens nor
narrows that surface.

## 0634 — PPT per-slide re-parse borrows the retained stream

Step 2 of the optimization order (unnecessary copying and allocation) applied to the last copying path inside the PPT reader that change 0606 named in its own "Limitations" and change 0630 carried as queue item 16. `Presentation::slides`, `slide_at`, `text` and `extract_text_fast`, the `NotesIndex` whole-document re-parse behind them, `SpeakerNotes` and the source-backed editor's slide resolve now all parse from the `PowerPoint Document` stream the presentation already retains instead of copying each record's payload out of it: 160 fewer allocations and 160 fewer `memcpy` calls per `slides()` on `45543.ppt`, 98 per speaker-notes read on `headers_footers_2007.ppt`, and 561,052 → 416,036 retained live bytes for an open plus slide list. No P1 audit row closes: this is one format's reader, the source-backed PPT path still pays the whole-artifact SHA-256 that 0587 ranked first, the editor paths still copy through the strict entry point, and no RSS, cold-cache, physical-device or producer-corpus measurement was taken. Lossless preservation, typed refusals and every record limit are unchanged — `max_copied_payload_bytes` is still charged per record although nothing is copied — and all 30 `.ppt` fixture reader dumps, including the four encrypted refusals, hash to the digest change 0606 recorded. OLE2/OOXML optimization remains active; ODF is deferred until completion and iWork excluded. [Change and limitations](0634-ppt-slide-factory-borrowed-reparse.md); [retained evidence](results/change-0634/README.md).

## 0630 — the queue is worked, and the goal's next obstacles are decisions, not code

`docs/GOAL.md` asks that the largest measured bottlenecks be addressed in Amdahl and return-on-investment order. Change [0630](0630-queue-refresh-after-the-first-wave.md) records that the 36-item queue of change 0587 has been worked to the end: 25 items landed with paired measurements, 11 were frozen, declined or falsified on measurement, and every one of the survey's own errors that the work exposed is annotated beside the item it corrects. Three goal clauses moved. "Selective reads perform work proportional to accessed content" is met further on OOXML (a member's first read is one request, the open's requests halved) and on XLS (one validated scan per sheet, text observations once per read). "Targeted updates avoid parsing unrelated content" moved on DOCX (two scans, not four), XLSX (the commit's second whole-sheet scan is gone), XLSB (no reparse on a no-op) and XLS (one fewer complete parse per commit). "Unchanged members flow without unnecessary work" moved at publication (unchanged relationships and content types are copied, not reserialized and audited). Two methodological rules join the standing ones: a count of removed work is not a saving until the cost it trades into is measured natively (0604, 0612), and a per-crate gate list cannot see the harness's own suite or the facade's feature-gated tests (0619, 0629). Still required and unchanged: cold-cache and physical-device distributions, a real range source, peak RSS for read paths, concurrency scaling (no parallel path is reachable from a format crate, 0615), cross-platform confirmation and coverage-guided fuzzing. The next obstacles are decisions: the original-bytes compactness contract at publication (0602, 0613), proposed ADRs 0030 and 0031, the sector-layout clarification for a DOC copy-through writer (0617), and the whitespace-compaction question (0591). OLE2/OOXML stay first; ODF is deferred until that goal completes and iWork is excluded; the broad goal remains active.

## 0623: one read per contiguous run of structural members — the OOXML open's request count falls again, for no extra bytes

Record: [0623](0623-zip-structural-span-accessor-and-prefetch.md).

Change 0623 is the first change in this programme to land a frozen design
written by an earlier record without amending it, and the reason is worth
recording: 0577 froze not just the mechanism but the *invariants*, and two of
them turned out to do all the work. Invariant 1 — "every byte read lies inside a
structural member's own local record" — was written before change 0611 existed,
and after 0611 it has a sharper form that can be *computed* rather than argued:
join two members into a run only when the earlier member's own first-read span
already reaches the later member's local header. That single rule makes the run
the union of spans the archive already reads, which is why the measured byte
totals are identical rather than 15% higher as 0577 modelled, and why no
container in the corpus reads a byte the before leg did not. Invariant 5 — "a
run read that fails falls back to per-member reads" — is the invariant change
0611 explicitly *refused* for its own one-member read, and both records are
right: a read that is a member's own first contact with the source must not be
retried, and a read that belongs to no member must be abandonable. A design
record that states its invariants precisely enough to be contradicted by a later
change is more useful than one that states a mechanism.

The second lesson is about where an optimization is allowed to apply. This one
is gated to the ordinary unmanaged, exact-policy open. A speculative read
spanning several members is reserved and committed against
`Resource::InputBytes` and observed for cancellation as a unit, so under an
`ExecutionContext` it would move where those observations land — and change
0611's justification for accepting that class of movement (its span read *was*
the member read, one for one) is not available here. Declining to apply a saving
on the managed path is a cheaper answer than sixteen corrected test triggers,
and it is the honest one: the mechanism genuinely is not invisible there.

## 0631 — "exact no-ops stay exact" restored in three crates, and the limits of what the corpus can prove

Record: [0631](0631-ooxml-relationship-order-verdict-sites.md).

`docs/GOAL.md` puts correctness and lossless preservation above speed and states that exact no-ops stay exact; ADR 0006 states that *"Serialization is deterministic unless a `Clock`, actor identity, or cryptographic RNG is explicitly supplied"*; ADR 0003 states that a source-checked patch conflicts on a real overlap and that *"there is no last-writer-wins behavior"*. Two of the three defects this change fixes broke all three sentences at once: a patch that staged **nothing** — no cell edit in XLSX, no content-control edit in DOCX — captured against a document and published against the same document, was refused whenever the two loads happened to draw opposite `HashMap` iteration orders for two relationships. The third produced a public `u64` that was a different number on every run. Each fix sorts on the key its own neighbours already use, so the rule reads the same in all three crates: *the relationships in the order the published `.rels` member lists them*. Because the sort key is the map's own unique key, a collection with zero or one matching relationship keeps exactly its old bytes, revision and verdict, and three of the ten new tests exist only to pin that on both legs. Two audit rows this change opens rather than closes, and they are the honest limits of the corpus evidence. First, **no fixture in this repository reaches the XLSX site at all**: every real workbook carries core-properties and extended-properties package relationships, and `validate_package_relationships` refuses the value-only closure for any package relationship that is not the officeDocument owner or a signature origin, so all 180 XLSX fixtures stop before `capture_auxiliary` runs; likewise exactly one of 78 PPTX fixtures carries `ActiveX` controls and none of its parts owns two relationships. Their proof is the new tests and the crate suites, and nothing is claimed about how common either shape is in the wild. Second, the **message-only** ordering class 0628 left alone is still open and is now measured rather than assumed: `validate_package_relationships` and its peers name whichever offending relationship they meet first, which makes 173–176 of 180 XLSX fixtures report a different diagnostic string between runs — **on both legs**, with a stable error variant. The one place this change happens to fix it is the XLSX auxiliary walk, because ordering the relationships also orders which of two bad auxiliary parts is reported. The wider 0628 audit — public list-order variation in the PPTX chart, slide-master, comments, modern-comments and tracks loaders and the XLSX pivot, chart, shapes and slicer loaders; the positional pairing hazard in `smartart.rs`; the ambiguous `find` selections in PPTX theme-override and XLSX `connections::remove_from_package`; and the two PPTX copy planners whose public refusal *discriminant* flips — is untouched: the brief named three sites and this change fixed three. No latency, RSS or throughput improvement is claimed. OLE2/OOXML remain active; ODF is deferred until completion and iWork excluded.

## 0611: one bounded positional read per ZIP member first-read — a source-backed OOXML open halves its requests, and halves its wall clock on a latency-bearing transport

Record: [0611](0611-zip-single-read-per-member.md).

Change 0611 is the first change in this programme to price a read-grammar change
on a latency-bearing transport rather than on a warm local source, and it is
worth recording why that was the right instrument. Change 0561 attributed this
span model and declined to implement it because change 0493 had already measured
the same mechanism, as an opt-in read-ahead window, at ±2% on a warm local
source and −85% at 1 ms per request; its conclusion was that collapsing these
reads is worth nothing locally and must be argued against a range-source
measurement. This change is the measurement 0561 asked for, and it confirms both
halves: on local, in-process sources every selector sits inside the A/A floor and
nothing is claimed, while on the simulated transport a source-backed OPC open
halves its wall clock. The standing instruction that instruction counts rank
work and not latency has a counterpart here — **request counts rank latency on a
transport and not work**; the deterministic count is the durable result and the
timing is its price at one service time. A third lesson is about *test triggers* rather than assertions: sixteen tests in
two consumer crates identified "the read that fetches this member" by its start
offset, and every one of them broke without a single contract moving. Three
successive predicates were needed before one held under both grammars, and each
wrong one was found by a different test — a zero-length terminating read landing
on the next member's header, a preservation probe at the same offset as the
member read, and a bulk publication copy spanning the member from 64 KiB away.
A read-grammar change should expect to spend as much effort on triggers as on
the code.

The second lesson is about oracles. The
differential change 0582 built, the gate 0587 names for any read-grammar change,
exercises only the strict-layout APIs and never calls `IndexedArchive::read_entry`
— the path every ordinary OPC part read takes. Re-running it unchanged would
have proved nothing about this change. It had to be extended before it could
gate, and once extended it immediately found a real divergence that no test in
either crate caught.

## 0629 — a fence the program wanted, a test the gate list could not see

Record: [0629](0629-facade-docx-budget-test-bisect.md).

`docs/GOAL.md` puts correctness and bounded resources above speed, and this record is an instance of the program paying that bill correctly and then failing to notice the bill. Change 0495 added a conservative parser workspace to the managed source-backed DOCX read path — `xml_len * 32 + 131_072` bytes of memory, plus objects, depth and an `(len + 1)^2` work ceiling, reserved before quick-xml sees caller-supplied XML. That is exactly the direction ADR 0005 and the optimization order ask for, and the 0495 record describes it in prose. What 0495 could not see is that one test in a different crate had sized a budget to the input length and would now always refuse. The audit's standing rule that a failure be attributed before it is acted on is met here by measurement: eleven bisect verdicts place the first bad commit at `44a471069`, its parent passes, and three `--release` legs confirm the verdict outside the `test` profile, so **changes 0588, 0591, 0592, 0593 and 0594 owe no correction** and change 0621's pre-existing call is confirmed rather than overturned. The structural cause is worth the audit's attention and is the same one change [0619](0619-harness-xls-lifecycle-assertion.md) found one layer out: there, `tools/perf-baseline` is a separate Cargo project a workspace `cargo test` never reaches; here, `litchi`'s `default = []` means `cargo test -p litchi` compiles almost none of the facade's own tests, so a crate can pass every gate in the standing list while the unified facade that fronts it is red. A feature-bearing `cargo test -p litchi` belongs in the standing gate list. Evidence tiers: **measured** for all eleven bisect verdicts, the three release legs, both feature-bearing suite runs and every byte of the arithmetic (194, 1,154, 137,280, 137,474); **modelled**: nothing; **unknown**: whether the 131,072-byte floor is right for a small document — this record does not measure quick-xml's real peak working set and proposes nothing about it. What this batch does not discharge: the fence is deliberately not weakened, because shrinking it would move where a refusal happens and needs a frozen design record; no sweep was made for other budgets in the tree still sized to a pre-`44a471069` number; and the five wave changes are cleared for this test only, each remaining on its own evidence. `performance_claim: none`; no library code path changed. OLE2/OOXML remain the active priority; ODF stays deferred; iWork is excluded.

## 0627 — a benchmarked dimension `docs/GOAL.md` names for both priority formats had no OLE2 measurement at all

Record: [0627](0627-ole2-range-source-selectors.md).

`docs/GOAL.md` names caller-supplied remote and range sources as a benchmarked dimension for OLE2 *and* OOXML. Change 0587's evidence gap table recorded that only one half of that had ever been measured: `SimulatedRangeSource` over OPC and XLSX, a dedicated `pptx_range_source` module and provider pacing on the OOXML side, and not one CFB, XLS, DOC or PPT case on the other. **The row closes for XLS and PPT and stays open for DOC and for CFB itself**, which is the honest reading of a gap that names four formats. The audit's own standard — that a claim be scoped to scenario, corpus, machine, build and metric — is what shapes the result: request counts, request bytes and request-size distributions are **measured** and identical across 50 retained samples of both transports; the request *sequence* is proved a pure function of fixture and scenario by a per-sample digest, which is change 0572's determinism property applied per selector; elapsed time on the range leg is **modelled** and labelled so, because the configured service floor is 86.7–94.8% of the median and, per changes 0447 and 0448, the delay counters are requested targets and not observed sleeping time; the corpus-wide rate of the two typed refusals met here is **unknown**, and deliberately left to change 0587's evidence gap 3 rather than guessed at. Two audit-relevant facts the record produces without proposing anything. First, the readers' admitted populations differ by scenario on real files: `ConditionalFormattingSamples.xls` opens, lists, answers a cell query and walks a worksheet but refuses full text, and `45543.ppt` opens but refuses the text of a shape its own eager reader reports with text — with four further real decks spot-checked refusing the same way, while `litchi-ppt`'s test for that API uses a generated fixture. Second, the audit's standing worry that "the measured path is not the path real files take" now has a measured instance in the opposite direction: a `SourceSnapshot::open` reads 2.02× its file before answering anything, which no in-memory selector had made visible. What this batch does not discharge: no DOC or CFB selector; one transport point; in-memory sources only, so the `fstat` per source observation on the `from_path` route that change 0587 identified is still unmeasured; and no save, edit or publication scenario runs over a range source. `performance_claim: none`; no file under `crates/` changed. OLE2/OOXML remain the active priority; ODF stays deferred; iWork is excluded.

## 0626 — deliverable 8's smoke check starts comparing against history, and stays off the merge gate on purpose

Record: [0626](0626-perf-ci-smoke-baseline-fetch.md).

`docs/GOAL.md` deliverable 8 asks for "a lightweight, stable performance smoke
check suitable for CI, plus a fuller manually triggered or scheduled benchmark
workflow", and adds "do not make noisy cloud-hosted microbenchmarks a hard merge
gate until variance is understood". Both halves existed; the first half was
comparing today's report with itself. It now compares against the last
successful `full` run's artifact when one is bound to the same policy identity,
corpus key manifest digest, harness binary profile and runner labels, which is
the first time any push or pull request in this program has been measured
against prior history at all. The second half is honoured structurally rather
than by assertion: the comparison runs under the allocator policy, so
`perf_compare` reports `latency_claims: withheld_instrumentation` and compares
zero latency results, and the twenty metrics it does compare are deterministic
allocation counters; on top of that `enforcement: advisory` means a regression
or an unusable comparison annotates the run and does not fail it. The audit rows
this leaves open are named in the record: nobody has yet observed a GitHub
Actions run of this workflow, because neither `gh` nor `actionlint` is installed
on the program host, so the fetch step's shell has never executed; whether a
hosted runner's build identity is stable enough between runs for the fetched
mode to fire in practice is exactly the variance deliverable 8 says must be
understood, and it is unknown; and the reference is one case over one pinned
corpus, not the 201-result release matrix a pull-request job cannot afford.
Deliverable 6's machine-readable results gain a small, honest addition — the
selection, comparison and classification of every smoke run upload as
`container-performance-smoke-comparison-<run id>`, each stating in its own text
whether it detected regressions or merely checked plumbing. Two audit
corrections are owed to earlier entries: change 0587's evidence-gap table
described the smoke self-comparison as "a real, working plumbing check", which
was true when written but had not been true since `126c4a8b2`; and change 0421's
"policies opt into the new identity once both sides are freshly captured"
applies to the CI allocator policy, where both sides always are, and was never
acted on. No production code, contract, limit or refusal is touched. OLE2/OOXML
remain active; ODF is deferred until completion and iWork excluded.

## 0628 — ADR 0006's determinism clause, checked across OPC and found to hold everywhere but one place

Record: [0628](0628-opc-relationship-iteration-order.md).

`docs/GOAL.md` puts correctness and lossless preservation above speed, and ADR 0006 states the rule this change checks: *"Serialization is deterministic unless a `Clock`, actor identity, or cryptographic RNG is explicitly supplied."* Change 0600 recorded a suspicion — `Relationships::iter()` walks a `HashMap` — and this change turns it into a result. The audit is the deliverable: five separate paths through which a relationship order could have reached a published byte, a typed refusal, a catalog verdict or a read order were each traced to a sort, a lookup, an aggregation or a uniqueness requirement, and each is cited by file and line in the record so the next reader does not have to redo it. The one escape is relationship *reuse*, and its fix is chosen so that nothing else moves: `reuse_candidate` matches on exactly the old predicate, so it cannot fail where `find` succeeded or succeed where `find` failed; `min()` is infallible and allocation-free, so no refusal changes identity or position and no bound loosens; a collection with zero or one match behaves exactly as before, which three of the seven new tests exist to pin. Three alternatives are recorded with why they lost: a numeric rId order (needs a fallible parse, and OPC does not require the `rId<digits>` spelling, so it would put a second and differently-behaved ordering in one module), a retained insertion order (makes the answer a function of the call history rather than of the collection, and costs either a second table to keep consistent across `remove`, `retarget` and `retain` or a linear duplicate check), and refusing outright when duplicates exist (all four corpus packages are files real producers emitted, and a duplicate group is legal OPC, so this would trade preservation for tidiness). Audit rows this change opens rather than closes: **86 order-sensitive observable sites** remain in `litchi-xlsx`, `litchi-docx`, `litchi-pptx` and `litchi-ooxml-common` out of 497 production sites audited, and three that change a *verdict or a value* rather than a message were verified directly — the XLSX value-only snapshot's styles-and-theme array, which is compared with slice `==` and can refuse a patch as stale on the ordinary two-relationship shape; the PPTX ActiveX descriptor and binary states, where an unsorted `Vec` is compared against a sorted one and folded into the public snapshot `Revision`; and the DOCX content-control signature staleness token, whose bytes are emitted in hash order inside an otherwise sorted structure. One row is opened inside `litchi-opc` itself and deliberately not acted on: `get_or_add` drops the open-time source capture even when it reuses an existing relationship and changes nothing, so a `.rels` member change 0593 would have copied verbatim is reserialized instead; making it return early would move published bytes wherever the source spelling differs from the canonical one, which is a contract change and stops at a report. No latency, RSS, allocation or throughput improvement is claimed by this batch; the change makes a reuse very slightly more expensive. OLE2/OOXML remain active; ODF is deferred until completion and iWork excluded.

## 0613 — the original-bytes publication audit, priced and its memo declined

Item XLSX-3 of the 0587 queue is answered for its memo half and closed without a
production change. The audit of the original bytes of a replaced Part is
measured at 27.59% of source-backed XLSX publication instructions (0587 modelled
"at most about 28%"), so the opportunity is real, but the mechanism proposed for
it is not reachable: publication consumes its package, so a memo held inside
`SourceBackedPackage` has no second reader, and measurement confirms zero hits.
Reaching it requires either a non-consuming publication door or a caller-owned
cross-lineage memo with a digest key — both contract changes needing their own
records — and both sit behind change 0602's D0, which this change does not
touch. The P1 row "Finish source-backed CRUD adoption across formats" is
unaffected; the P0 publication-intersection rows gain one priced owner. 737
`litchi-opc`/`xml-minifier` tests and 3,629 `litchi-xlsx`/`litchi-docx`/
`litchi-pptx` tests pass on the retained candidate patch; the committed tree
changes no file under `crates/`. `performance_claim: none`. OLE2/OOXML
optimization remains active; ODF is deferred until completion and iWork
excluded. [Change and limitations](0613-opc-original-audit-memo.md);
[evidence](results/change-0613/README.md).

## 0624 — the audit's parallelism row now has a size, and its threshold has a correction

Design only; no production file was modified, so every limit, refusal, audit and
fence is exactly as change 0618 left it. Against `docs/GOAL.md` Workstream F's
"compression of changed output members" row, which change 0615 recorded as
having no parallel path at all, this record supplies the size the row lacked and
two corrections the audit should carry. First, the size: deflate of the changed
set is 24.19% to 61.94% of a save's native user cycles across five scenarios
spanning real fixtures, the harness corpora and the authored path, which makes
it the largest single term in every save measured. Second, the correction to the
gate: the "two or more members of 256 KiB or more" rule the survey inherited
from changes 0009, 0498 and 0499 does not transfer, because those records
measured *decompression batches* contending on one positional source with
per-wave thread creation and per-task memory reservations, whereas a deflate
task borrows an owned slice, shares nothing and returns a `Vec` — measured
profit begins at four 871-byte members, three orders of magnitude under the
proposed floor, and the real discriminator is whether one member dominates the
set. Third, the reason the row stays open: `docs/GOAL.md` rule 9 requires CPU
parallelism to be "opt-in and controlled by an explicit execution context with
thread, memory, I/O, cancellation, and task-granularity budgets", and a grep at
the base commit finds zero `ExecutionContext` or `ExecutionLimits` references
under `litchi-opc/src/pkgwriter.rs`, `atomic.rs`, the `soapberry-zip` writers or
the `litchi-cfb` writers; proposed ADR 0031, which would supply them, is not
accepted, and change 0615 §7 names its acceptance and gates G1–G3 as the
prerequisite for exactly this work. The audit rows this does **not** close:
nothing parallel was built and nothing end-to-end was timed, so every
end-to-end figure is Amdahl arithmetic and is labelled modelled; the DOCX
multi-part case the survey names was measured through the `OpcPackage`
publication route over all 62 DOCX fixtures rather than through `litchi-docx`'s
own editor, and the largest DOCX fixture's twelve ≥256 KiB members are embedded
fonts no paragraph edit regenerates; the authored path, which carries the
largest modelled win, sits behind the harder writer boundary and is out of scope
for a first implementation; and compression level remains fixed at 6 everywhere
and is still unmeasured, as SAVE-6 noted. Cold cache, peak RSS, allocation
profile, physical-device and cross-platform behaviour were not measured, and
allocation matters here because change 0618 rejected a pooled compressor on
glibc page-fault behaviour — this record makes re-measuring that an admission
gate (A5) rather than assuming per-worker state behaves differently.
`performance_claim: none`. OLE2 and OOXML remain active; ODF is deferred until
that goal completes and iWork is excluded.
[Change](0624-parallel-changed-member-deflate-design.md);
[evidence](results/change-0624/README.md).

## 0622: sixteen bytes per cell carried from planning delete the XLSX commit's second whole-sheet scan

Record: [0622](0622-xlsx-compact-source-facts.md).

Change 0622 closes item XLSX-1 of the 0587 survey and, with it, the line
0550 opened ("the next task is a private source-bound layout proof that could
avoid the second complete source scan while preserving lexical spans, scanner
facts, error order, resource limits and output validation"). It satisfies 0550's
admission condition — a candidate must improve planning *plus* commit and
publication rather than shift work — with a measured 33.6-41.4% reduction in
planning-plus-commit instructions and no change to publication. It also clears
the two gates that rejected changes 0552 and 0553: valid no-op planning is
unaffected (the builder never runs when the plan is empty, and an empty plan
returns before the fact route is consulted), and process peak RSS rises 4.23% at
the median on the shape where 0552 measured +6.97% and 0553 +5.16%, inside that
leg's own 6.49% run-to-run spread. The optimization order in `docs/GOAL.md` is
respected: this removes unnecessary work and unnecessary parsing, ahead of
layout, algorithms and parallelism; nothing was vectorized and no parallelism was
introduced. **The audit gap this change does not close**: the route it optimizes
is unreachable on real producer files. Of 391 real worksheet parts in
`test-data/`, the value-only planning validator accepts exactly one, and that one
publishes facts; change 0602 already recorded that the editor admits none of the
95 real `.xlsx` fixtures. The measured win is on the harness corpora.

## 0621 — a full XLS text projection took four source observations per shared string; it now takes one per read

Record: [0621](0621-xls-open-fence-count.md).

GOAL step 2 (eliminate unnecessary I/O) applied to the one kind of I/O the program had never priced on the text path: the freshness observation itself, which on a `from_path` workbook is a `statx`. The measurement-blocker-is-work rule was applied rather than noted. Change 0587 recorded that no record attributes the residual observations and that the XLS survey could only model them "from code, **unmeasured**"; this change wrote a probe that captures and buckets a backtrace at every `ReadAt::version()` call, so every one of the 89,789 observations on a `54016.xls` text projection is attributed to a file and a line on both legs, and the probe's source is retained. The decision rules are satisfied in order: before measurements from the shared read-only checkout of the base, hypothesis and mechanism stated per site, smallest coherent change, correctness and adversarial evidence, after measurements with identical setup. The **evidence-tier** rule decided how the two survey items are reported: CORE-2's size was **modelled** at 1-1.5% of an open and measures 25 → 22 observations with a p50 move inside the A/A floor, so it is reported as a count with no latency claim and its falsification condition is stated as met; XLS-4's size was **modelled** at ≥48,165 `fstat` and measures 53,283 of 89,794, so its falsification condition is stated as not met. **Reporting the item that failed is the point of the record**, not an aside: "collapse the 25 per-open fences to one per operation boundary" cannot be done, because 18 of the 22 that remain are structural and 16 of those are the per-read fence ADR 0005 needs and change 0558 already halved, and saying so is what stops the item being re-ranked in a later survey. Two audit rows this change opens. The first is that no harness selector attributes source observations per site, so the probe had to be a scratch Cargo project with path dependencies; a `--mode` on `xls_source_attribution` that reports observations per call site would make this class of analysis repeatable. The second is that the `facade-file` control's `statx` counts are whole-child and cannot be decomposed, because the release binary carries no symbols for `strace -k`; a debug-info release profile for the attribution binary would close it. No cold-cache, physical-device, range-source, RSS, allocation or concurrency result is claimed, and the count claim's relevance on a network filesystem — where a `statx` is a round trip rather than 167 ns — is modelled, not measured, which is change 0587's own blocker B6 restated. OLE2/OOXML remain active; ODF is deferred until completion and iWork excluded.

## 0625 — `docs/GOAL.md`'s determinism requirement, restored on the OLE2 write path

Record: [0625](0625-cfb-writer-deterministic-storage-order.md).

`docs/GOAL.md` puts correctness, lossless preservation and bounded resources above speed, and ADR 0006 states the rule this change restores: *"Serialization is deterministic unless a `Clock`, actor identity, or cryptographic RNG is explicitly supplied."* `OleWriter` is supplied none of the three and was non-deterministic anyway, through hash-table iteration order rather than through any declared source of entropy — which is why no audit row had caught it and why change 0617 only found it by digesting the same build twice. Every OLE2 writer in the workspace reaches this code: `Package::render` in `litchi-ole-common`, behind every length-changing XLS, DOC and PPT save; the custom-XML codec; `litchi-crypto`'s encrypted-OOXML container writer and its `rebuild_ole`; `litchi-vba`'s project writer; `litchi-sign`'s CFB signature writer; and `litchi-ograph`'s package writer. The fix is confined to `write_to` and is chosen so that no other contract moves: the canonical order is infallible, so no refusal changes identity or position; it is total over the distinct keys of a `HashSet`, so it needs no tie-break; it places every ancestor ahead of its descendants, which is what `add_storage_path` wants; and it keeps the two hash tables, so `create_storage` keeps the `try_reserve` allocation-failure path that `reserve_hash_set_entry` exists to provide. Three alternatives were considered and are recorded with why they lost: the MS-CFB sibling key via `canonical_cfb_path` (fallible — it would move a refusal), `BTreeSet`/`BTreeMap` fields (no `try_reserve` — it would abort instead of returning `OleError::Allocation`), and an insertion-ordered collection matching `SequentialOleWriter` (linear duplicate check, so *n* storages become O(*n*²) under adversarial control, and the output would be a function of the call history rather than of the document). Two audit rows this change opens: the corpus differential is writer-level rather than save-level, so a digest-comparing differential through each format crate's own editor would be a strict upgrade; and `SequentialOleWriter`, which is already deterministic, still uses declaration order, so the two writers agree only for callers that sort. No latency, RSS, allocation or throughput improvement is claimed by this batch; the change makes a save very slightly more expensive. OLE2/OOXML remain active; ODF is deferred until completion and iWork excluded.

## 0620 — the XLS edit-and-save path attributed; one duplicate source parse removed; the readback-owner swap frozen as inadmissible

GOAL step 1 (eliminate unnecessary work) in the narrowest admissible form: the
work removed is one complete parse of bytes the same call chain had already
parsed, under the same limits and the same compatibility profile, whose three
verdicts were already recorded. No step was skipped and no later step was
reached — no layout change, no algorithm substitution, no parallelism, no SIMD —
and no validation was moved, relaxed or deferred, because the candidate readback
this record is about is untouched: the target is still materialized, reopened
through a complete `Workbook::new` and validated again by a second one. The
decision rules are satisfied in order: before measurements from the shared
read-only checkout of the base, hypothesis and mechanism stated, smallest
coherent change, correctness and adversarial evidence, after measurements with
identical setup. The **measurement-blocker-is-work** rule was applied rather than
noted: 0587 recorded that no attribution of this path existed and that the
harness's XLS attribution binary supports only `open`, `list` and `one-cell`, so
this change wrote a probe that drives the real editor on real fixtures and
retains its source. The **instructions-rank-work, cycles-price-latency** rule
decided how the result is stated: callgrind reports −14.0% and native `perf stat`
−25.1% on the same operation, and the gap is exactly the software SHA-256
valgrind is forced into, which inflates the fingerprint terms the change does not
touch. Evidence tiers: **measured** for every per-call-site Ir figure (32
isolation pairs, 64 annotated profiles), the 32 native counter pairs, the 48
counter rows, the 5,760 timed operations across four rounds and the 3,546
corpus-differential rows; **modelled** for nothing — this record makes no
arithmetic prediction; **unknown** for the cost of the two remaining target
parses beyond their measured Ir, and for the shares on any workbook with a record
mix unlike the three measured. Host quiescence is established for the retained
timing run and is *not* assumed: the first timing run of this change was taken
while a corpus sweep occupied two other cores, its A/A floor reached 30% at p50,
and it was discarded and re-run rather than reported. The honest control band is
stated as ±5% at p50 — larger than the measured ±2% A/A floor — because the two
binaries differ in code layout and the unchanged `open` phase shows it. OLE2/OOXML
optimization remains active; ODF is deferred until completion and iWork excluded.
[Change and limitations](0620-xls-edit-save-attribution.md); [retained
evidence](results/change-0620/README.md).

## 0600 — one fewer observation per cold Part read, a bounded monitored-read scope, and an allocation-free member lookup

`docs/GOAL.md` puts unnecessary work ahead of I/O and unnecessary allocation
ahead of layout and algorithms, and this change takes only those two steps: not
one positional request, not one requested byte and not one output byte moves on
any measured phase, on any of three real fixtures, in either read order. The
"exact no-ops stay exact" rule is what bounds the design — the observation
removed is the only one of five that a strict-subset argument reaches, and the
relocation onto the failed-publication branch keeps change 0317's precedence
rather than trading it for a count. The `monitor_reads` bound was allowed only
after checking that change 0327's publication protocol never reads the flag: its
fences are direct `ensure_current()` calls, so bounding the flag cannot weaken
them, and what the flag does guarantee — a per-chunk fence *during* an
incremental copy — is now scoped to exactly the operation that needs it. The
audit's standing caveat that instructions rank work rather than latency applies
to the 3.45% figure and is stated in the record; `performance_claim: none` and
no claim-registry entry. The pre-existing non-determinism found on the way — the
relationship iteration order of `Relationships::iter()` varies per process
because it walks a `HashMap` — is reported, not fixed.
[Change and limitations](0600-opc-cold-read-observations-and-name-lookup.md);
[evidence](results/change-0600/README.md).

## 0612 — a `docs/GOAL.md` step-1 candidate closed by measuring the cost it trades into, not the work it removes

Record: [0612](0612-xls-skip-uninterpreted-globals-design.md).

`docs/GOAL.md`'s optimization order puts "eliminate unnecessary work" first and "unnecessary I/O" second, and 75.30% of the globals bytes a source-backed XLS open reads are payloads no consumer interprets — by byte volume the largest remaining instance of both on this path. This record establishes that removing them is not free, because the same order requires "bounded read-ahead and range coalescing based on measured access patterns" and avoiding "request amplification from many tiny remote reads", and the skip trades directly against both: it converts one large read into more, smaller reads, and each read on a `from_path` workbook is a `pread64` **and** an `fstat` freshness observation, 29 rising to 85 on one flagship open. Measured, the trade loses on every source kind this repository can build — +5.54% of the open in cycles on an owned in-memory source, +17.22% on a file source — and the only source kind on which it could win, a range or remote source, does not exist in the corpus, so its side of the trade is modelled from change 0564's 116 ns per request and **never measured**. That gap is the audit row this record leaves open, and it is the same row change 0587 opened for the whole read path. Two further audit rows are narrowed rather than closed: `xls_source_attribution` still has no full-text, no all-cells and no `from_path`-observation selector, so the scenarios this design would make *worse* cannot be timed (both re-enter the same open, so both would be worse, not better); and the design's own admission gates — a `ParsedGlobals` differential over all 123 modellable fixtures, error identity on the three encrypted refusals, and paired timing on dense and sparse globals in both directions — are frozen in the record so a later batch inherits them rather than re-deriving them. The decision rules were applied as written: the schedule was modelled before anything was built, three measurement scaffolds were built, measured and reverted rather than argued about, the design was priced in its **best** form (skip plus retained cursor, which is worse still at +18.68%/+15.99%), and the result that contradicted the ranking was reported rather than folded into a mean. `docs/GOAL.md` rule 10 is untouched (no `unsafe`; `#![forbid(unsafe_code)]` stands in `litchi-xls`) and rule 12 is untouched (no limit, budget, cancellation point or malformed-input defence is weakened, because nothing changed). No latency, RSS, allocation or throughput improvement is claimed by this batch. OLE2/OOXML remain active; ODF is deferred until completion and iWork excluded.

## 0618 — the save path stops rebuilding its compressor for every member

Retained implementation against the audit's standing "eliminate unnecessary
work before unnecessary I/O" ordering, on the authored-publication clause. Every
authored OOXML save allocated, zeroed and released one ~300 KiB Deflate state
per output member; a 39-member authored save faulted in and released 39 of them,
at a net 95 minor page faults per save, and now constructs one and faults in
none. Correctness is
established by whole-corpus byte identity rather than by inspection: 336 OOXML
fixtures × 4 mutation scenarios (1,344 rows) and the same fixtures × 3
multi-member regeneration scenarios (1,008 rows) published identically on both
legs, including 8 preserved open refusals and 27 preserved save refusals matched
by typed-error text; 422 rows from the real `tabs`, `edit_cells` and
`append_plain_paragraph` routes over the whole corpus, 110 published digests and
312 typed refusals, all identical; 48 authored-PPTX digests across six deck
sizes; and 78 opened decks still saving as exact-source passthrough on both
legs. The audit rows this does **not** close: authored DOCX and XLSX publication
uses the same writer and the same code path but was covered only by the oracles
and the test suites, not by its own timing, because `tools/perf-baseline` still
has no ordinary-save selector; and the preservation numbers are the fixed
per-member cost on `.rels` members of a few hundred bytes, since no fixture in
the corpus regenerates a large content part. Cold cache, peak RSS, real-device
behaviour and non-glibc allocators remain unmeasured — and the last of those
matters here, because the rejected pooled variant was rejected on glibc page-fault
behaviour. `performance_claim: none`. OLE2 and OOXML remain active; ODF is
deferred until that goal completes and iWork is excluded.
[Change](0618-zip-writer-deflate-state-reuse.md);
[evidence](results/change-0618/README.md).

## 0619 — the instrument was wrong, not the library; one standing failure closed

Record: [0619](0619-harness-xls-lifecycle-assertion.md).

This is the audit's own evidence hygiene turned on the program's instrument. `docs/GOAL.md` puts correctness above speed and requires every claim to be scoped; change 0601 recorded a standing harness failure as pre-existing, reproduced it on an untouched checkout and left it open with a pointer at change 0595. That pointer was wrong, and this record establishes it by measurement rather than by reading: eight bisect legs place the first bad commit at `c1d2caf85` (change 0565), whose parent passes, so the failure predates 0595 by four production changes and 0605 by ten. The audit's standing requirement that a regression be attributed before it is acted on is therefore met without touching either record: **no correction is owed to 0595 or 0605**. The correction is owed to **0565**, which relaxed one gate function and — correctly, in the sense that the record measured the exact 93 bytes and named the corpus — did not notice that the same contract was asserted three more times in the harness's unit test and encoded in two published booleans. The cause is structural and worth the audit's attention: `tools/perf-baseline` is a separate Cargo project, so a workspace `cargo test` never reaches it, and a change that edits harness code can pass every gate in the standing list while leaving the harness's own suite red. That is the same class of gap change 0587's evidence table calls "the measured path is not the path real files take", one layer down — the *gating* path is not the path the harness's own tests take. Evidence tiers: **measured** for all eight legs, the fifteen-read trace, the five range-set byte totals and the after-state; **modelled**: nothing; **unknown**: whether any fixture outside this corpus and 0565's 104-fixture survey drives the over-read above 3,949 bytes. What this batch does not discharge: the two published booleans keep the strict pre-0565 meaning and now publish `false` honestly rather than stating the bounded invariant, which is a schema change left open along with its `verify-capture.py` migration; the four pre-existing Python gate failures 0601 reported are untouched; and no production fix for the over-read exists or is proposed, because the globals end is not knowable until a `BoundSheet8` has been framed, which is why 0565 bounded the contract instead of preserving it. `performance_claim: none`; no file under `crates/` changed. OLE2/OOXML remain the active priority; ODF stays deferred; iWork is excluded.

## 0615 — the execution context bounds one session at a time, and the workspace runs three

`docs/GOAL.md` rule 9 requires CPU parallelism to be opt-in and controlled by an explicit execution context with thread, memory, I/O, cancellation and task-granularity budgets, and Workstream F names seven capabilities such a context must control. This record audits `litchi_core::ExecutionLimits` against that list and finds four and a half present: maximum worker count, the memory and decompression-in-flight budget, cancellation, and half of the task-size threshold — `min_parallel_bytes` is an aggregate with no per-task floor. Absent: I/O concurrency, a CPU task budget distinct from byte-denominated `Work`, and an optional caller-provided executor or scoped worker facility. Two further Workstream F requirements have no implementation at all: bounded overlapping pipelines (every session is wave-synchronous, and change 0499's contract is explicit that later waves cannot begin until the current one completes), and the dependency DAG itself — independent worksheets, slides, stories, sections, disjoint edit subplans, independent validation domains and compression of changed output members have no parallel path, and the three that do exist have no format-crate or facade caller. **No audit row closes.** The measured coexistence counts (6, 12 and 20 session-owned worker threads at `workers` 2, 4 and 8, from one budget root and one policy) are the evidence that the missing budgets must be hierarchical `Resource` dimensions rather than more per-session policy fields; the design is drafted as proposed ADR 0031, which amends the execution paragraph of accepted ADR 0005 and is therefore left for human review rather than implemented, exactly as `docs/GOAL.md` line 80 requires. No code changed, so lossless preservation, typed refusals and every limit are untouched. OLE2/OOXML optimization remains active; ODF is deferred until completion and iWork excluded. [Change and limitations](0615-execution-context-completeness-design.md); [retained evidence](results/change-0615/README.md).

## 0617 — `docs/GOAL.md`'s "unchanged-stream or unchanged-sector copy-through on save" designed, priced and blocked on an ADR

Record: [0617](0617-cfb-copy-through-writer-design.md).

The LEGACY CFB-SPECIFIC WORK list asks for "unchanged-stream or unchanged-sector copy-through on save". This change writes that design out in full — a fourth `litchi-cfb` publication path alongside `OleWriter`, `SequentialOleWriter` and the same-length overlay, built by generalizing `ComposedOverlaySource` from spans-over-a-source to spans-plus-an-appended-tail, with a planner that rebuilds only the FAT, DIFAT, MiniFAT, directory and changed streams and copies every other directory byte verbatim — and states its nine invariants, its seven admission gates and its seven-part parity gate. It is **not** implemented, for two independent reasons. First, measurement: on XLS the container rebuild is 1.07% of a save in cycles, inside this host's floor, so the design would buy nothing there; on DOC it is 13.25-29.99%, so it would. Second, and blocking regardless: **no ADR addresses the physical sector layout of a saved CFB.** ADR 0005 authorizes the mechanism ("Preserve-mode save raw-copies unchanged compressed ZIP entries or CFB streams when possible") and ADR 0008 names sector *size* in its preservation list, but nothing says whether a saved file may reuse a `FREESECT`, must append, or must preserve placement. All three policies preserve every logical byte; append-only grows `FloatingPictures.doc` by about 70 KiB on every save, and free-sector reuse relocates bytes a byte-comparing differ can see. Choosing one in code would be making preservation policy inside a performance change, so the record asks for the clarification and recommends reusing only the sectors the operation itself released. Two audit rows this change opens: the unchanged-stream share, which decides whether CFB-2 is worth a rank at all, is measured on five fixtures and ranges 2.3-81.4%, and a census over all 212 OLE2 fixtures is the cheapest next measurement; and no `.ppt` fixture in this repository can reach a length-changing shape-text edit at all, so that landed capability has no corpus. No latency, RSS, allocation or throughput improvement is claimed by this batch. OLE2/OOXML remain active; ODF is deferred until completion and iWork excluded.

## 0609: the facade's `.doc` slurp is the cheaper route — routing it to the source-backed DOC reader would cost 2-5× the cycles and admit four artifacts the eager reader refuses

Record: [0609](0609-facade-doc-source-route-design.md).

**A source-backed route is not automatically the cheaper route, and this one is
6.3× to 9.9× worse at p50.** `docs/GOAL.md` hypothesis 1 ranks whole-input
ingestion as the first thing to remove, and the facade's `.doc` slurp is the
last such ingress on a priority format. Change 0609 measured it against the
alternative the repository already owns and found the slurp cheaper on every
axis but heap peak: 2 `pread64` against 30, 11 `statx` against 153, one complete
file read against six, 2.12× to 5.07× fewer native cycles. The reason is a
structural one this audit should carry forward — the eager reader's cost is
proportional to *document content* and the source-backed reader's to *artifact
bytes*, because ADR 0006's complete-artifact identity fence is levied on the
file's length before any paragraph is resolved. Removing whole-input ingestion
is only a saving when what replaces it reads less, and a fence that hashes the
artifact three times (six reads) reads more. The second finding is an admission
one: the two readers' refusal sets are **not nested**, so "try the cheap reader,
fall back to the strict one" is not a safe pattern in this repository whenever
the cheap reader skips a validation the strict one performs — here it would have
admitted four `.doc` fixtures the library calls corrupt. The third is a cost of
the fallback shape itself: a source-backed probe that is *expected* to be
refused still costs two complete artifact passes, measured at +25.8%, +192.7%
and +245.4% on top of the eager open for three real fixtures, growing with file
size. Evidence gap 5 of change 0587 is unchanged by this batch: there is still
no `perf-baseline` selector that opens a `.doc` or `.ppt` through the facade, so
every figure here comes from a retained scratch probe rather than the harness.

## 0610 — hypothesis 8 is confirmed, and the waste is measured

Design only, against the audit's standing eager-OPC-materialization row.
`docs/GOAL.md:350`'s hypothesis 8 — *"XLSX selective sheet loading may be
undermined by eager OPC materialization"* — is now **confirmed by measurement
rather than by source reading**: the XLSX layer above OPC already defers
worksheet work, and an open-then-hide-a-tab-then-save reads exactly one part,
`/xl/workbook.xml`, on every one of the 33 corpus fixtures that admit the
operation, while `load_parts_eager` has already inflated all 550 parts and
8,274,037 bytes before the workbook model exists. 99.27% of those bytes are
decompressed, charged, retained and never read. The record also closes a gap
0581 left explicitly open: 0581 inferred that its measured open retention is
`archive + Σ decompressed payloads` and said the inference was not a
measurement; the new census measures Σ and the formula accounts for 98.32% of
0581's figure. `GOAL.md`'s standing instruction — *"Do not change an
architecture solely because it appears suboptimal in source code. Require
profiles and scenario measurements"* — is honoured: the measurements exist and
the architecture is still not changed, because gate 1 requires a human to accept
[ADR 0030](../adr/0030-lazy-opc-part-decode.md) first. Audit rows this does
**not** close: no DOCX or PPTX semantic-editor read set is measured, because no
example opens a real `.docx` or `.pptx`, edits through the model and saves —
0587 §4 recorded that gap, 0593 recorded it again, and it remains open; a DOCX
save regenerates every model-owned part when one whole-model flag is set and a
PPTX save regenerates every slide, so those read sets are **unknown, not small**.
`tools/perf-baseline` still has no ordinary open-and-save selector. One new
audit item is opened: the pre-existing `NotCompact` publication refusal blocks a
plain tab-hide on 59 of 180 real `.xlsx` fixtures, which is a correctness defect
sized here and owned by no record. `performance_claim: none`, and no timing,
peak-RSS or cold-cache figure is measured. OLE2 and OOXML remain active; ODF is
deferred until that goal completes and iWork is excluded.
[Change](0610-opc-lazy-part-decode-design.md);
[evidence](results/change-0610/README.md).

## 0608 — a `docs/GOAL.md` step-1 candidate closed on reach rather than on size

Record: [0608](0608-xls-lazy-sst-index-design.md).

`docs/GOAL.md`'s first optimization step is "eliminate unnecessary work", and the eager SST walk looked like the largest remaining instance of it on the XLS open path. This record establishes that the work is not unnecessary for the scenarios the corpus is made of. The walk's size is confirmed and sharpened — 3.19% to 35.73% of an open in instructions, 2.99% to 30.99% in native cycles, measured on three fixtures with two A/A controls in the same window — but the corpus census shows a prefix index must walk 100% of the table for a full text on every fixture that has one, a median 63.88% for one uniformly chosen string cell, and 0% only for open-and-list. The audit rows this leaves open are the ones the record could not close by measurement, and both are standing XLS blockers rather than new ones: `xls_source_attribution` still has no full-text, no all-cells and no `from_path`-observation selector (change 0587's blocker), so the scenarios this design would make *worse* cannot be timed at all and its re-read cost is modelled from byte counts; and no fixture in the repository is refused by the per-string SST walk — all four SST refusals in the corpus are header refusals — so the synthetic malformed-SST differential the design's admission gates require cannot be derived from this corpus and does not exist. The decision rules were applied as written: the ceiling was captured before any design was believed, the one form that preserves every refusal was built and measured rather than argued about, and it was reverted because a 0.44–3.13% cycle saving on open-and-list bought by doubling the walk on every other path is not practically useful. `docs/GOAL.md` rule 10 is untouched (no `unsafe`, no `MaybeUninit`; `#![forbid(unsafe_code)]` stands in `litchi-xls`) and rule 12 is untouched (no limit, budget, cancellation point or malformed-input defence is weakened, because nothing changed). No latency, RSS, allocation or throughput improvement is claimed by this batch. OLE2/OOXML remain active; ODF is deferred until completion and iWork excluded.

## 0606 — PPT record tree borrowed payloads

Step 2 of the optimization order (unnecessary copying and allocation) applied to the last unaddressed PPT parser, which change 0301 had explicitly left out of scope. The eager and the source-backed `Presentation` now parse a record tree whose payloads span the `PowerPoint Document` stream the presentation already retains, and the eager per-open full-text extraction that only a unit test consumed is deferred to first use. Retained bytes per byte of stream for an eager open fall from 2.87x to 1.24x on `45543.ppt`, 3.30x to 1.31x on `SampleShow.ppt` and 3.24x to 1.25x on `headers_footers_2007.ppt`; allocations halve. No P1 audit row closes: this is one format's reader, the source-backed PPT path still pays the whole-artifact SHA-256 that 0587 ranked first, `SlideFactory` still copies per slide, and no RSS, cold-cache, physical-device or producer-corpus measurement was taken. Lossless preservation, typed refusals and every record limit are unchanged, and all 30 `.ppt` fixture reader dumps — including the four encrypted refusals — are identical across the two legs. OLE2/OOXML optimization remains active; ODF is deferred until completion and iWork excluded. [Change and limitations](0606-ppt-record-tree-borrowed-payloads.md); [retained evidence](results/change-0606/README.md).

## 0601 — reproducible baselines on the corpus class the audit was missing

Record: [0601](0601-perf-harness-real-producer-shape.md).

Retained against the audit's standing "corpus classes including multi-producer"
and DEFINITION OF DONE clause (a) rows. The audit has carried "most real
DOC/PPT and a chunk of real XLS/XLSX/DOCX/PPTX fixtures can't reach the measured
path" as a gap since the 0587 entry; this change closes the generator half of
it for OOXML. The new corpora are deterministic — two independent release-mode
processes produce byte-identical archives, and both the census diff and the
corpus-identity diff are empty — and they are shown to trip the same library
gates a real Excel file trips, part for part, through the same code path. The
`--real-file PATH` opt-in adds the other half for XLSX reads: the only input in
this harness whose bytes come from outside the process, bounded to 32 MiB,
absent from the default matrix, with the file's path, size and SHA-256 bound
into the corpus identity. `--producer-evidence PATH` makes the marker and
refusal census a first-class harness output, which is the audit's evidence-gap
3 ("a tracked refusal census instead of narrative repetition per record").
**What this batch does not discharge**: the selector gap is untouched for DOCX
and PPTX edit/save, the eager XLSX paths, file-backed and range sources and
XLSB; the dense shape is 128 × 128 rather than 256 × 256 because the codec makes
the larger size seconds per sample; the CI smoke job still compares against
itself; and no instruction-level attribution was taken, so the numbers rank
latency on one host, not work. One standing failure was confirmed on the way and
not fixed: `xls_source_backed_lifecycle_selectors_are_matched_and_local` asserts
`open_reads_zero_worksheet_payload == [true]` and gets `[false]`, identically on
the untouched checkout at `f8cf7d2a1`, and its panic poisons the shared
allocation-metrics mutex so six later tests fail as cascades.

## 0603 — the value-only worksheet vocabulary, not the marker gates, is what excludes real producers

Record: [0603](0603-xlsx-fused-traversal-marker-admission.md).

Change 0602 measured that the source-backed value editor admits 0 of 95 real
`.xlsx` packages at its package-root relationship gate. This record measures the
layer beneath: of the 207 worksheet parts in the same corpus, change 0546's
fused traversal completes on **none**, before this change or after it, because
each is refused by the sixteen-name value-only element allow-list or its
attribute allow-list, or exceeds the 131,072-event provisional cap, long before
the compatibility-marker gates decide anything. Admission itself widens from 71
to 102 of 207 parts (71 borrowed, 31 newly proved), and the 105 that stay
refused are refused for one reason only: `dyDescent`, whose values the fallback
captures in a separate x14ac pass and the shared reader does not. The audit's
standing P1 row "cover the high-impact CRUD categories and real producers"
therefore gains a second measured obstruction on this path, and it is the same
vocabulary obstruction change 0602 recorded: the marker gates were never the
binding constraint. The harness gap change 0587 reported is unchanged — no
generated worksheet carries `mc`, `x14ac`, `dyDescent` or `cols` markup — so the
shapes measured here are change 0602's projections, regenerated by that record's
own scripts. `performance_claim: none`; `claim_authorized: false`. OLE2/OOXML
remain active; ODF is deferred until completion and iWork excluded.

## 0607 — PPTX opened-document CRUD has one route, and the example that takes the other is broken

Record: [0607](0607-pptx-authored-slide-regeneration-design.md).

The audit's standing P1 row "finish source-backed CRUD adoption across formats"
gains a measured boundary on PPTX. There are two writer surfaces in
`litchi-pptx`: the mutable `MutablePresentation`, reachable only from
`Package::new()`, and the `opened::Transaction` capture-and-patch route,
reachable only from an opened package. They are mutually exclusive by
construction, and the census confirms it on the whole corpus: 78 of 78 `.pptx`
fixtures open and 78 of 78 refuse `presentation_mut`, while an unedited save of
each is byte-identical to its source. Opened-document PPTX CRUD is therefore
entirely `opened/`'s (changes 0590 and 0598), and the mutable writer is an
authoring surface only. One in-tree consumer has not noticed:
`crates/litchi/examples/office_crud_demo.rs:289-290` performs its PPTX UPDATE
step as `Package::open` followed by `presentation_mut`, which cannot succeed, so
the CRUD demonstration fails at that step. Reported, not fixed: the fix is to
port the step to `opened_presentation_transaction`, which belongs to the
`opened/` owner. Nothing in the repository materializes an authored presentation
more than once, so the authored-save opportunity this record freezes has no
caller and no harness selector to be measured against, and that is its admission
gate. `performance_claim: none`; `claim_authorized: false`. OLE2/OOXML remain
active; ODF is deferred until completion and iWork excluded.

## 0605 — retained whole-sheet XLS walk; retained sheet index frozen as a design

GOAL step 1 (eliminate unnecessary work) in its purest form: the work removed is
N−1 complete validated scans of a worksheet substream, and nothing about the
remaining scan changed. No step was skipped — no layout change, no algorithm
substitution, no parallelism, no SIMD — and no validation was moved, relaxed or
deferred, because the walk runs to the worksheet's EOF exactly as a selected-cell
query does; an early exit would be change 0574's opportunity 6, which ADR 0005's
mandatory-validation clause rejects. The decision rules are satisfied in order:
before measurements captured from the shared read-only checkout of the base,
hypothesis and mechanism stated, smallest coherent change, correctness and
adversarial evidence, after measurements with identical setup. Evidence tiers:
**measured** for the 42 logical-counter cells, the 24 walk-against-per-cell
cells, the 18 callgrind isolation pairs (36 annotated profiles), the 18
`perf stat` pairs, the 120 timing runs across four rounds and the corpus
differential over every XLS fixture; **modelled** for the 24.6 s whole-sheet
per-cell extrapolation, which rests on three measured points that are linear to
within 5%, and for every number in the part (2) design. The rule that a measurement blocker is itself work
was applied rather than noted: change 0587 recorded that no source-backed XLS
all-cells or full-text selector existed, so this change built them, and their
all-cells oracle is a differential that fails the run on any disagreement between
the walk and a selected-cell query. The standing instruction to price
pointer-chase work in cycles is honoured, and it is what caught the one real
problem: the first shared loop cost +1.37% instructions and +1.36% cycles on the
`54016` one-cell query, `#[inline]` hints and a restored branch shape did not
recover it, and folding the sink's four constant arguments into one `ScanContext`
did, to +0.011%; the three trial profiles are retained rather than discarded.
Host quiescence is not established — eight measurement agents shared 32 cores,
1-minute load 8.27 — so the floor is measured in the same window: A/A p50 0.71%,
max 2.14% over the 18 cells that exist on both legs, but B/B up to 46.51% on the
heavy new scenarios, which is why no wall-clock statement is made about them
beyond their ratios. OLE2/OOXML optimization remains active; ODF is deferred
until completion and iWork excluded. [Change and
limitations](0605-xls-retained-sheet-index.md); [retained
evidence](results/change-0605/README.md).

## 0594 — one Deflate decoder per OOXML open, above a measured threshold

`docs/GOAL.md` puts unnecessary allocation ahead of layout, algorithms and
parallelism, and requires every claim to be scoped to scenario, corpus, machine,
build and metric. This change removes allocation without touching I/O, limits,
error identity or output bytes — positional requests, requested bytes,
`version()` calls, `syscr` and `rchar` are identical on every measured phase —
and it is scoped by a threshold rather than by assertion, because the measurement
found a regime where the trade loses. The bounded-resource cost is measured, not
argued: one retained 80,320-byte workspace, worth +15 minor page faults per
fresh-process open of a 132-member workbook and −2 on a 445-member PPTX, against
3.29 MB of allocation and zero-fill removed. The audit's standing instruction to
price pointer-chase work in cycles is what caught the first implementation:
instructions fell and cycles did not. `performance_claim: none` and no
claim-registry entry. The P1 row "finish source-backed CRUD adoption across
formats" is unchanged. [Change and limitations](0594-zip-session-reuse-per-open.md);
[evidence](results/change-0594/README.md).

## 0604 — a `docs/GOAL.md` rule-10 blocker resolved on paper, and closed on measurement

Record: [0604](0604-cfb-append-reads-design.md).

`docs/GOAL.md` hypothesis 11 — "CFB may … allocate and zero a sector `Vec` for frequent reads … and copy streams unnecessarily" — is confirmed as a mechanism and closed as an opportunity for this corpus. The safe shape that rule 10 demands (no new `unsafe`, no `MaybeUninit`) is now written out in full: a provided `ReadAt` method appending into a caller's `Vec`, whose default preserves today's zero-fill so no implementor breaks, with overrides for `OwnedSource`, `SliceSource` and `OwnedArcSource`; one `litchi-cfb` chain walk parameterized by a slice or appending sink so the chain-validation error order stays single-sourced; tail-only zeroing in the sector helpers; and `try_reserve_exact` up front so the typed `OleError::Allocation` refusal stays ahead of any read (rule 12). It is not implemented, because measurement puts its ceiling at 0.57–2.94% of an open in cycles, inside this host's 4% p50 floor, and because no stable positional API accepts uninitialized memory — so `FileSource` and every `File`-backed reader, which is what `Package::open`, `SourceBackedWorkbook::from_path` and `SourceBackedPackage::from_path` use, gain nothing at all. The audit row this leaves open is the corpus one: no OLE2 fixture above 1.6 MB, none with a DIFAT sector and none with 4,096-byte sectors exists, so a large-source claim for this or for CFB-3 and CFB-5 remains unmeasurable until a synthetic corpus is built. No latency, RSS, allocation or throughput improvement is claimed by this batch. OLE2/OOXML remain active; ODF is deferred until completion and iWork excluded.

## 0597 — XLSX selected-cell ineligibility gate, frozen at its design

Rank-6 survey item XML-2 is now priced (−63.19% of a source-backed one-cell read on the control fixture, −14.08% on the real one, measured as callgrind isolation pairs) and frozen: the gate moves which typed error a malformed `<cols>`-bearing worksheet reports, so GOAL's "capture BEFORE measurements, make the smallest coherent change, never trade a typed refusal for a partial result" line puts it behind a frozen design rather than in this batch. The same brief's oracle closed a correctness gap instead: `SourceWorksheet::cell`/`cells` refused 434 reads over 24 of 180 fixtures with `worksheet mergeCells appears before sheetData`, a question answered from state the scanner stops maintaining after `mark()`. After the fix the source-backed path agrees with the mandatory materialized parser on all 3,948 comparable rows (372 disagreements at the base). 1,297 tests pass. The unmeasured XML-1 codec expansion, the harness's lack of any marker-bearing or ineligible worksheet corpus, and the gate's two open prerequisites keep the non-iWork performance goal open. OLE2/OOXML remain active; ODF is deferred until completion and iWork excluded. [Change and limitations](0597-xlsx-selected-cell-ineligibility-gate.md); [retained evidence](results/change-0597/README.md).

## 0598 — PPTX cross-package copy revision reuse

Record: [0598](0598-pptx-cross-copy-revision-cache.md).

`docs/GOAL.md`'s first two optimization steps — eliminate unnecessary work, then
unnecessary I/O, serialization and hashing — applied to the cross-package slide
copy that 0587 ranked fifth and measured at 85% of a media-rich whole child.
Five of nine complete-archive serializations and four of twelve complete-graph
hashes per lifecycle are removed by reusing values the same call already proved
on the same bytes. ADR 0005's "cache behavior is semantically invisible" is the
licence for the one new piece of state: a private `OnceLock` on the immutable
`Snapshot`, keyed on the archive bound so no `Error::Limit` moves, never
inherited across `Snapshot::rebound_to` (where `packages_equal` proves graph
equality and says nothing about the retained archive), and always a plain
recomputation on a miss. ADR 0005's bounded-memory rule is also why
`bounded_package_bytes` keeps its `Vec` rather than becoming a hashing sink: the
`Vec` is the candidate archive that `from_vec_reusing_payloads` reopens, so the
saving is taken by digesting it in place. ADR 0003 and 0006 are untouched
because no revision value or proof format changes and the durable `LPCP0002`
encoding stays bit-identical; the four staleness proofs still run on the **live**
packages before any capture, so a stale or foreign source or destination is
refused exactly where it was. **Measured** −24.61% `Ir` per media-rich lifecycle
with exact before/after call counts (twelve semantic hashes and nine archive
serializations before; eight and four after), and −8.39% native whole-child
cycles against a −0.36% A/A floor. What stays open: the candidate is still built,
deflated and captured twice per lifecycle (5.96 G `Ir` of deflate on the copied
closure), and retaining the planned archive to remove the second build changes
`CrossSlideCopyPlan`'s memory profile under ADR 0005; `validate_candidate`'s
second capture is the proof that binds the candidate to the plan, so reusing it
is a contract question, not a value-identical reuse; the four remaining semantic
hashes need 0587's PPTX-1(c), which redefines the durable revision. No speedup,
RSS, allocation, cold-cache or real-producer claim follows, and the wall-clock
floors in this window (A/A −4.99% at p50 on the media-rich selector) are stated
beside every timing.

## 0602 — real-producer CRUD coverage is blocked before the editor, and before publication

Record: [0602](0602-xlsx-real-producer-admission-design.md).

The audit's standing P1 row "cover the high-impact CRUD categories and real
producers" now has a measured obstruction on the XLSX source-backed value path.
The source-backed value editor admits 0 of the 95 real `.xlsx` fixtures: 93
refuse at the package-root relationship allow-list with a typed
`value-only edits refuse package relationship` error naming `docProps/app.xml`
or `docProps/core.xml`, and 2 fail earlier in the read door. Opening that gate
exposes seven more, of which the vocabulary gates are the expensive ones: the
value-only worksheet element allow-list is sixteen names and admits no
`sheetPr`, `mergeCells`, `pageMargins`, `pageSetup`, `autoFilter`,
`conditionalFormatting`, `dataValidations`, `hyperlinks` or `extLst`. Beyond the
editor, the OPC publication audit applies `verify_authored` to the **original**
bytes of every replaced part, and 94 of 95 real fixtures are non-compact
(`FormattingWhitespace`); the only compact package in the corpus is the one
litchi wrote. This confirms as a blocker what change 0587's XLSX section listed
as an item to verify. No harness corpus can exercise any of it: none carries a
shared-string part, a worksheet relationship, an out-of-`sheetData` child or
non-compact XML, and the record specifies the generator shape that would.
`performance_claim: none`; `claim_authorized: false`. OLE2/OOXML remain active;
ODF is deferred until completion and iWork excluded.

## 0588: the cheapest read-side transform in the OOXML path, and a contract nobody had written down

`docs/GOAL.md`'s optimization order puts "eliminate unnecessary allocation and
copying" ahead of layout and algorithms, and the MCE codec had both in the same
function: two heap allocations per attribute, an owned `QualifiedName` per
qualified name on three paths, and an exact reservation per written run.
[0588](0588-mce-codec-namespace-emission.md) removes them **without changing a
single output byte** — 57.84% of the codec's instructions on a real Excel
worksheet, 52.08% of a public eager open plus one cell, p50 -43.88% wall clock —
which is the cleanest form the decision rules ask for: a change that cannot move
a refusal, cannot move a limit and cannot change what any consumer sees.

The more durable result is what stopped the rest. 0587 ranked XML-1 first and set
the precondition "confirm the processed stream is never published". Two writers
publish it, but the binding constraint turned out to be different and unnamed:
because the writer re-declares every in-scope binding on every emitted start tag,
**any element span of the processed buffer can be sliced out and parsed
standalone**, and `litchi-docx`'s paragraph, table and row views do exactly that.
That is a contract the codebase relies on and no record stated. It is now a test
(`every_element_span_of_the_output_is_namespace_self_contained`). It is also
already broken where the codec's fast path applies: on the one `test-data` `.docx`
whose `document.xml` declares no MCE namespace, `Paragraph::extensions()` already
refuses at the base revision, so `docs/GOAL.md`'s correctness-first rule is
engaged independently of any optimization. The audit's
"Apply layout, cache, or SIMD tuning only from measured hot loops" row gains a
sibling requirement from this: an internal representation that several crates
slice needs its slicing contract written down before it is optimized, or the
optimization is discovered by a single consumer test. OLE2/OOXML remain the
active priority; ODF stays deferred and iWork excluded.

## 0596: the eager DOC open stops decoding text four times, resolving each PAPX three times, and copying the WordDocument stream twice

Record: [0596](0596-doc-eager-open-terms.md).

**A ninety-instruction edit moved one harness scenario by fifteen percentage
points.** Change 0596 removed 19.8% of the instructions of the 1.6 MB DOC open
and 21.2% of its native cycles — the two agree because the removed term is a
`memcpy` of a 697,827-byte stream. But its first candidate build was 13% to 15%
worse at p50 on `doc_semantic_paragraph_count`, reproducibly, in two independent
windows against a floor under 0.3%, on a path where the query's own instruction
count moved by +0.03% and its native cycles by -0.66%. 99.98% of that query is
`ParagraphExtractor::count_paragraphs_in_range`, source the change does not
touch, which the candidate build inlines differently. Hoisting one FIB slice
resolution out of `get_all_subdoc_ranges` — worth 299 to 289 instructions per
call — turned that scenario into -2.63% and -0.32% in the next two windows. The
lesson for this audit is that at these scenario sizes harness percentiles rank
candidates rather than measure them, and that the A/A floor must be read per
scenario and per window, not once per run: in one window of this batch the floor
was -10.26% at p50 on `doc_semantic_open/large` and -0.11% on
`doc_semantic_paragraph_count/large` at the same time, and in another it reached
-80% at p95.

## 0599: an XLSB cell-value commit parses the workbook once, not twice — and a proven no-op parses it not at all

Record: [0599](0599-xlsb-commit-single-parse.md).

| priority | item | what it needs |
| --- | --- | --- |
| P1 (progressed) | Finish source-backed CRUD adoption across formats — the XLSB publication half | Change 0599 removes the duplicated whole-workbook readback from the XLSB cell-value commit and skips it entirely for a proven no-op, the first XLSB performance evidence in the program. What stays open is unchanged by it: XLSB cell reads still go through the eager `OpcPackage` door (0581's frozen ADR question), `SourceBackedWorkbook` still feeds only a sequential text writer, and the publication path for `apply_workbook_structure`, `apply_sparklines` and `apply_cell_watches` is unmeasured because **no harness selector opens it**. Registering one is the prerequisite for finishing the pattern in this crate. |
| P1 (progressed) | Cover the high-impact CRUD categories and real producers — XLSB corpus ceiling | The ceiling change 0587 recorded (largest `.xlsb` 22,715 bytes, one sheet, 48 stored cells, so a one-cell read and a full scan are 0.3% apart) is now liftable on demand: `tools/perf-baseline`'s `xlsb_synthetic_fixture` generates a deterministic workbook of any shape through the public writer and proves it reopens before writing it. Change 0599 measured 4-sheet fixtures of 15,000 and 90,000 cells with it. The gap it does **not** close is a *real-producer* large `.xlsb`: the generator emits no pivot cache, table, chart sheet, drawing, connection, external link or VBA project, which is exactly the surface whose parsing dominates the figures. Change 0599's saving therefore has a measured range, 7%–44%, and no single number. |

Supporting note for the audit body: change 0599 is a GOAL step 1 result —
*eliminate unnecessary work* — and its evidence separates the two methods rather
than leaning on either. The five `xlsb_crud` read-only cases never reach
`apply_cell_values`, so their paired delta (−3.54% to +1.93%) bounds drift and
code layout in the same window; subtracting that offset from the no-op delta
gives about −43% against the deterministic −43.94% of instructions. That window's
A/A floor was **worse than the host's standing figure** — |p50| median 0.84%,
p90 3.22%, max 10.04% over 64 control pairs — and the record respects it: the two
synthetic fixtures' commit deltas fall inside it and **nothing is claimed from
their timing**. No allocation, RSS, syscall, cold-cache, physical-device or
cross-platform measurement was taken; three allocations per commit are removed by
construction and none was counted, because `xlsb_crud` still has no
allocator-metrics hookup.

## 0593 — unchanged-member preservation stops paying for discarded work

Retained implementation against the audit's standing "close ZIP64/CFB and
output-source preservation intersections" row, on its unchanged-member
passthrough clause. Every OOXML save through `PackageWriter` previously paid a
whole-package XML serialization and audit pass for members it then copied
verbatim; the open-time canonical capture now proves the unchanged case in one
pointer comparison, and the content-types manifest is decided from provenance
before it is built. Correctness is established by whole-corpus byte identity,
not by inspection: 336 OOXML fixtures × 4 mutation scenarios = 1,344 published
digests and typed errors, before and after, with an empty diff, including 27
preserved save refusals and 8 preserved open refusals; plus 12 of 12 identical
editor-level rows, among them the two pre-existing `NotCompact` refusals 0587
reported. Measured per publish: native cycles −62.20% to −71.11% on four
fixture/scenario pairs, callgrind Ir −27.48% to −61.37% across ten, publish-only
p50 −61.52% to −68.68% against a p50 A/A floor of ±0.7%. The audit row this does
**not** close: no DOCX or PPTX semantic-editor save was measured through the
ordinary save path, because no example opens a real `.docx` or `.pptx`, edits
through the model and saves — 0587 §4 recorded that gap and it remains open, as
does the absence of an ordinary-save selector in `tools/perf-baseline`. Cold
cache, peak RSS and real-device fsync distributions remain unmeasured, and
`save(path)` remains fsync-bound at ~7.5 ms of an 8.3 ms median on this host.
`performance_claim: none`. OLE2 and OOXML remain active; ODF is deferred until
that goal completes and iWork is excluded.
[Change](0593-opc-publication-pristine-members.md);
[evidence](results/change-0593/README.md).

## 0595 — retained lean XLS frame loop and cheaper eager SST walk

Closes two rows of change 0587's ranked queue by GOAL step 1 (eliminate
unnecessary work) and step 3 (unnecessary allocation and copying), with no step
skipped: no layout change, no algorithm substitution beyond replacing an O(n)
re-summation with a maintained running total, no parallelism and no SIMD. The
decision rules are satisfied in order — before measurements captured on a
read-only detached checkout of the base, hypothesis and mechanism stated per
site, smallest coherent change, correctness/preservation/adversarial evidence,
after measurements with identical setup. Evidence tiers: **measured** for the 18
logical-counter cells, the six callgrind isolation pairs, the six `perf stat`
pairs and the 72 timing rounds; **modelled** for the two `SourceTextSheet::insert`
sites, which sit on the text-extraction path that `xls_source_attribution`
still has no selector for — change 0587's recorded measurement blocker, now
blocking a second batch. The audit row "apply layout, cache, or SIMD tuning only
from measured hot loops" is untouched. The standing instruction to price
pointer-chase work in cycles is honoured: native cycles fall 25.42% and 18.44%
on the `54016` open and one-cell against callgrind's 25.75% and 24.61%, and IPC
falls on five of the six cells because what was removed was the most
superscalar-friendly work in the loop. Host quiescence is not established — the
load average was 44.41 to 40.45 on 32 cores — so the noise floor is measured in
the same window (−2.64% to +3.06% at p50 over 36 same-binary comparisons)
rather than assumed.
OLE2/OOXML optimization remains active; ODF is
deferred until completion and iWork excluded. [Change and
limitations](0595-xls-frame-loop-and-sst-walk.md); [retained
evidence](results/change-0595/README.md).

## 0592: the first item off the 0587 queue, and what the harness could not see

`docs/GOAL.md`'s optimization order puts "eliminate unnecessary work" first, and
[0592](0592-docx-lazy-paragraph-index.md) is that step applied to the smallest
aligned item the 0587 survey found: a cache built on every DOCX main-document
open for the benefit of three of its dozen readers. It is retained on
deterministic evidence — callgrind isolation pairs and allocation counts on two
document shapes, plus a 332-document differential whose signature files are
byte-identical between legs — and `performance_claim: none`.

**The batch confirms one of the survey's falsification conditions rather than
refuting it.** 0587 wrote that DOCX-2 would be "falsified if the full-text
harness timer starts after `document()`, in which case the gain is invisible to
the selector". It does, and it is: `docx_semantic_full_text` moves −1.83% and
+0.10% across the two orders against an A/A floor of −0.80%. The gain is real and
is measured by the file-backed selectors, whose timed region does contain
`document()` (−40.22% at p50 on `docx_file_eager_full_text`, A/A floor 1.52%),
and by the instruction counts. The lesson for the goal's evidence rules is the one 0587
already recorded in another form: **a selector's timer boundary is part of the
claim's scope**, and three of the four `docx_semantic_*` DOCX read selectors
exclude the open from the interval they time. Anything that changes `document()`
must be measured with a probe or with the file-backed lifecycle selectors, and
this record says so rather than quoting the semantic selector's flat result as
"no effect".

**The A/A floor in this window was not the host's usual one.** With eight agents
building and measuring concurrently, the same-binary floor at p50 reached
**17.34%** on `docx_semantic_open` and **8.54%** on
`docx_file_eager_open_full_text_lifecycle`, so neither of those two carries a
result here; sub-microsecond selectors (`docx_semantic_one_paragraph` times a
130 ns array lookup) are below the clock's usable resolution entirely. The
counts, not the timings, carry this record's result, and the record says which is
which.

## 0591 — the ordinary DOCX edit scans the main part twice

Record: [0591](0591-docx-edit-single-scan.md).

`docs/GOAL.md`'s first optimization step — eliminate unnecessary work — applied
to the DOCX opened-document CRUD route that 0587 ranked fourth. Two of the four
whole-part scans per edit are removed because neither produced information the
caller did not already hold: one existed to feed a byte comparison that needs
only the bytes, and one rebuilt a layout that the splice determines exactly.
Both reuses are conditional and fall back to the original route, so no refusal,
no output byte, no limit and no validation moved; the semantic readback after
every paragraph rewrite and the `same_source` check on every patch application
both still run, which is what ADR 0005's mandatory validation requires, and ADR
0006 preservation is untouched because compaction — the part that would change
untouched bytes — was deliberately left alone. **Measured** with exact before and
after call counts (four scans and two `document_snapshot` builds per edit
lifecycle before, two and one after) and −33.27% instructions in the isolated
timed region of the 10,000-paragraph one-edit, matched by −36.98% p50 against a
2.06% A/A floor. What stays open: item DOCX-1(c) removes the last removable scan
but rewrites whitespace in paragraphs the edit never touched, so it needs the
owner's answer to 0587's ADR 0006 observation before a design record exists;
eleven other `with_rewritten_xml` sites still rescan, several of them rewriting a
paragraph inside a table or a block control, a case the derivation declines
outright; `document_snapshot` still copies the main part instead of sharing its
`Arc`, so the reuse proof still pays a full byte comparison; and DOCX-2's eager
paragraph index is untouched. No speedup, RSS, allocation, cold-cache or
real-producer claim follows.

## 0590 — PPTX opened-transaction revision reuse

Record: [0590](0590-pptx-opened-transaction-revision-reuse.md).

`docs/GOAL.md`'s first optimization step — eliminate unnecessary work — applied
to the PPTX opened-document CRUD route that 0587 ranked third. Two of the four
complete-package SHA-256 passes per lifecycle are removed by reusing values
already computed on identical content: the commit's own unsign-check revision,
and the commit's own capture, which the facade had been discarding in favour of
re-deriving it. Both reuses are conditional and fall back to the original path,
so no refusal, no output byte, no limit and no validation moved; ADR 0003's
complete-package revision binding is preserved because the fingerprint
definition is untouched and the durable patch encodings that carry revisions are
bit-identical. **Measured** −15.77% Ir per lifecycle on `pptx_eager_batch_edit_save`
with exact before/after call counts (six hashes and five captures per lifecycle
before, four and four after). What stays open: the remaining four hashes need
0587's item PPTX-1(c), which redefines the durable revision and therefore needs
a frozen design record and a magic bump for `LPRM0001` and `LPCP0002` before any
code; the notes-index memo of item (d) needs an owner for the memo and an
ADR 0013 invalidation rule; the cross-package copy path (PPTX-2) is untouched;
and no eager PPTX save timing is believable until the `open`/`from_vec` corpus
variant 0587 asks for exists, because the present corpora re-deflate all media.
No speedup, RSS, allocation, cold-cache or real-producer claim follows.

## 0589: an empty overlay hashes the artifact once, not twice — DOC and PPT source-backed opens lose half their SHA-256 work

Record: [0589](0589-ole2-snapshot-fingerprint-passes.md).

| priority | item | what it needs |
| --- | --- | --- |
| P1 (progressed) | Source-backed CRUD adoption: price the identity fence | Change 0589 closes the DOC/PPT half. The complete-artifact fingerprint designed by 0100/0105/0119 and read-coalesced by 0143 was never priced in CPU; it is now, and its value-identical duplicate is gone. What remains open is the **write** side: the commit and save paths gain the same halving on a no-op publication and are covered only by unit tests, because no `perf-baseline` selector opens a source-backed DOC or PPT path at all. Registering such a selector is the prerequisite for any further DOC-1 work, and for any attributable measurement of the DOC read path now that hashing no longer dominates its profile. |

Supporting note for the audit body: change 0589 demonstrates the measurement
discipline GOAL step 1 asks for — the saving is proven to be *unnecessary work*
rather than relocated work, because the deterministic read counts (`read_calls`,
`read_bytes`, `len_calls`, `version_calls` per operation) are byte-for-byte
identical between the two legs on all 38 fixtures, while the hash-pass count
halves. No allocation, RSS, syscall, cold-cache, physical-device or range-source
measurement was taken, so those rows of the Phase-1 baseline are untouched. The
largest fixture either measured path admits is 1.45 MB, so DIFAT-scale behaviour
remains unmeasured, and the host has SHA-NI, so a host without it would see a
larger relative saving that was not measured.

## 0587: the survey the goal asks for, taken after 586 changes

`docs/GOAL.md`'s DELIVERABLES section requires `HOTSPOTS.md` to carry "ranked
opportunities by expected total CRUD impact, risk, and ADR compatibility", and
its REPOSITORY EXPLORATION section requires the complete data path to be mapped
before it is optimized. The ranked queue this program carried was written before
change 0190 and headed "provisional until baseline measurements are recorded";
every later record ranked its own area. [0587](0587-remaining-opportunity-survey.md)
discharges the deliverable: eleven parallel surveys mapped every OLE2 and OOXML
area of the path at `2fc5fc657`, screened each candidate against the record set,
and produced one queue of 36 items with evidence tiers, prerequisites and an ADR
compliance matrix filled in advance. It authorizes nothing and lands nothing.

**Two findings bear on the goal's evidence rules directly.** The first is that
the goal's clause "every claim must be scoped to a named scenario, corpus,
machine, build, and metric" has been honoured to the letter and missed in
spirit: the corpora were named, and they were synthetic in a way that excludes
the paths real Office files take. The largest OOXML read cost this program has
found — the MCE codec's 16.9× expansion of a real Excel worksheet, 12.9× on an
open plus one cell — is invisible on every generated corpus, and three landed
optimizations (0525, 0546, the selected-cell stream) are disabled on most real
files by the same markers. The corpus requirement in PHASE 1, "files from
multiple real-world producers where licensing permits", is the clause that was
under-served, and a real-producer shape in the generator is now the first
prerequisite of the queue. The second is methodological: callgrind runs
SHA-256 in software, so every profile that could have found the whole-artifact
hashing in the DOC and PPT snapshots, the PPTX opened transaction and the
cross-package copy either discounted it as harness cost or was never taken on
those paths. Hashing must be priced natively, with `perf stat`, the same way
0579 established for pointer chases.

**What this record does not discharge.** No measurement in it is a paired,
controlled timing; every fresh figure is a single-leg count or instruction
profile with the stated caveats, and the top item is measured on one fixture.
Still required by the goal and unchanged: cold-cache and physical-device
distributions, real range sources, peak RSS for read paths, concurrency scaling,
real-producer breadth, cross-platform confirmation, and coverage-guided fuzzing
on this host — the evidence-gap table in 0587 ranks these by which
DEFINITION OF DONE clause each blocks, and finds the CI smoke check compares a
run against a byte-copy of itself.

OLE2/OOXML stay first; ODF is deferred until that goal completes and iWork is
excluded; the broad goal remains active.

## 0585-0586: one hint pays 65.5%, the other pays nothing, and the difference is scope

`docs/GOAL.md`'s DEFINITION OF DONE requires that "selective reads perform work
proportional to mandatory metadata plus accessed content". On the OLE2 side the
shared-string path violated that in a way no record had named: a source-backed
XLS text extraction rewalked the `Workbook` allocation chain from its first
sector **once per string cell**, so the work was proportional to
*strings × chain depth* rather than to the strings read.
[0585](0585-cfb-resumable-cursor-construction.md) removes **65.5%** of those
links corpus-wide, measured, with byte-identical output on all 86 measurable
fixtures.

The two records together answer a question the program had not asked explicitly:
**when does change 0579's mechanism pay?** 0585 pays because one text extraction
resolves hundreds or thousands of strings through a document-scoped resolver.
[0586](0586-doc-paragraph-hint-rejected.md) pays *exactly nothing* — 3,991 chain
links before and after, on every measurable fixture — because its hint was scoped
to a single `resolve_paragraph` call, and one call issues too few reads for a
prefix to exist. Same mechanism, same crate boundary, opposite outcome, decided
by the hint's lifetime rather than by the format. That is the reusable lesson,
and it is why 0586 is retained as a rejection rather than discarded.

**Two methodological points are worth carrying forward.**

The first is that the measurement validated itself before its result was
believed. The instrumented counter carries a `pre-0579` control leg that
reproduces change 0579's published pre-change figure of 5,796 chain links
exactly, and a `0579-only` leg that reproduces its published post-change 2,099
exactly. A switch that reproduces a previously published number on both sides is
measuring what it claims to. Without that control, `pre-0579 → this-change` would
have credited 0585 with 0579's saving.

The second is that a prediction was superseded rather than confirmed. Change
0584's byte-layout model forecast 74.4% corpus-wide against the 65.5% measured;
on the largest fixture its naive SST walker had decoded only 829 of 16,055
entries, so its backward-step distribution came from a 5% sample. The model was
useful for *targeting* and wrong for *sizing*, and the record says so rather than
quoting whichever number is larger.

**What this batch does not discharge.** No timing, cycle, cache-counter,
allocation, peak-RSS, cold-cache, range-source, concurrency or cross-platform
measurement was taken for 0585. Chain links are a count of dependent loads, and
`GOAL_AUDIT`'s own standing note — that 0579 removed 1.24% of instructions and
6.19% of cycles on one fixture — means the relationship between 65.5% and any
wall-clock figure is unknown in both directions. The per-sheet cursor term
(28,143 links on a 16-sheet fixture, ~91% resumable) is priced and unimplemented.
`ConditionalFormattingSamples.xls` could not be measured end to end because the
library refuses its text extraction, so 0584's 703,937-link prediction for it
describes a scan that is not reachable on that file.

OLE2/OOXML stay first; ODF is deferred until that goal completes and iWork is
excluded.

## 0584: the OLE2 ranking is refreshed, and two long-standing coverage gaps close

`docs/GOAL.md`'s REPOSITORY EXPLORATION section requires the complete data path
to be mapped before it is optimized, and its DELIVERABLES section requires
`HOTSPOTS.md` to carry confirmed bottlenecks, disproven hypotheses and ranked
opportunities. [0584](0584-ole2-profile-at-head.md) refreshes that ranking for
OLE2 at `e927e139c` and closes two gaps the program had carried for a long time.

**The first is DOC and PPT.** Change 0574 states plainly that "no DOC or PPT
scenario was measured at all", and that remained true through change 0583. It is
now false. The result is not yet actionable — DOC cost is not proportional to
document size, a 65 KB fixture costing 9% more per open than a 1.6 MB one, and no
single symbol is large enough to explain it — but the question is now posed with
numbers behind it rather than absent.

**The second is that the profile now attributes per caller, not merely per
operation.** 0579 could count chain links per open; 0584 splits a flagship
one-cell query's 4,428 links into 2,099 from `read_stream_range_hinted`, 2,238
from `stream_cursor_at` and 91 from `normalize_state`. That split is what makes
0579's own open question answerable.

**One methodological point is worth carrying forward.** The record's ranking
would have been wrong in three of five rows without screening each candidate
against the existing record set first. Lazy SST indexing looks like the single
largest OLE2 opportunity in the profile at 49.5% of one fixture's open, and
change 0576 has already deferred it because it moves *when* a malformed SST is
refused. Skipping uninterpreted globals bytes looks like 34% of the flagship
open, and its density gate is the one change 0568 recorded as untestable on any
real fixture in this corpus. A profile ranks work; only the record set says
whether that work is available.

The record also restates the standing measurement caveat in measured form:
instruction counts rank work, not latency, and callgrind's per-byte `rep`
accounting makes bulk-copy shares upper bounds. `performance_claim: none`, and
no candidate it names is authorized by it.

Still required by the goal and unchanged by this record: cycles, cache and branch
counters for these scenarios, cold-cache and physical-device distributions, peak
RSS for read paths, concurrency scaling, real-producer breadth, cross-platform
confirmation, and coverage-guided fuzzing on this host.

OLE2/OOXML stay first; ODF is deferred until that goal completes and iWork is
excluded.

## 0577-0583: the OOXML selective-read goal is met, and the approval flow is corrected

This batch closes the OOXML half of the goal's definition-of-done clause that
change 0572 showed was violated. "Selective reads perform work proportional to
mandatory metadata plus accessed content" was breached on the OOXML side in a way
the clause does not name: the work was proportional to the archive's **member
count**, paid in positional requests before the first payload byte. Change 0573
halved it and [0580](0580-zip-target-scoped-strict-layout.md) takes a single-member
read of a 132-member workbook from 264 requests to **20**. The strict-layout proof
is no longer the dominant term of a selective OOXML read.

That result required narrowing a refusal, and how that decision was reached is
the part of this batch the goal's evidence rules bear on most directly. Change
0575 established that the invariant had **no ADR basis, no threat model and no
commit rationale**, and the project owner approved narrowing it. Change
[0582](0582-zip-strict-scope-differential-fuzz.md) then measured the delta and
found the approval had been sought on **3.3%** of it: the request was framed
around one adversarial overlap witness, while the change drops every
local-versus-central consistency check on every record the caller does not read.
The owner re-approved on the corrected, measured delta. No accepted ADR was
weakened at any point, and the one rival design that would have breached ADR 0005
was recorded and declined rather than implemented.

The goal's rule 12 held on its own terms. Change 0582's harness found one genuine
weakening of a malformed-input defence, and [0583](0583-zip-local-size-span-bound.md)
closed it in the same batch, returning 32,472 verdicts to refusal while making
none newly readable.

One required gate is **not** satisfied and is recorded as outstanding rather than
waived: `docs/GOAL.md`'s verification list names existing fuzz targets, and
`parse_zip` cannot run on this host for want of `cargo-fuzz` and a nightly
toolchain. The deterministic differential harness built in its place covers the
changed surface more exhaustively than the fuzz target does per input, but has no
coverage feedback, no mutation engine and no sanitizers, and the record says so.
It also proved its own limits: a regression it did not find was found by reading
the code.

On the OLE2 side [0579](0579-cfb-resumable-chain-walk.md) continues the
instruction-side work change 0574 opened, and corrects the method that ranked it.
Instruction share mis-ranked this opportunity in both directions, because the
removed work is a dependent-load pointer chase; the rejected candidate removes
more instructions and is slower. Future OLE2 rankings in this program should
price cycles, not instructions, where a pointer chase is involved.

Three goal items advance without code. [0577](0577-ooxml-open-relationship-parts.md)
establishes that the OOXML open's relationship reads are **mandatory**, not
incidental, so the remaining 89-request open cost is a cost to be reduced rather
than eliminated, and designs that reduction to 10.
[0578](0578-zip-passthrough-is-already-bounded.md) refutes the standing suspicion
against the "unchanged media without logical-byte copies" clause: the ZIP save
path already satisfies it, flat at 12,971 bytes across a thousandfold range of
member sizes. [0581](0581-opc-package-retention.md) finds what the goal's bounded
-memory clause actually implicates — **the ordinary documented open-then-save path
in every OOXML format is the eager one**, differing 254-fold in peak from the
bounded path while emitting byte-identical output.

Still required by the goal and unchanged by this batch: coverage-guided fuzzing on
this host, cold-cache and physical-device distributions, real range sources rather
than a simulated transport, peak RSS for read paths, concurrency scaling,
real-producer breadth, and cross-platform confirmation. Change 0580's reverse-order
regression, up to 2n−1 reads, is recorded and unaddressed. Change 0577's designed
coalescing is blocked on a read-side accessor that does not exist. Change 0581's
candidates are unimplemented by design.

OLE2/OOXML stay first; ODF is deferred until that goal completes and iWork is
excluded; the broad goal remains active.

## 0572-0576: the range-source gap is closed, and the OLE2 open stops being I/O

The previous batch's audit listed "remote and range-source behaviour" among the
items this goal still required and no batch had supplied. [0572](0572-ooxml-range-source-attribution.md)
supplies it for OOXML. It is a model rather than a device — a simulated
transport, not a network — and the record says so, but it is the first evidence in
this program of what the OOXML read path costs a caller-supplied source, captured
through the explicit provider the ADRs require and with no ambient networking
added. All five of its frozen gates pass, and its request sequences are identical
across transports and repeats.

What it found bears directly on the goal's definition of done. The clause
"selective reads perform work proportional to mandatory metadata plus accessed
content" is violated on the OOXML side in a way the clause's own wording does not
name: the work is proportional to the archive's **member count**, not to its
uncompressed size, and it is paid in requests before the first payload byte.
[0573](0573-zip-single-local-header-read.md) halves that, and
[0575](0575-zip-lazy-strict-layout-design.md) establishes both how much further it
can go and that it cannot go to O(1), with a proof rather than an estimate.

0575 also surfaces something the program should decide rather than inherit. The
archive-wide strict-layout proof that costs those requests has **no documented
rationale**: both introducing commits have empty message bodies, no accepted ADR
mentions it, there is no threat model, and the ordinary read path already bypasses
it. The goal forbids reinterpreting an accepted ADR to enable an optimization; it
does not require preserving an undocumented behaviour forever. The design records
the conflict, rejects the one candidate that would breach ADR 0005, and leaves the
remaining decision to human review with a reproducible witness. No production
change was made on the strength of it.

On the OLE2 side the goal's target is now met further up the stack.
[0574](0574-ole2-next-opportunity-survey.md) establishes that changes 0565, 0568
and 0570 moved the bottleneck **off I/O**: a source-backed open now spends 75.9%
of its time outside the source. That is a goal milestone worth naming, because it
means the remaining OLE2 work is instruction and allocation work rather than read
scheduling, and the optimization order in `docs/GOAL.md` moves accordingly from
step 2 to step 1 — eliminating unnecessary work.
[0576](0576-xls-sst-scan-without-materialization.md) does exactly that, removing
up to 88.2% of the instructions in an open without changing one byte of I/O or one
character of an error message.

Still required by the goal and unchanged by this batch: cold-cache and physical
device distributions, real range sources rather than a simulated transport, peak
RSS, concurrency scaling, real-producer breadth, and cross-platform confirmation.
Three new limitations are recorded rather than presented as covered. Only one PPTX
fixture in the corpus has enough slides to exercise a genuine middle slide, and
one DOCX fixture extracts zero bytes under a scenario labelled "full text". Ten of
93 XLSX fixtures refuse the cell scenario outright, two of them fixtures the frozen
plan itself named. And 0576's allocation-failure error is unreachable on the new
measure path; the divergence is bounded and stated rather than left silent.

A new OOXML hotspot is recorded and not yet addressed: the open reads every
`*/_rels/*.rels` part in the package, so a 132-member workbook pays 89 of its 354
requests before a cell read begins. Compressed-entry passthrough on save is also
still a logical-byte copy — `write_precompressed_file` takes a fully materialized
`&[u8]` — which the goal's "unchanged large media … without unnecessary
decompression or logical-byte copies" clause names directly.

OLE2/OOXML stay first; ODF is deferred until that goal completes and iWork is
excluded; the broad goal remains active.

## 0568-0571: the XLS selective-read goal is met, and a recommendation is withdrawn

This batch closes the other half of what change
[0325](changes/0325-cfb-frame-transaction-rejected.md) had in scope.
[0568](0568-xls-worksheet-window.md) takes a source-backed one-cell query from
319 logical reads to 61, its worksheet-scan component from 266 to 8, while open
and list stay byte-for-byte identical. With change
[0565](0565-xls-globals-single-pass.md) before it, the goal's target for this
path — work proportional to mandatory metadata plus accessed content — is now met
for both the globals and the worksheet scan, and in the worksheet case with no
over-read at all, because the sheet boundary is validated at open.

[0570](0570-cfb-fat-run-batching.md) collapses contiguous container FAT runs into
single reads, and demonstrates why a corpus can answer "how often" without
answering "how much": across 98 real fixtures it removes 30 reads, and on one
large synthetic container it removes 257 from a single open.

On the OOXML side the batch is mostly a retraction. Change
[0567](0567-ooxml-single-index-per-open.md) had recommended a zero-code
documentation fix; [0569](0569-ooxml-detect-then-open-priced.md) measured it and
**rejected** it, finding that one of the named openers does no detection at all,
that the classification is computed and discarded so the workbook opener cannot
tell a Word document from an image, and that the pattern is unavoidable for any
caller dispatching across the three facade types. [0571](0571-ooxml-prepared-source.md)
implements the fix that measurement pointed to instead, taking a detect-then-open
sequence from two archive indexes to one and saving 35.9 to 201.7 microseconds,
and repairs the discarded-classification defect as a separate, explicitly
non-performance change in the same diff.

Two findings in this batch are about the program's own instruments rather than
the library. A test runner aborting at the first failing target had been hiding
every later target, and a lint gate was already red at HEAD on crates the
standing gate list does not cover. Both are fixed. The noise floor is now
measured in the same window as the result it qualifies, rather than assumed:
p50 4.10% and p99 13.70% for one binary against itself, which is what makes the
batch's single p99 review trigger legible as quantization rather than regression.

Still required by the goal and unchanged by this batch: cold-cache and physical
device distributions, remote and range-source behaviour, peak RSS and allocation
accounting, concurrency scaling, real-producer breadth, and cross-platform
confirmation. The density gate change 0568 adds cannot fire on any input in this
repository and its only coverage is synthetic, which is recorded as a limitation
rather than presented as tested. A complete text extraction of the flagship XLS
fixture remains unmeasurable through the public text API because it fails
identically on both sides.

OLE2/OOXML stay first; ODF is deferred until that goal completes and iWork is
excluded; the broad goal remains active.

## 0565-0567: the largest remaining OLE2 gap is closed, and an OOXML hypothesis is disproved

[0565](0565-xls-globals-single-pass.md) closes the gap
[0564](0564-xls-open-read-attribution.md) identified as the biggest measured
remaining item on the OLE2 priority path. A source-backed XLS open falls from 655
positional reads and 636 freshness observations to 53 and 34, with each globals
byte read once instead of the headers being read twice. Change 0564 recorded that
this conflicted with four tested contracts and, per this goal's rule, stopped.
This batch made that decision explicitly rather than by side effect: each
contract is replaced by an enforced invariant, three further consequences are
written down, and a survey of 104 fixtures establishes that the conservative
alternative — clamping before any read-ahead — saves 8% of the reads where the
implemented schedule saves 95%.

The goal's target for this path, work proportional to mandatory metadata plus
accessed content, is now met for the globals: the scan reads the globals once
plus a bounded read-ahead of at most one fill, measured at zero bytes on the
flagship fixture and at most 3,949 across the surveyed corpus.
[0566](0566-xls-worksheet-window-design.md) designs the same treatment for the
worksheet scan, the other half of what change 0325 had in scope, and finds it
easier because the sheet boundary is validated at open.

[0567](0567-ooxml-single-index-per-open.md) **disproves** this goal's hypothesis
12 inside one library call: detection does not repeat container indexing, because
the facade builds the package once and hands it on. It also corrects two earlier
records, one of which had left a follow-up open on a read that turns out not to
belong to this library at all. The OOXML repeated indexing that does exist is
across two calls, and the cheapest fix is documentation rather than code.

An important limitation is now quantified rather than assumed. This host's A/A
noise floor drifts several percent at p50 and over ten percent at p99 with one
binary against itself, so the 5% review trigger is not comfortably above the
tail-statistic floor, and the nanosecond-scale semantic selectors carry no
information in either direction.

Still required by the goal and unchanged by this batch: cold-cache and physical
device distributions, remote and range-source behaviour, peak RSS and allocation
accounting, concurrency scaling, real-producer breadth, and cross-platform
confirmation. The goal's own list names a simulated high-latency range source,
and both change 0561's conclusion and this batch's owned-source rows point at it
as where the remaining I/O work pays.

OLE2/OOXML stay first; ODF is deferred until that goal completes and iWork is
excluded; the broad goal remains active.

## 0564: the largest remaining OLE2 target, and why it is blocked

[0564](0564-xls-open-read-attribution.md) locates the biggest measured remaining
gap on the OLE2 priority path. A source-backed XLS open costs 655 positional
reads and 636 freshness observations against an owned source's identical logical
work, and 621 of those reads are a four-byte-at-a-time BIFF header scan whose
bytes the following bulk read fetches again in full. Removing it is worth
roughly the entire file-source penalty.

It is blocked, and the blockage is a genuine design decision rather than a gap
in the work. Four tests assert the current shape deliberately: the exact open
range list, that a `FILEPASS` payload is never read, that a skipped payload is
never read, and that open never touches a sheet body. The goal's own rule is to
record such a conflict rather than implement through it, so this record
enumerates the four things a fused pass would have to establish — a
buffered-but-never-published rule, a bound on reading past an as-yet-unknown
`global_end`, how the per-record limits fold into a chunk, and an accepted
reduction in freshness granularity from 621 observations to about 9.

This record also reopens a line the program had closed. Change 0325 rejected a
prefetch candidate and concluded that the open and list selectors were
"zero-opportunity"; that was true of the candidate it tested and false of the
path, because 95% of open's reads are in a loop that record never examined.

Still required by the goal and unchanged by this batch: cold-cache and physical
device distributions, remote/range-source behaviour, peak RSS and allocation
accounting, concurrency scaling, real-producer breadth, and cross-platform
confirmation.

OLE2/OOXML stay first; ODF is deferred until that goal completes and iWork is
excluded; the broad goal remains active.

## 0563: the OOXML syscall line of work, and what it is worth

[0563](0563-opc-single-warm-part-observation.md) removes one source observation
from every warm OPC part read and pins the warm and cold counts as enforced
invariants. It is a real work reduction, but the batch's more important result
is a steer for the goal itself.

Change [0493](changes/0493-managed-opc-source-read-ahead.md) had already built
and measured the coalescing mechanism the 0561 traces point at, as an opt-in
`litchi-opc` forward-start window. It collapsed physical reads from 19 to 3 per
DOCX lifecycle and measured +0.98%/−2.01% p50 on a zero-delay local source
against −84.87%/−82.17% on a 1 ms-service range source, concluding that the
local arm "is not evidence of a useful local-source speedup".

So the OOXML repeated-read finding is exact as a count and close to worthless as
a local-latency target. The goal's own list of required scenarios includes "a
simulated high-latency range source with configurable latency, bandwidth,
request overhead, and maximum range size", and that is where this work pays.
The next OOXML step should therefore be to measure the existing
`SourceReadPolicy::forward_start` window on the file-source selectors against
that transport — which costs no production code — rather than to remove more
warm-cache syscalls. Enabling it is a policy decision, not a default change:
`crates/litchi-opc/tests/source_read_ahead.rs:433` pins `exact()` as the default
and requires existing constructors to stay uninstrumented.

Still required by the goal and unchanged by this batch: cold-cache and physical
device distributions, peak RSS and allocation accounting for these paths,
concurrency scaling, real-producer breadth, and cross-platform confirmation. A
file-backed repeated-part-read selector is a newly identified coverage gap.

OLE2/OOXML stay first; ODF is deferred until that goal completes and iWork is
excluded; the broad goal remains active.

## 0561-0562: the OOXML read pattern, attributed and first fix

[0561](0561-opc-repeated-positional-reads.md) and
[0562](0562-zip-descriptor-read-once.md) open the goal's OOXML I/O workstream
with measurement rather than assumption. About nine in ten positional reads on
file-backed OPC paths re-read a range the same child already read; each member
read costs four reads, two of which recover values the archive index already
holds. Removing the duplicated data-descriptor resolution cuts whole-child
positional reads by 24.84% on the PPTX captures and 15.79-16.31% on the DOCX
ones without changing which ranges are read.

The goal's target for this path — work proportional to mandatory metadata plus
accessed content — is not yet met. About 88% of reads still repeat a range, but
a later segmentation of the same traces **retracted** this record's first
projection: those repeats are across the ten archive constructions each traced
child performs, not within one archive instance, where every member's header and
descriptor is already read exactly once. A per-entry memo would therefore remove
zero reads on this corpus. The correction and its evidence are in change 0561.

The follow-ups the measurement actually supports are: coalesce the local header,
payload and descriptor into one bounded read per member read, which is the only
candidate that reduces first-read cost and which the payload size distribution
supports; and stop reconstructing the package on every open, which ADR 0011
makes `litchi-opc`'s to own. Neither is implemented or measured.

Still required by the goal and unchanged by this batch: cold-cache and physical
device distributions, remote/range-source behaviour for OPC consumers, peak RSS
and allocation accounting, concurrency scaling, real-producer breadth, and
cross-platform confirmation.

OLE2/OOXML stay first; ODF is deferred until that goal completes and iWork is
excluded; the broad goal remains active.

## 0560: XLS freshness observations reduced to one

[0560](0560-xls-single-observation-freshness.md) continues the goal's first
optimization rule on the OLE2 spreadsheet path. A retained XLS metadata query
took six source observations; it now takes one, and the collapse is strictly
equivalent because nothing between the removed observations consumed a source
byte. A counted-observation test makes the invariant enforceable.

The measurement also re-established where the remaining XLS file-source cost
lives: after changes 0558 and 0560 the path still takes one observation and one
positional read per BIFF record, so the open/list/one-cell gap over an owned
source is now dominated by read count rather than probe count. Bounded span
batching is the next hypothesis for that gap, and it must preserve the exact
source-byte accounting the program's evidence depends on.

A separate OOXML diagnostic recorded during this batch is not yet addressed: one
file-backed `pptx_file_source_selected_slide` operation issues 13,608 positional
reads over only 1,262 distinct byte ranges, so 90.7% of its physical read calls
re-read bytes already read, and 25% are immediate back-to-back duplicates of the
identical range. Attribution is in progress; no fix is claimed.

Still required by the goal and unchanged by this batch: cold-cache and physical
device distributions, remote/range-source behaviour, peak RSS and allocation
accounting, concurrency scaling, real-producer breadth, and cross-platform
confirmation.

OLE2/OOXML stay first; ODF is deferred until that goal completes and iWork is
excluded; the broad goal remains active.

## 0558-0559: OLE2 read fencing and name folding

[0558](0558-ole2-single-read-fence.md) and
[0559](0559-cfb-ascii-simple-uppercase.md) advance the goal's first two
optimization rules — eliminate unnecessary work, then unnecessary I/O — on the
shared OLE2 substrate rather than on one selector.

Progress: a confirmed file-source hotspot is halved. The shared CFB reader
observed `ReadAt::version()` twice per read, so every positional read on a
`FileSource` cost two `statx` syscalls; one selective XLS open issued 1,266
observations for 655 reads. Keeping one fence per read cuts observations and
whole-child `statx` by about 49% with read calls and read bytes byte-identical,
improving all 48 paired cells. Deterministic branch attribution then exposed
CFB directory-name case folding as the leading branch-miss owner in OLE2
edit/save; an exhaustively proven ASCII branch removes it. Against the pre-0558
baseline, 222 of 376 guardrail comparisons improve in both directions with two
p99-only review triggers.

Still required by the goal, unchanged by this batch: cold-cache and physical
device distributions; remote/range-source behaviour for CFB consumers; peak RSS
and allocation accounting for these paths; concurrency scaling and Amdahl
analysis; real-producer breadth beyond the one XLS fixture and the generated
CFB/OLE2 corpora; cross-platform confirmation. The remaining file-source gap
over an owned source is now dominated by the surviving one probe per read plus
positional read cost, so bounded span batching — which reduces the number of
reads rather than the probes per read — is the next hypothesis for that gap.
The `litchi-xls`, `litchi-opc` and `litchi-doc` freshness sites that take two or
three observations with no intervening read are a separate, strictly equivalent
class and are not addressed here.

OLE2/OOXML stay first; ODF is deferred until that goal completes and iWork is
excluded; the broad goal remains active.

## 0551: layout handoff design prerequisites narrowed

[0551](changes/0551-xlsx-layout-handoff-feasibility.md) adds reproducible
scanner partitions, source-bound edit-row coverage and a field/refusal audit.
This is diagnostic progress toward eliminating repeated OOXML work, not a
runtime improvement or a completed equivalence proof. Compact cell metadata,
error-order differential tests, managed/no-op retention and fresh matched
workflow admission remain required. OLE2/OOXML stay first, ODF deferred and
iWork excluded; the broad goal remains active.

## 0550: current XLSX commit attribution

[0550](changes/0550-xlsx-source-commit-attribution.md) Fresh source-backed
XLSX profiles change the next action toward eliminating a second source layout
scan. Four metadata checks and capture oracles pass; no new broad test suite
or optimization is claimed. Single-sheet API profiling, managed/refusal
performance, route/copy counters, physical providers and scaling remain gaps.
OLE2/OOXML stay first; ODF is deferred and iWork excluded.

## 0549: checked bitset candidate rejected

[0549](changes/0549-cfb-checked-test-and-mark.md) Fresh matched evidence
rejects another CFB loop replacement and preserves its exact semantic tests
and raw measurements. This is diagnostic progress, not a production speedup.
OLE2/OOXML remain active; ODF is deferred until that goal completes and iWork
is excluded.

## 0548: measured rejection; OLE2/OOXML goal stays active

[0548](changes/0548-ole2-checkpoint-collector-rejected.md) adds a public malformed
CFB guard and rejects the checkpoint implementation on frozen native and
instruction gates. Exact baseline production is restored while candidate
snapshots and all evidence remain available. This is diagnostic and validation
progress, not a runtime performance gain. ODF remains deferred until the
OLE2/OOXML optimization goal completes; iWork is excluded.

## 0545: bounded diagnostic progress, broad goal remains active

[0545](changes/0545-xlsx-scanner-diagnostic-rejected.md) rejects a dense-only scanner win using sparse and early-exit evidence
before another integration campaign. All captures, 18 individual reviews and
replay checks remain; the owned build/fixture tree is cleaned and the batch
committed. OLE2/OOXML performance is the active priority. ODF work is deferred
until that goal completes; iWork remains outside this workstream.

## 0544: oracle and cap benchmark progress; runtime restored

[0544](changes/0544-xlsx-event-preflight-rejected.md) completes a fresh preflight campaign and direct XML-reader oracle.
The candidate fails native and cap admission, so production is restored and
only the measured benchmark example is retained. All 450 adverse/drift rows
remain explicit. OLE2/OOXML goal remains active; ODF waits until it completes.

## 0543: valid-input regression prevents retention

[0543](changes/0543-xlsx-shared-traversal-cap-rejected.md) completes the main pilot and conditional lanes, then rejects the
candidate on a supplemental valid event-boundary sweep and warning-denied
Clippy. Production is restored; the public post-EOF regression test and complete
evidence remain. A conservative admission preflight is an unmeasured proposal.
OLE2/OOXML optimization remains active; ODF is deferred until that goal completes.

## 0542: concrete progress; broad OLE2/OOXML goal remains active

[0542](changes/0542-xlsx-shared-traversal-refusal-rejected.md) adds a measured valid/late-error planning guard and identifies why a
shared XLSX traversal candidate cannot yet be retained. Valid native and memory
gates pass, but three late-raw latency rows and a retained-validator overlap
require a new candidate. Production is restored; the unmeasured follow-up and
full evidence are retained. ODF optimization remains deferred until the
OLE2/OOXML optimization goal completes.

## 0541: public XLSX planning first-error guards

[0541](changes/0541-xlsx-planning-error-order-guards.md) adds the behavioral guard. The
shared-traversal prerequisite now has public first-error controls, combined
validation/preprocessing/parser cases, retries and cross-sheet ordering. This adds tests
rather than a runtime optimization. OLE2/OOXML remain active, ODF deferred and iWork
excluded; native improvement is still unproven.

## 0540: XLSX validation/parser boundary attribution

[0540](changes/0540-xlsx-validation-parser-boundary.md) records the fresh attribution.
Fresh profiles and a combined source audit establish the next larger XLSX target after
rejected 0539: shared worksheet traversal with separate validation/parser states and
preserved first-error ordering. No implementation is yet admitted. Exact source equality
reuses six final 0539 quality checks and 1,292 tests. The broad goal stays active;
OLE2/OOXML first, ODF deferred, iWork excluded.

## 0539: XLSX transient ownership candidate rejected

[0539](changes/0539-xlsx-transient-ownership-rejected.md) records the matched
experiment. The 0537 candidate has now been measured using the 0538 planning observer
and rejected under frozen native gates. Baseline, candidate and restored suites each
pass 1,292 tests, with final quality checks passing. Next is a source and profiling
audit of the worksheet validation/parser boundary. The broad goal remains active;
OLE2/OOXML take priority, ODF is deferred, and iWork is excluded.

## 0538: close the XLSX planning allocation evidence gap

[0538](changes/0538-xlsx-planning-allocation-metrics.md) implements the
measurement prerequisite identified in 0537, with real normal/allocator binary
checks and acquisition-order phase alignment. The initial filesystem test's
`/tmp` quota failure is retained and its successful owned-TMPDIR rerun is
recorded separately. Production remains unchanged. Next is fresh matched
measurement of the transient attribute candidate; the broad goal remains
active, prioritizing OLE2/OOXML with ODF deferred and iWork excluded.

## 0537: raw XLSX attribute ownership and planning allocation gap

[0537](changes/0537-xlsx-transient-attribute-attribution.md) replays four sealed
historical planning profiles and binds 13 relevant current source/lockfile
inputs. Raw attribute decoding forces temporary strings to become owned; its
direct allocation-child instruction costs are retained separately from nested
decoder/scan totals. A draft keeps transient references and numeric metadata
borrowed while preserving owned cell-type retention, normalization and complete
checked scans. No runtime optimization or fresh performance claim is accepted.

The next measurement must first cover planning allocations: the current region
starts at commit. Fresh matched native, phase/profile, allocation and eager
readback guards remain required. OLE2/OOXML stays first; ODF is deferred and
iWork is excluded. This diagnostic does not complete the broader goal.

## 0536: reject collector cold-error helpers after matched measurement

[0536](changes/0536-cfb-collector-cold-error-layout-rejected.md) restores the
baseline runtime after testing eight private CFB diagnostic helpers. All eight
primary XLS p50 comparisons improve, but only four reach the frozen 3% gate;
XLS constructor inclusive Ir also increases slightly in both repeats. The
collector shrinks from 1,436 to 1,238 bytes while self Ir changes by less than
0.005%. No runtime optimization or speedup is retained.

All 24 allocation pairs retain identical allocation-call, reallocation-call,
allocated-byte and incremental-peak vectors. Evidence retains 48,000 native
and 1,440 allocator samples, 80 timed and 12 setup profile dumps, 32 matched
adverse flags and 60 same-build variations. Candidate and restored-final
quality each pass 14 gates and 4,382 test executions (8,764 new executions).
All 96 receipts pass serial/ABBA verification. Full postcleanup replay passes,
owned temporary files are removed, and the evidence is sealed. OLE2/OOXML
remains active, ODF is deferred until that goal completes, and iWork is excluded.

## 0535: CFB collector attribution advances the goal; completion remains open

[0535](changes/0535-cfb-collector-instruction-attribution.md) completes a
CFB collector diagnostic with unchanged final-0534 production and harness
source. All 19 serial receipts, 4,000 native samples, 20 timed instruction
dumps and two setup dumps passed evidence verification. The mapped ordinary
loop accounts for 99.9325% of XLS collector self Ir and 99.9723% of CFB
collector self Ir. Checked bitset updates and vector append already occur
inside that loop. A cold-error-layout experiment remains conditional on
fresh native, allocation and correctness gates; no optimization or speedup
is accepted here.

All 13 same-build variation flags remain retained, including CFB few-large
p50 drift of +10.0959%. Exact-source custody confirms reuse of 14 prior
quality gates and 4,382 Rust test executions; no new Rust tests are claimed.
Full postcleanup verification and five verifier tamper probes passed. Both
owned scratch paths were removed, and the evidence is sealed. OLE2 and OOXML
remain the active optimization priority and the broader goal remains open.
ODF is deferred until that goal completes; iWork is excluded.

## 0534: CFB paired-prefix candidate rejected; full goal remains open

[0534](changes/0534-cfb-physical-paired-prefix-rejected.md) advances the
active OLE2/OOXML investigation with a complete paired role/FAT
reconciliation comparison, then rejects the candidate because all eight
primary XLS p50 rows regress in both repeats. The allocator guard passes and
the source contract tests remain on the restored baseline, but no runtime
speedup is retained. The evidence keeps 48,000 native samples, 1,440
allocation samples, 16 profile children, 80 timed dumps, 12 setup dumps, and
all 153 review flags. Retained candidate receipts record 14 checks and 4,382
candidate-stage test executions; restored final source also passed 14 checks
and 4,382 executions. Cleanup and post-cleanup verification passed; the evidence is sealed. This closes one physical-reconciliation hypothesis
and does not complete the broader OLE2/OOXML optimization goal. ODF stays
deferred until the active goal completes, and iWork remains excluded.

## 0533: CFB candidate clears measured gates; full goal remains open

[0533](changes/0533-cfb-claim-cold-error-layout.md) advances the active
OLE2/OOXML goal with a source-bound CFB `claim_sector` code-layout candidate.
The four frozen primary XLS p50 gates pass in both repeats at
20.5382–24.5798% improvement, parent-constructor Ir falls in every profiled scenario,
and all allocation vectors remain identical. Four matched adverse rows and 30
same-build variations remain retained after individual review, including the
owned-list repeat-2 maximum outlier; no uniform tail claim follows. The 14
quality checks pass with 4,378 executed tests, and the independent adverse
review permits adoption. Post-cleanup verification and sealing are complete;
this batch still does not complete the full optimization goal.

The measured scope is warm synthetic in-memory XLS/CFB. It does not close
cold/range, physical-provider, native-producer, scaling, fuzz or broad CRUD
requirements. A fresh paired walk of residual physical reconciliation is the
next OLE2/OOXML hypothesis; rejected 0524 visited-bit fusion remains rejected.
ODF stays deferred until OLE2/OOXML optimization completes, and iWork remains
excluded.

## 0532: OLE2 attribution advances the goal; completion remains open

[0532](changes/0532-cfb-claim-success-path-attribution.md) completes a fresh
current-head CFB/OLE2 ownership and reconciliation attribution baseline. It
binds 24,000 native samples, 720 allocation samples, eight profiles with 40
timed and six setup dumps, and mechanism evidence for the per-sector
`claim_sector` success path. No runtime optimization or speedup is adopted.
The next bounded measurement is private cold error helpers plus ordinary
`claim_sector` inlining, preserving all checked bounds, error precedence,
fallible operations, collect-then-claim order and physical reconciliation.
Five fresh quality checks pass with 1,957 test executions; exact-source 0531
reuse remains separately bound at eight gates and 4,757 executions. The
OLE2/OOXML optimization goal remains incomplete, ODF stays deferred and iWork
is excluded; cleanup and evidence sealing are complete.

## 0531: final OOXML MCE candidate rejected; OLE2 remains next

[0531](changes/0531-ooxml-mce-namespace-search.md) makes progress by completing
the fresh final-binary native comparison, then rejecting the shared MCE
namespace-search candidate. The authoritative [final native comparison](results/change-0531/final-native-comparison.json)
fails the required dense-sparse repeat-2 total gates at 0.8937% p50 and
0.6056% mean reduction; the final conditional profile is mechanism evidence,
not a retained speedup. Production is restored, while nine MCE regression tests
and a test-only lint correction remain. All eight restored-source quality gates pass, including 4,757 successful
test executions. This batch does not complete the optimization goal.
The next OLE2 measurement is the bounded CFB ownership/reconciliation
attribution selected in [the OLE2 review](results/change-0531/next-ole2-review.md).
The overall goal remains open; ODF is deferred and iWork is excluded.

## 0530: progress through planning attribution

[0530](changes/0530-xlsx-planning-attribution.md) identifies a concrete namespace-search follow-up after the rejected 0529 probe. No runtime optimization is adopted. Broader OLE2/OOXML work remains incomplete; ODF stays deferred until that goal completes, with iWork excluded.

## 0529: progress through measured rejection and reusable instrumentation

[0529](changes/0529-xml-attribute-probe-pilot.md) rejects the zero/one-attribute probe after native gates fail, retaining publication counters, four public regression tests and all evidence. OLE2/OOXML remains the active priority; ODF is deferred until that optimization goal completes and iWork stays excluded. The full goal remains incomplete.

## 0528: progress through a new publication baseline

[0528](changes/0528-xlsx-publication-attribution.md) The batch verifies a small resolver-cost ceiling and captures four fresh publication profiles, identifying shared authored-XML validation as the next concrete lead. Production remains unchanged and the full goal remains incomplete. OLE2/OOXML stays first, ODF deferred until that optimization goal completes, and iWork excluded.

## 0527: progress through measured rejection

[0527](changes/0527-xlsx-row-primary-arena-pilot.md) The fresh XLSX pilot completed 2,440 native durations and 40 allocation samples. A failed dense repeat-2 gate caused restoration of production and retention of compatible tests/evidence. OLE2/OOXML performance remains the active priority; ODF is deferred until that optimization goal completes and iWork stays excluded. Overall completion is not claimed.

## 0526: next scanner experiment prepared; full goal remains open

[0526](changes/0526-xlsx-primary-span-layout-audit.md) The previous turn made progress through accepted commit `67028ab60`. This batch revalidates its source/evidence and prepares a concrete scanner-storage experiment with independent cost/design/test reviews. No production optimization is retained here. Fresh before/after measurements and correctness gates are the next action; broader CRUD, provider, cold/range and scaling requirements remain open. OLE2/OOXML continues ahead of deferred ODF; iWork is excluded.

## 0525: measured OOXML reconstruction progress; full optimization goal remains open

[0525](changes/0525-xlsx-unchanged-cell-readback.md) advances the active
OLE2/OOXML goal with source-bound omission of unchanged XLSX cell owners and
independent parsing of actual changed output. The matched primary matrix passes
the frozen native total/commit gates in all four shape/repeat pairs, and the
profile gate passes at 31.5439–32.4414% lower commit Ir. Allocation calls,
bytes and incremental region peak also fall in both shapes with exact repeat
vectors. The candidate is retained after supplemental eager confirmation passed both
pairs with zero >5% matched or drift flags. All 12 quality gates pass with
1,293 successful executions; cleanup and exact evidence replay pass. They do not close
physical FileSource/range, cold-cache, native Office-producer, fuzz, scaling or
broad CRUD requirements, and do not complete the full optimization goal. ODF
remains deferred until OLE2/OOXML work is complete; iWork is excluded.

The native review retains 31 over-5% rows (7 open, 19 publication, 4 reopen,
and 1 managed dense-sparse repeat-2 whole-child-RSS row at 85,648 → 90,244
KiB, +5.3661%) and 76 same-build drifts; no main commit or total-elapsed row
crosses 5%. Reopen diagnostics are outside measured edit/save elapsed time,
and the accepted decision weighs these phase/RSS tradeoffs alongside all
four primary total p50 gains. Quality has 12 passing gates and 1,293 successful
executions; the candidate is accepted and both owned scratch trees are absent.
The next audit targets scanner layout/address work (53.26% of these XLSX
commit profiles), with Store merge (5.97%) secondary; rejected mechanisms
remain excluded.

## 0524: visited-bit experiment closed; broader optimization goal remains open

[0524](changes/0524-cfb-visited-bit-evaluation.md) completes and rejects the
bounded CFB candidate under matched native admission. It retains an independent
collector differential and stabilizes an existing filesystem substitution test;
production is unchanged. The complete evidence includes adverse results and
both preflight/environment and later filesystem-test failures. This does not
close provider, producer, cold/range, fuzz, scaling or broad CRUD requirements.
OLE2/OOXML performance remains first until that goal completes; ODF is deferred
and iWork excluded.

## 0523: CFB allocation gap closed; OLE2/OOXML optimization continues

[0523](changes/0523-cfb-open-allocation-attribution.md) adds missing direct-CFB
allocation observations and a current OLE2 baseline: 24,000 native samples,
720 allocation samples and eight constructor profile children. Production is
unchanged; this is no speedup or coverage-catalog promotion. The next bounded
experiment is checked visited-bit lookup/set in chain collection, subject to
fresh matched admission. Provider, producer, cold/range, fuzz, scaling and broad
CRUD requirements remain open. OLE2/OOXML stays first; ODF is deferred until
that optimization goal completes, and iWork is excluded.

## 0522: complete candidate evaluation; OLE2/OOXML goal stays open

[0522](changes/0522-xlsx-cell-reference-guard.md) completes the narrow
common-cell scanner evaluation and rejects it under the frozen native gate.
Medium primary total p50 is mixed across repeats despite lower commit
instructions and allocations. The scanner is restored; stronger codec tests,
an explicit prefixed/multi-attribute harness guard and complete evidence are
retained. No coverage-catalog, provider, cold/range, producer, fuzz or scaling
gap is closed by this synthetic comparison. OLE2/OOXML remains the active
priority until its optimization goal completes. ODF, including the historical
ODG regression row below, is deferred; iWork remains excluded.

## 0521: XLSX ownership optimization accepted; broader goal remains open

[0521](changes/0521-xlsx-borrow-validation-events.md) closes the narrow
per-event ownership candidate and operation-local allocation attribution task.
Primary edit/save p50 improves 6.35–7.59%, commit allocation calls about 56%,
and scoped commit instructions about 10.4%. Incremental peak is unchanged.
All 23 adverse phase metrics remain explicit. The synthetic source-backed
matrix is not a coverage-catalog promotion or completion of physical providers,
cold/range, native producers, fuzz, broader CRUD or scaling requirements.
The next investigation follows current candidate reconstruction and worksheet
layout/rewrite owners, preserving validation and readback. OLE2/OOXML remains
the active priority until its optimization goal completes; ODF is deferred
and iWork excluded.

## 0520: current XLSX phase attribution complete; optimization remains open

[0520](changes/0520-xlsx-source-edit-phase-attribution.md) closes the immediate current phase/commit-owner attribution task for
the two synthetic source-backed one-percent edit/save shapes. It retains 400
native samples, four isolated commit profiles and executable output/preservation
oracles. It does not close operation-local allocation, physical FileSource/range,
cold-cache, native producer, broader CRUD or scaling requirements, and does not
promote the selector in the checked coverage catalog. The next work targets
measured candidate reconstruction and source-layout costs. OLE2/OOXML remains
active, ODF deferred until that optimization goal completes, and iWork excluded.

## 0519: reuse the immutable OPC publication XML proof

[0519](changes/0519-opc-publication-xml-proof-reuse.md) reuses complete XML
validation when destination and proof `ReadLimits` are exactly equal. The
publication owner retains Work, source/context, PartBytes, destination identity,
security, topology, transfer, and DOCX candidate readback checks. Different
limits retain full XML validation; matching limits no longer reserve unused
parser workspace.

Across 48 matched synthetic DOCX comparisons, lifecycle p50 improves
1.43–26.41% and publication p50 improves 36.31–80.85%. Scoped publication
instructions fall 47.33–47.37% for p128 and 77.47–77.53% for p512. The separate
allocator probe records 109→96 calls and 75,001→74,218 allocated bytes;
incremental region peak remains 72,535 bytes. Work, source I/O, exact output,
and release counters are unchanged. RSS has no >5% adverse flag.

All 107 phase flags remain explicit, including two second-campaign lifecycle
p99 regressions dominated by edit outliers. The separate eight-child tail
guard improves lifecycle p99 by 14.47–18.07% in all four comparisons, while
retaining three short commit-phase flags and all original flags. All nine
quality gates pass, including 5,012 executed tests. No general tail-latency,
cold/range, native Office producer, or scaling improvement is claimed. The
remaining local publication profile is mostly ZIP preservation and copying;
its instruction count alone does not prove further removable work. OLE2/OOXML
remains active, ODF deferred, and iWork excluded.

## 0518: reuse the retained DOCX source snapshot at publication

[0518](changes/0518-docx-source-snapshot-reuse.md) reuses the immutable
patch source after OPC checks the current Part, source identity, exact bytes,
limits, and security state. Candidate reparse/readback and destination
publication validation remain intact. Across 48 same-API comparisons,
lifecycle p50 improves 1.42–34.59% and publication p50 improves 43.75–67.42%.
Method instruction counts fall about 54–66%. In the separate instrumented
probe, publication allocation calls fall 96.65–99.13%, while incremental
region peaks improve only 2.70–6.55%.
All lifecycle/publication tails improve; RSS stays within the 5% review
threshold. The 69 adverse open/commit/drop flags remain explicit, with causes
unproven. These results cover the named synthetic DOCX matrix only.

The repeated current-snapshot owner is now negligible in these profiles;
topology publication dominates the remainder and is the next proof-review
candidate. Historical 0499/0500 flags, broader CRUD/corpus coverage, cold/range
and scaling requirements remain open. OLE2/OOXML work continues, ODF is
deferred until that goal is complete, and iWork is excluded.

## 0517: remove a duplicate shared OPC source-XML validation pass

[0517](changes/0517-opc-source-xml-validation.md) retains the initial complete
source/destination/XML proof and replaces the later duplicate scan with a
source freshness and cancellation check. All 48 same-route DOCX comparisons
improve lifecycle p50 by 1.87–14.09%; scoped publication instruction counts fall
about 17.9–21.1%. No adverse lifecycle, publication or whole-child RSS threshold
is triggered; short open/commit/drop flags remain explicit in the report.
This measures synthetic DOCX cases, not every OPC caller. The source-snapshot
construction path remains roughly half to two-thirds of publication CPU and
is the next measured investigation target. Historical 0499/0500 flags and
broader coverage/scaling requirements remain open. OLE2/OOXML remain the
priority; ODF is deferred until that optimization goal is complete.

## 0516: reject XLSX emitted-output parser fusion

[0516](changes/0516-xlsx-output-fusion-rejection.md) passes all 1,284 candidate
all-features tests but fails performance admission: changed-edit pilot p50
rises 4.46–8.30% across the 12 main rows and 5.82–8.64% across the six warm
changed-edit guard rows. Dense allocation calls also increase, and scoped
commit instruction references rise 3.95%. No formal ABBA campaign or speedup
claim follows. The exact candidate and all adverse evidence are archived;
production is restored. Two independent no-op fixture fixes remain, with
1,263 tests passing on unchanged production. OLE2/OOXML work remains active;
ODF is deferred and iWork excluded.

## Current priority after 0520: OLE2 and OOXML

Per the user's instruction, prioritize OLE2 and OOXML performance
until their full optimization goal is complete. Further ODF optimization is
deferred until then. This overrides the ordering of older entries below.
CFB profiling and FAT batching are complete in 0511. The 0514 source-pass
fusion and 0516 emitted-output fusion are rejected. The 0517–0519 DOCX publication
changes remove repeated XML and current-snapshot construction work. Further
publication investigation must distinguish required byte copying from
removable work and compare its end-to-end value with broader OLE2/OOXML
coverage. Required candidate validation and readback remain intact.
The [0520 attribution](changes/0520-xlsx-source-edit-phase-attribution.md)
now identifies source-backed XLSX candidate reconstruction and source-layout
scanning as the next commit owners to investigate. Operation-local allocation
and a distinct measured work-elimination design remain required; do not repeat
the rejected fusion proposals. Retain the 0499/0500 review flags. This investigation queue is not a completion checklist. The broader requirements and
outstanding coverage remain open; iWork remains outside this workstream.
See the [priority review](results/change-0510/ole2-ooxml-priority-review.md).

## 0515: isolate XLSX changed-output work

[0515](changes/0515-xlsx-output-attribution.md) separates changed-output parsing
from source Store parsing with caller-context profiles and captures compaction
alone. The unchanged-source commit profiles differ by 0.008% in simulated
instructions. Output parsing remains about 25.8% of commit work and compaction
about 20.4%, but sharing the reader directly targets only about 6.9% of commit
instructions. Semantic cell processing, materialization and namespace resolution
remain required; the full validation share is not a removable-work estimate.
The 720 normal samples show p50 repeat drift from −1.01% to +1.10%, with no
frozen latency or RSS review flag. Whole-child RSS is 138,480/135,732 KiB;
this is repeatability context, not an optimization or document-memory result.

The next implementation should investigate feeding emitted-equivalent events
from compaction into the existing parser only for effective changed outputs.
Reuse needs a proof for actual normalized bytes, exact-output x14ac/MCE/UTF-8
checks, deferred error ordering and an authoritative fallback. Style checks,
web validation, change readback, package reopen and the bounded Store handoff
remain intact. See the [semantic review](results/change-0515/output-semantics-review.md)
and [scope review](results/change-0515/scope-review.md). This evidence-only batch
makes no speedup or new phase-memory/hardware/cache/scaling claim. The full
OLE2/OOXML goal remains active; ODF is deferred and iWork excluded.

## 0514: reject speculative XLSX parser/layout fusion

[0514](changes/0514-xlsx-fusion-rejection.md) evaluates a shared-event source
parser and lossless rewrite layout. The candidate passes all 966 XLSX owner
unit tests, including 17 new differential/concurrency tests, but fails the
predeclared admission gate: cold same-value pilot p50 increases 64.43–84.64%
across all six size/update rows. Repeated dense no-op allocation observations
show incremental peak live demand rising 27.04% for one-cell edits and 18.64%
for one-percent edits. A 6.30% reduction in scoped save Callgrind instructions
does not justify that extra work. Changed-edit pilot results are mixed; no
full candidate native comparison or accepted speedup is claimed.

Production and candidate tests are restored exactly to the control revision.
The verified candidate patch, source-bound evidence and reusable public no-op
and cold-read guard remain available. Next work should investigate the larger
changed-output XLSX validation/compaction path or DOCX publication CPU cost;
0511's completed CFB profiling and FAT batching should not be repeated as an
unstarted item. See the [conditional follow-up](results/change-0514/follow-up-options.md).
The full OLE2/OOXML goal remains active, ODF stays deferred, and iWork remains
excluded.

## 0513: XLSX operation allocation baseline

[0513](changes/0513-xlsx-operation-allocation.md) adds allocation observations
to four existing XLSX commit/save cases and a private save profiling boundary.
The 4,800 matched native samples show p50 changes of −0.48% to +2.32%, with
no latency/throughput/RSS or repeat-drift flag. The separate 240 allocator
samples establish candidate-only baselines: dense one-percent commit allocates
273,128,176 bytes in 2,400,578 allocation calls; commit/save allocates
286,872,324 bytes in 2,532,326 calls. Both have 58,496,820 bytes of derived
incremental live demand above region entry. Absolute region peaks differ
because commit-only retains the prior result; neither is document peak or RSS.
This is a measured harness enabler, with no format speedup or memory-reduction
claim. The next production candidate must measure parser/snapshot reuse while
preserving error order, original-source spans and bounded temporary overlap.
The full OLE2/OOXML goal remains active; ODF stays deferred.

## 0512: current XLSX commit attribution

[0512](changes/0512-xlsx-commit-attribution.md) isolates three dense one-percent
commit bodies per profile, resetting after fixture generation. Two captures
differ by 0.012% in simulated instruction references. Direct source Store and
validation parsing consume about 52%, rewrite 27% and compaction 20%; nested
snapshot scanning is 25.40%, while shared-formula resolution is only 0.05%.
These are scoped instruction diagnostics, not phase clocks or a speedup.
Twelve unchanged native cases retain 720 durations across two repeats with no
same-build drift flags. Whole-child hardware/RSS remain separate; operation
allocation and exact save-phase boundaries are still missing. Next work should
add those observations before attempting parser/snapshot fusion with preserved
error order and bounded memory. No production or default-matrix change occurs.

## 0511: CFB FAT helper work reduced; broader goal remains open

[0511](changes/0511-cfb-fat-entry-reservation.md) removes repeated FAT-entry
helper work using the proven exact reservation. Eager XLS and CFB native
opening improve; plain OwnedSource results are mixed. Required chain,
sector-ownership and physical-layout checks retain their work and dominate
the remaining source-constructor profile. FileSource, cold/provider behavior,
operation-local hardware scope and broader scaling coverage remain open.
Continue the OLE2/OOXML investigation queue with dense XLSX commit/save and
DOCX provider/publication costs, revisiting CFB only with fresh attribution
and an exact safety proof. This batch does not complete the full optimization
goal or unlock deferred ODF work. Default coverage remains 41 cases, 213 rows,
43 corpora and 18 mapped correctness-only selectors; iWork remains excluded.

## 0509: profiled ODT export allocation reduction

[0509](changes/0509-odt-sink-buffer-reuse.md) follows fresh allocation evidence
with one bounded operation-local paragraph buffer. Stack allocation calls
fall 99.99% in the large synthetic export, with exact output/sink identities
and 1,491 passing Rust tests/doctests. Native captures retain 14,000 samples,
including a longer tail follow-up after an initial +9.14% large p99 flag.
The follow-up stays below 5%, but no tail-latency or peak-memory improvement
is claimed. This is scoped implementation progress; the remaining 18
correctness-only selectors, provider overhead, native-producer coverage,
historical comparisons and full non-iWork requirements remain open. The
program goal remains active.

## 0508: concrete default conversion coverage expansion

[0508](changes/0508-default-semantic-text-export.md) preserves all 201 prior
default identities and adds twelve measured text-export rows. The default
now has 41 cases/213 rows/43 corpora. Two fresh full matrices validate 6,390
samples; 15 mapped measured selectors bind 60 case/corpus rows, with 18
selectors still correctness-only. The XML-compaction checklist mismatch is
corrected. This advances representative conversion/export coverage without
claiming native-producer, cold-provider or complete checklist certification.
The [ODG audit](results/change-0508/odg-priority-review.md) distinguishes
matched 0504–0507 improvements from the historical old-parser comparison.
Remaining priorities include the 0499 local Part-batch overhead, 0500 K1
latency/RSS flags, the remaining coverage gaps and real-producer evidence.
The program goal remains active.

## Current audit: 0507 reduces repeated ODG value work

[0507](changes/0507-odg-attribute-value-batches.md) records this batch.
The next attributed bottleneck is reduced by two fixed-size request groups,
retaining scalar error precedence, semantic validation order and exact source
preservation. All 124 tests pass; the capture improves median time 17–34%
and plain-large RSS about 9%. This is scoped progress with no >5% adverse
paired flag, not closure of all older-parser comparisons or the broader goal.
The ten-entry strict registry and outstanding CRUD/provider requirements
remain unchanged.

## Current audit: 0506 batches ODG source-span discovery

[0506](changes/0506-odg-shape-attribute-span-batch.md) records this batch.
The measured source-span bottleneck is reduced while exact field ordering,
namespace resolution, checked attributes and fallback errors remain covered.
The 108-test suite includes all 16 field mappings and private error parity.
The plain-large ~10% RSS regression is accepted for the measured latency/work
reduction, with unresolved allocator/working-set attribution explicitly retained.
This does not complete the broader goal, close older-parser comparisons or
change the ten-entry strict registry.

## Current audit: 0505 reduces ODG attribute lookup work

[0505](changes/0505-odg-attribute-name-prefilter.md) records this batch.
A local-name filter retains checked iteration and exact source preservation,
while reducing p50 8.43–9.88% in the matched richer-parser capture. Four new
tests cover namespace selection and trailing invalid attributes. Full CRUD,
provider/scaling coverage and older-parser regression attribution remain open.
The strict claim registry remains at ten entries; this is scoped progress.

## Current audit: 0504 reduces a measured ODG open bottleneck

[0504](changes/0504-odg-direct-transition-reuse.md) implements local reuse of
successfully parsed direct page-transition values. The 3,200-sample matched
richer-parser comparison reduces metadata-large open/traversal p50 by about
34% and metadata-small by about 10%, with no paired >5% adverse latency or
whole-child RSS flag. Per-page inheritance validation remains in force.
This is production and scoped measurement progress, not completion of the
ODG regression review: `parse_content` remains the dominant profiled subtree,
unique-style-heavy input is unmeasured, and the older 0502 baseline exposed
less metadata. Full CRUD, provider, memory-attribution and scaling requirements
remain open; the ten-entry strict registry is unchanged.

## Current audit: 0502 adds bounded ODG metadata but retains large open regressions; the full goal remains open

[0502](changes/0502-odg-metadata-open.md) adds typed ODG transitions and
auxiliary/enhanced shape metadata, with presence-gated validation, parsed-style
reuse, and boxed cold metadata. The committed probe evidence is four
deterministic corpora, two serial repeats, 25 warmups, and 200 samples per
child: 1,600 samples per phase and 3,200 matched samples in
[the retained summary](results/change-0502/summary.json). The timing comparison
is clean HEAD `f3f9221` versus the final working tree; source manifests and all
raw reports are retained under `results/change-0502/`. The summary was
recomputed from those reports and matches the retained JSON.

The final p50 changes are +5.903/+6.386% for plain-small, +6.062/+5.949% for
plain-large, +53.402/+55.215% for metadata-small, and +110.090/+111.136% for
metadata-large across repeats. The p95/p99 rows show the same direction,
including +110% metadata-large tails. Boxing reduces `Shape` from the 936-byte
pre-boxing candidate to 800 bytes, versus 792 bytes at clean HEAD. A separate
whole-child heaptrack comparison reports peak heap 13.69M to 12.57M with the
same 174,853,447 allocation calls; its instrumented p50 and RSS are not
operation-local measurements. Hardware counters were unavailable, so no CPU,
cache, branch, or IPC conclusion follows. The ODG work is therefore a bounded
correctness/layout enabler with an explicit regression queue, not an accepted
end-to-end performance improvement.

| Goal area | 0502 evidence | Audit status and boundary |
| --- | --- | --- |
| ODG open and metadata traversal | 3,200 matched samples over plain and metadata corpora | Measured regression; follow-up parser/representation work is required |
| Layout and allocation shape | `Shape` 792 → 800 bytes; pre-boxing 936; heaptrack peak display −8.2% with unchanged call count | Whole-child/layout evidence only; no allocation or RSS claim for the timed operation |
| Native producer and semantic correctness | Retained ODG tests include LibreOffice fixture checks and bounded/refusal cases | Correctness evidence only; the probe uses generated packages and does not establish native-producer performance |
| Full non-iWork goal | No CRUD selector, edit/save, cold/range, concurrent, or scaling timing is added | Open |

## Current audit: 0501 replay evidence remains retained, but its current verifier cannot be replayed

The matched [0501](changes/0501-pptx-exact-payload-comparisons.md) reports,
catalog checks, and 38 static coverage tests remain present and pass their
current validators. The committed `verification-final.json` is the historical
formal custody receipt for the eight before and eight after reports. A fresh
run of `verify-final.py` is currently unavailable because the frozen before
executable under `/tmp/litchi-goal-0501` is missing; the replay reports
`before: frozen binary is missing`. This is a replay/custody limitation, not a
reason to discard the retained report rows or to describe the historical receipt
as a new run. The strict claim-registry check currently validates ten claims.

The formal comparison and its 52 favorable flags remain scoped to the named
PPTX source-backed media/plain workflows; native producer closure, cold and
physical-I/O behavior, and exhaustive CRUD promotion remain open. The prior
0501 timing, profile, cleanup, and exact-payload boundaries below continue to
apply.

## Current audit: 0501 scopes PPTX payload hashing; the full goal remains open

[0501](changes/0501-pptx-exact-payload-comparisons.md) removes redundant
private image/chart payload hashing while retaining exact payload comparisons,
candidate reread, graph and source proofs, cancellation, budgets, and partial
publication behavior. The fresh before phase has eight reports and 240 measured
samples, and the matched after phase has the same eight reports and 240
measured samples. The supplementary whole-child
SHA-256 profile reports 31.01% owned and 30.08% warm-file, with incomplete
caller recovery, so it does not establish touched-digest attribution or a
timer-local causal speedup. Same-data export recovery and two subsequent owned
profiling workload reruns are recorded separately and do not alter the formal
matrix. The comparison retains all 208 rows, passes every lifecycle oracle, and
retains 52 favorable timing/throughput flags with no adverse change above five
percent. The separate default CRUD refresh
passes two serial 201-row lanes, totaling 6,030 samples across 37 cases and 31
corpora; both generated report/catalog validators and 38 static coverage tests
pass. It closes the timing-report baseline gate only. The scoped production
gates also pass: default all-targets 848, all-features library 552, doctests 6
with 2 ignored, focused 58, private guards 5, Clippy, fmt, rustdoc, downstream,
boundaries, and 38 CRUD static tests. Independent strict verification passes.
Final cleanup removed 2,154,708,992 unique-inode allocated bytes and
retained eight replay files in local tmpfs: two binaries and six raw perf files,
including two failed attempts. Historical 15.5–16.1%
touched-digest attribution is from 0449. Native producer notes/chart closure,
cold and physical-I/O behavior, exhaustive CRUD promotion, and other goal
areas remain open.

## Current audit: 0500 targets managed paragraph batching; the full goal remains open

[0500](changes/0500-managed-paragraph-batches.md) targets the managed refusal
at the existing `Edit::replace_body_paragraph_texts` seam. The repeated
baseline has 12 children, 720 measured samples, and 72 warmups; at 32 selected
paragraphs, edit accounts for about 93% of the timed lifecycle. The candidate
builds one source-checked final projection and preserves finite ownership,
atomic failures, inverse proofs, and monotonic accounting. The 24-child after
phase completes a 36-child, 2,160-measurement comparison. Against repeated
scalar in the same final executable, K=32 lifecycle p50 improves 6.947–7.138x
and edit p50 12.640–13.456x; K=8 lifecycle improves 2.324–2.346x. K=1 p128
owned retains a 7.34% lifecycle flag, and p512 K=8 warm-file retains a 5.77%
RSS flag. These are scoped API-choice results; CRUD-index promotion, native
producers, controlled-cold behavior, one-percent coverage, and other goal
areas remain open.

## Current audit: 0499 targets bounded Part worker lifetime; the full goal remains open

[0499](changes/0499-operation-local-part-workers.md) reuses a bounded,
operation-local scoped worker set when an ordered Part batch spans waves, with
the existing single-wave path retained. The implementation preserves source
and cancellation fences, worker/task/byte limits, owner-retained results,
monotonic `Work` and `InputBytes`, typed error ordering, and complete joins.
An unwind guard also closes a provider-panic cache flight; recovery is limited
to unwind builds because release uses `panic = abort`. The matched capture has
60 children, 3,600 measured samples, and 360 warmups. Many-small owned p50
improves 3.59x/2.92x/2.08x and warm-file 2.64x/2.26x/1.81x at widths 2/4/8,
but both remain slower than their ordinary serial controls. Five aggregate and
ten per-repeat flags remain retained, so this is scoped workload evidence
rather than a program-wide speed result. CRUD, native-producer, controlled-cold,
history/composition, and other goal areas remain open.

## Current audit: 0498 adds explicit bounded Part batches; the full goal remains open

[0498](changes/0498-bounded-source-backed-part-batch.md) adds an owning ordered
OPC Part batch with caller-controlled worker, task, byte, cancellation, and
memory limits. The 30-child same-executable matrix has 1,800 measured samples
and 180 warmups with byte/budget/cleanup checks. Delayed-source scaling is
useful, while many-small owned/file reads regress materially. Apparent
superlinear few-large observations remain unexplained; no scheduler-only or
program-wide speedup claim follows. Broad CRUD, native producers, controlled
cold behavior, allocation attribution, and history requirements remain open.

## Current audit: 0497 adds an atomic DOCX publication capability; formal analysis is verified

[0497](changes/0497-docx-atomic-publication.md) adds consuming
`ParagraphStreamPlan::write_to_path` and
`ParagraphStreamCommit::write_to_path` methods for the existing bounded DOCX
logical-tail append route. The methods use OPC's sibling-temporary atomic
replacement helper and retain the existing source, replay, budget,
cancellation, candidate, and inverse-proof boundaries. A callback failure before
replacement preserves the destination and removes the private temporary;
post-replacement parent-directory sync failure returns typed
`OpcError::Committed` and must not be blindly retried.

The default hashing route remains the only before/after comparison and keeps
its existing report shape. Counting is a bounded non-retaining after-only
capability; atomic publication is also after-only because the before revision
has no atomic destination method. The frozen formal inventory contains 288
children and 8,640 samples, with a separate 72-child/216-sample pilot. Formal
verification passes all 288 child terminals, including the 97-child resume
after the original `ENOSPC` interruption. The descriptive analysis retains 864
matched default-hashing comparison cells and 144 after-only capability rows.
It retains all 100 >5% adverse flags: 72 allocator live-byte endpoint, 16
latency, and 12 whole-child RSS. The original ordinal-191 raw observation
remains archived without fabricated evidence. Cleanup overlapped two formal
whole-child intervals, so no isolated-host or individual-outlier attribution is
made. Final cleanup verification passes; the evidence inventory is recorded
by `results/change-0497/seal.py`.

| Goal area | 0497 candidate evidence | Audit status and boundary |
| --- | --- | --- |
| Atomic filesystem publication | Consuming DOCX plan/commit path methods delegate to the existing OPC atomic sibling replacement | Capability scope only; no general atomic-save or durability claim |
| Default publication comparison | Historical hashing sink remains unchanged and is the only before/after route | 864 descriptive comparison cells; no broad speedup or causal claim |
| Counting and atomic routes | Non-retaining counting and filesystem atomic routes are balanced after-only capabilities | 144 descriptive capability rows; no atomic before/after comparison or synthetic write-call/digest values |
| Broader non-iWork goal | No new representative CRUD selector; borrowed input, producers, cold behavior, scaling, history, and broad CRUD remain open | Open |

## Current audit: 0496 verifies descriptive phase attribution; the full goal remains open

[0496](changes/0496-docx-edit-phase-attribution.md) is a harness-only,
opt-in diagnostic follow-up to the unresolved 0495 whole-child RSS and latency
flags. Its verified formal run has 32 children and 960 samples: 12
before-unmanaged, 12 after-unmanaged, and 8 after-managed, with two reversed
repeats, three warmups, and 30 measured samples per child. The default report
schema and existing correctness oracles remain unchanged. The claims scope is
descriptive phase latency, whole-child RSS, and full-lifecycle allocation;
managed-after rows remain capability observations and no CPU, optimization, or
causal claim is authorized.

The paired unmanaged review contains 294 comparison cells, with 74 threshold
flags retained and classified by evidence scope.

Normal unmanaged lifecycle p50 latency is milliseconds, paired by repeat:

| Arm | Before R1 | After R1 | Before R2 | After R2 |
| --- | ---: | ---: | ---: | ---: |
| owned | 2.221 | 2.153 | 2.126 | 2.176 |
| file-warm | 4.794 | 2.302 | 2.279 | 2.326 |
| short | 5.359 | 5.340 | 5.238 | 5.419 |

Publication is the largest named normal phase in all 16 normal children. The
after-managed file-warm lifecycle/publication p50s are 3.563/2.465 ms and
3.553/2.453 ms; short is 6.509/5.406 ms and 4.092/2.993 ms, showing repeat
instability. Phase clocks are wall time, not CPU time; allocator and RSS remain
full-lifecycle/separate scopes.

The 74 retained threshold flags are 60 phase rows (18 commit-drop, 18
published-snapshot-drop, 10 diagnostics/XML identity, 7 open, 6 publication,
1 phase-sum), 9 allocator reallocation rows, 4 whole-child RSS rows, and 1
full-lifecycle latency row. The allocator p50s are 22,859 calls/5,721,334
bytes/+606,959 peak increment before-unmanaged, 9,696/1,623,696/+609,903
after-unmanaged, and 20,535/4,270,196/+622,454 after-managed. These rows are
descriptive flags, not independent regressions; `historical_flags_resolved` is
false and causality remains unresolved.

Formal verification passes all 32 children and 960 samples. The final helper
gate passes 8/8 seal-helper tests; the evidence seal remains a separate custody
gate. Cleanup passes source-manifest and protected-file checks after removing
10.709 GiB of disposable custody data and retaining four replay binaries. The
full non-iWork goal remains open: borrowed input, atomic save, independent
producers, cold intersections, scaling, durable history/composition, and
broader CRUD/security evidence still require separate work.

| Goal area | 0496 evidence | Audit status and boundary |
| --- | --- | --- |
| Opened-document edit/save | Verified descriptive phase/lifecycle capture over the 0495 path | No optimization or causal claim |
| CPU/RSS attribution | Wall-clock phase vectors and whole-child RSS | Wall time is not CPU attribution; RSS is not phase-local |
| Allocation attribution | Full-lifecycle calls, bytes, reallocations, and peak increments | No nested phase allocation region; managed-after remains descriptive |
| Existing 0495 flags | 0496 retains 74 flagged comparison rows alongside unresolved historical flags | No flag is dismissed or deleted |
| Complete non-iWork goal | No new CRUD selector or production capability | Open |

## Current audit: 0495 enables and measures ordinary managed DOCX edit/save; the full goal remains open

[0495](changes/0495-docx-managed-document-edits.md) closes the finite-owner
and source-authority seam that 0494 identified as a prerequisite for an
ordinary managed opened-document edit. Its verified formal run retains 72
processes and 2,160 samples; pilot2 retains 36 processes and 108 samples. The
matrix covers six providers (`owned`, `instrumented`, `file-warm`, 4 KiB
`short`, delayed 64 KiB `delayed`, and zero-fixed-delay `range-zero`) in normal
and allocator roles over a 200-paragraph, 20-member DOCX with eight 2 MiB media
members. Every formal row emits the expected 16,793,048-byte artifact and
passes source, sink, semantic, and untouched-media checks. Managed rows also
pass their finite-budget checks, while allocator-role rows pass
allocator-conservation checks; normal and unmanaged rows do not expose those
allocator fields.

The normal managed p50 observations in milliseconds are 5.592/5.556 (owned),
3.277/3.298 (instrumented), 5.682/5.853 (file-warm), 4.066/4.060 (short),
572.748/576.147 (delayed), and 182.327/181.990 (range-zero) for repeats one
and two. These are managed capability and budget observations. The protocol has
no managed-before baseline, so they do not authorize a managed speed, RSS,
allocation, or provider-ranking claim. The only performance comparison is the
paired unmanaged before/after review, which retains eight whole-child RSS flags
and three latency flags above the five-percent threshold. Their causality is
unresolved, and repeat instability remains visible rather than being dismissed
as host noise.

All 720 managed formal rows report zero reservation failures and release
resource memory, objects, and depth to baseline. The lifecycle labels are
separate: `cache_before` is post-open, `cache_live` is post-edit/pre-publication,
resource `live` is post-publication, and `after_drop` follows package
consumption plus returned-snapshot/commit release. No post-publication cache
gauge exists. The profile review is whole-child evidence that includes setup,
output verification, and serialization; it does not attribute CPU or RSS to the
timed operation. The next measured follow-up is phase-local CPU and whole-child
RSS attribution, with repeats of the unstable file-warm and short/allocator
arms, plus a distinct post-publication cache observation if retention is being
claimed.

| Goal area | Current evidence | Audit status and boundary |
| --- | --- | --- |
| Opened-document edit/save | 0495 verified managed capability/budget baseline plus paired unmanaged before/after review | Descriptive evidence only; no broad speedup or managed-before comparison; atomic filesystem save remains open |
| Managed ownership and source authority | Owner-retained snapshots/views, finite admission, source-checked publication and complete-artifact inverse tests | Necessary ordinary-edit enabler; durable history, composition, broad mutators, and dependency-bearing edits remain open |
| Provider and cache boundaries | Six explicit provider arms, two roles, two repeats, and distinct cache/resource lifecycle gauges | Cache diagnostic is pre-publication; release gauge is post-publication/drop; no cache-after-drop or provider ranking claim |
| CPU/RSS attribution | Core counters, syscall traces, and owned stack captures in the profile bundle | Whole-child diagnostics only; phase-local operation attribution remains the follow-up |
| Borrowed and independent-producer coverage | No genuine borrowed source or native producer round trip | Open |
| Concurrency and scaling | One worker and serial lifecycle execution | Open; bounded 1/2/4/8-worker evidence and Amdahl analysis remain required |
| Complete non-iWork CRUD checklist | 0495 adds no representative selector or default row | Open; structural/deletion, cross-document, merge/split, patch, repair, dynamic-content, security, and broader format rows remain |

The final validation bundle records 478 harness tests (one ignored), 937 DOCX
unit tests, 119 DOCX integration tests, 388 OPC unit tests, 79 doctests (31
ignored), and 58 Python helper tests. These are scoped correctness, custody,
and measurement gates; they do not turn 0495 into a broad performance claim.
The full non-iWork goal remains open, and iWork is untouched.

## Current audit: 0494 adds opened-edit provider baselines; the full goal remains open

0494 supplies the missing descriptive baseline for one opened DOCX paragraph
replacement, commit, and sequential publication across six explicit warm
provider arms and a verified-cold `FileSource` lane. The warm capture retains
24 formal processes and 720 samples (plus 36 pilot samples); the cold capture
retains 120 formal samples and six pilot samples. The deterministic corpus has
200 paragraphs, 20 archive members, and eight 2 MiB media members. Warm and
cold rows use the same logical content; the cold lane uses a page-aligned copy
with a padded ZIP tail and therefore a distinct physical archive hash. Output,
semantic edit, untouched-part/media, and source-version checks pass for the
retained rows. Patch replay, inverse, and stale-source checks are untimed
preflight gates.

The result is baseline evidence with `claim_authorized: false` and no
before/after optimization claim. Warm normal p50 ranges from 2.298 ms for the
first instrumented repeat to 569.581 ms for the delayed-range arm; the two
repeat values vary substantially for several providers. The verified-cold
normal p50 is 148.065 / 238.430 ms and allocator p50 is 98.139 / 93.802 ms.
The cold lane has its own source-open and residency/I/O boundary, so these
values are not a warm-versus-cold performance comparison. The 4 KiB short-read
rows produce 4,217 source calls and 3,840 short reads; other traced rows
produce 377 source calls. Warm allocator rows retain a 606,986-byte peak
increment and 22,859 calls; cold allocator rows retain 607,702 bytes and
22,864 calls. The independent 63-row cold allocator audit passes.

| Goal area | Current evidence | Audit status and boundary |
| --- | --- | --- |
| Opened-document edit/save | One existing paragraph replacement through commit and sequential publication across six warm providers and verified-cold file input | Baseline coverage is present; 0495 subsequently enables the managed path, while this 0494 capture remains descriptive and general atomic filesystem save remains open |
| Read-only managed provider lifecycle | 0493 synthetic opt-in open/load/text evidence | Accepted as a separate read-only slice; it must not be conflated with 0494 edit/save |
| Provider and cache boundaries | Owned, instrumented, warm file, short-read, delayed-range, range-control, and verified-cold file cells | Descriptive provider baselines; no provider ranking or warm/cold delta is authorized |
| Borrowed and independent-producer coverage | No genuine borrowed source and no native producer round trip | Open |
| Concurrency and scaling | 0493/0494 lifecycle captures are serial; 0498/0499 add bounded low-level Part batches and 0500 adds managed batch comparison | Open; full lifecycle 1/2/4/8 scaling, lock-wait evidence, and Amdahl analysis remain required |
| Complete non-iWork CRUD checklist | The 0494 path adds no selector or representative-index row | Open; broader format and CRUD rows remain to be evidenced |

The ordinary managed edit boundary was intentional at the 0494 revision. The
0495 follow-up now provides the owner-retained managed capability; 0494 remains
its descriptive provider/cold baseline and does not itself establish a managed
before/after speed claim. iWork remains outside this audit while its separate
work is in progress.

## Current audit: 0493 production read-ahead is accepted; the full goal remains open

0493 completes the production OPC integration of the bounded forward-start
read-ahead policy selected by 0492. The policy is explicit and opt-in, the
default source-backed constructors remain exact, and publication or exact
preservation paths permanently close forward admission before their first
exact archive read. The DOCX facade forwards the policy without exposing
archive implementation types. The source and focused integration tests cover
budget charging, retained-window release, source-version changes, re-entry,
panic/poison recovery, queued cancellation, short reads, and untouched-member
preservation.

The accepted 0493 evidence is a policy comparison using the same retained
executable within each normal or allocator role for a fresh managed DOCX open,
main-document load, text extraction, and drop over a pinned in-memory provider.
It contains 24 pilot and 480 formal samples. The modeled
delayed provider reduces managed p50 by 81.82–84.87%; physical calls fall
19→3 while accepted bytes rise 3,966→5,445 (+37.29%). Zero-delay changes stay
within −2.01% to +2.28%, allocator cost is +3 calls and +4,384 bytes, and no
same-repeat latency or whole-child RSS regression exceeds five percent. These
figures establish a bounded provider/read-only result, not a default-local,
real-network, cold-filesystem, or broad format claim.

| Goal area | Current evidence | Audit status and boundary |
| --- | --- | --- |
| Explicit caller-supplied range reads | OPC-owned bounded managed window, physical-fill accounting, exact fallback, and DOCX forwarding | Partial completion for the opt-in source-backed path; default behavior remains exact |
| Read-only content extraction | Fresh synthetic managed DOCX open/load/text lifecycle with strict source, text, cache, physical-trace, and budget oracles | Accepted but narrowly scoped to the pinned in-memory provider and one worker |
| Opened-document edit/save | 0494/0495 add descriptive provider/cold coverage and owner-retained managed edit/save; 0497 adds an atomic logical-tail capability | Partial capability and baseline evidence; broad before/after coverage and general atomic save remain open |
| Provider, cold, and producer breadth | 0491 and 0494 add explicit provider and verified-cold baselines | Borrowed lifetimes, native producers, and cold intersections across the CRUD matrix remain open |
| Bounded concurrency and scaling | 0498/0499 provide explicit ordered Part batches; 0500 compares managed batch versus repeated scalar edits | Partial low-level evidence; full lifecycle 1/2/4/8 curves, lock wait, and Amdahl analysis remain open |
| Complete non-iWork CRUD checklist | No 0493 selector or representative-index promotion | Open; conversion, structural/deletion, cross-document, merge/split, patch, repair, dynamic-content, security, malformed, and broader format rows remain to be evidenced |

The source validation for this batch passed 628 OPC tests, 1,390 DOCX tests,
456 harness tests, and 27 selected Python helper tests, alongside warning-
denied lint/documentation and boundary gates. Those are scoped correctness and
integration gates; they do not turn the read-only synthetic slice into
opened-edit, native-producer, cold-cache, borrowed-lifetime, scaling, or full
CRUD evidence. iWork remains outside this audit while its separate work is in
progress.

## Current audit: 0492 measures a range read-ahead enabler; 0493 is the bounded production integration

[0492](changes/0492-docx-bounded-range-read-ahead.md) is a benchmark-only
4 KiB forward-window experiment over 24 pilot and 480 formal samples. On its
simulated 1 ms range source, p50 falls from about 20.34 ms to 3.46 ms while
physical calls fall 19→3 and accepted bytes rise 3,966→5,445 (+37.29%). The
zero-delay normal p50 changes +4.88%/+0.50%, with a repeat-one p99 increase of
5.20%; the window allocation is outside the operation allocator region. The
private mutex and one-worker route are correctness details, not scaling
evidence. [0493](changes/0493-managed-opc-source-read-ahead.md) carries the
bounded policy into OPC with explicit InputBytes and Memory accounting; neither
batch establishes real-network, cold-filesystem, borrowed-lifetime, or broad
CRUD performance.

## Current audit: 0491 establishes DOCX provider and verified-cold baselines

[0491](changes/0491-docx-provider-and-cold-baseline.md) adds descriptive
source-provider and filesystem observations: 600 provider samples, 360 fresh
child filesystem samples, and four prepared-query controls that are explicitly
cold-ineligible. The simulated 1 ms range arm has roughly 20.3 ms p50 and 19
calls; local provider medians are roughly 0.27–0.30 ms. Verified-cold p50 is
2.753/4.361 ms normal across repeats, with 2.724/2.797 ms allocator, and each
row proves one main-part materialization plus positive process I/O. Repeat
tails vary materially. Whole-child profiles include setup and report work, so
they do not attribute operation CPU or RSS. This is a baseline, not a speedup;
0492/0493 provide the subsequent bounded read-ahead experiment and integration.

## Current audit: 0490 leaves file-store tails unresolved and bounds sync conclusions

[0490](changes/0490-file-store-variance-and-sync-attribution.md) alternates
retained before/after binaries through six blocks, preserving 72 processes and
4,320 samples with exact source/output identities. Normal file-store p50 has a
mean 6.73% reduction with a block-bootstrap interval of −9.71% to −4.41%, but
normal and allocator p95/p99 intervals span both directions and adverse blocks
remain. Four sync-only diagnostics attribute 78.99–81.61% of traced median
operation time to the unchanged `fdatasync`; selected whole-child syscall
counts match. The derived 1.23–1.27× bound is a scoped Amdahl model, not a
durability optimization or hardware limit. No production change follows;
0491 onward supplies the provider/cold work that this follow-up identified.

## Current audit: 0489 reuses candidate XML audit work with retained adverse rows

[0489](changes/0489-opc-candidate-xml-audit-reuse.md) reuses one successful
candidate XML audit inside an immutable prepared splice plan while retaining
source/candidate byte and EOF authentication, freshness, cancellation, Work,
limits, and final reopen. Its 144-process, 4,320-sample comparison reports
roughly 20–21% lower source-heavy medians, 14–15% lower authored-heavy file
medians, and about 19.89% lower deterministic operation heap. Candidate bytes
are unchanged. A file-store tail regression and four whole-child RSS increases
above 5% remain in the review; the result does not authorize dropping freshness
or independent candidate proofs.

## Current audit: 0488 records evidence-preserving cleanup only

[0488 cleanup](results/change-0488-disk-cleanup/README.md) changed no
production code or performance behavior. It removed 503,605,403,648 allocated
bytes of rebuildable Cargo output and 12.25 GiB of inactive worktree output,
481.26 GiB in total, after process-use checks; protected and dirty worktrees
were excluded and reachable commits were retained. The post-cleanup 0487 seal
still verifies 144 formal processes, 4,320 samples, 12 diagnostic children,
and two fuzz campaigns. This is custody evidence; rebuilding is required
before new workspace tests, and the cleanup does not strengthen any performance
claim.

## Current audit: 0487 reduces replay sink fences but retains small-workload regressions

[0487](changes/0487-opc-replay-consumed-prefix-retention.md) keeps consumed
replay bytes in the existing adapter allocation across short reads. The matched
144-process/4,320-sample comparison lowers authored-heavy file p50 by
35.61–35.99% and leaves source-heavy medians within 1%; diagnostic `statx`
falls 59.90% on authored-heavy file input while `pread64` remains unchanged.
Four latency rows and three whole-child RSS pairs above 5% remain, all in the
small-workload controls, and operation-heap rows do not increase. The parallel
allocator harness assertion failed because its process-global counter is not
isolated; the serial retry passes and the failure remains retained. Exact
source/output and preservation oracles pass. The broader goal and candidate
audit reuse remained open at this change.

## Current audit: 0486 attributes DOCX replay metadata cost without a production change

[0486](changes/0486-docx-replay-metadata-callers.md) retains four tiny
whole-child caller profiles over the 0485 executable. The authored-heavy file
profile places about 27.38% of sampled period weight in stacks containing
`statx`; this is inclusive sampled attribution, not a syscall count, CPU
measurement of the operation, or speedup. Its review identified the consumed
prefix retention candidate and an unresolved partial-output contract boundary,
which 0487 implemented under additional tests. Source-heavy 0485 strace data
showed 25,219 `statx` calls, but the 0486 sampled absence of a `statx` frame in
other profiles does not imply zero calls. No production code changed and the
full provider, cold, scaling, native-producer, and CRUD requirements remain
open.

## 0485: bounded consumed-window batching; full goal remains open

[0485](changes/0485-opc-splice-consumed-window-batching.md) applies a private
OPC adapter optimization to the replayable DOCX logical-append route. It
batches hashing and sink output for bytes already consumed inside the existing
bounded window while retaining the separate source XML audit, replay
authentication, source freshness policy, per-fragment Work charges, and
source-bound preservation checks.

The matched evidence contains 144 formal processes and 4,320 samples. At 64
existing and 16,384 authored paragraphs, file-input p50 falls from
443.916/440.383 ms to 239.499/238.865 ms across the two repeats. At 131,072
existing and 64 authored paragraphs, owned-input p50 falls from
482.325/476.215 ms to 385.058/385.006 ms. Operation heap peaks are effectively
unchanged. Three latency quantiles and nine small-workload whole-child RSS
observations exceed the five-percent review threshold and remain part of the
accepted scope.

The profile bundle records authored-heavy file `statx` falling from 3,735,939
to 1,475,055 while `pread64` remains 114, and source-heavy file `statx` rising
from 15,415 to 25,219 while `pread64` remains 264. These are one-sample,
one-warmup whole-child diagnostics that include setup and oracle work. They
justify caller-level attribution before any further change; they do not justify
relaxing freshness checks or deleting the independent source proof.

The route passes its scoped all-feature/no-default tests, formatting, warning-
denied Clippy and rustdoc, boundary, helper, benchmark, and sanitizer gates.
It does not close cold-cache, concurrent, atomic-save, all provider/input
intersections, or broad native Office validation. It also does not add a
representative CRUD selector. The machine-readable index therefore remains 15
categories and 34 rows: 11 measured, 22 correctness-only, and one explicitly
unsupported dynamic-content row (33 selector-backed mappings).

## 0482: shared primitives complete; end-to-end append remains open

[0482](changes/0482-bounded-xml-opc-splice.md) implements bounded XML auditing and decoded OPC insertion publication.
Focused tests cover source/candidate proof refusal, opaque ZIP preservation,
no-ops, inverse restoration, partial output, cancellation and budgets. Primitive
measurements cannot prove the program's bounded append requirement: public
DOCX append still uses the materialized lifecycle, and generated fragments,
package metadata and caller storage have distinct ownership. The
[next work](results/change-0482/next-work.md) retains DOCX scanner integration,
replayable paragraph generation, durable patches, provider variants and
end-to-end scaling evidence. Broader CRUD and parallel scaling obligations
remain in scope. The full goal is not achieved.

## 0481: scanner allocation progress; full goal still open

[0481](changes/0481-docx-borrowed-scanner-names.md) removes measured unnecessary
name allocation work and records lower normal lifecycle means in both repeats.
It leaves peak operation heap and document-sized XML/index ownership unchanged.
The [explicit-window contract](results/change-0481/window-contract.md) separates
the first decoded-splice milestone from the required multi-paragraph producer,
replay, durable inverse and large-stream scaling work. The broader non-iWork
requirements remain open; this batch does not redefine the definition of done.

## DOCX publication duplicate payload removed; append bound open (0480)

[0480](changes/0480-docx-shared-publication.md) removes one complete target XML copy at the existing OPC handoff.
The large operation peak is 35,371,102 bytes instead of 41,793,870 bytes;
retention still grows with document size. This is measured ownership progress,
not a bounded-window append completion claim. Repeated scanner/layout work,
the separate tail-only capability and the broader non-iWork scenario/source/
native/parallel requirements remain open.

## DOCX logical append baseline captured; window requirement open (0479)

[0479](changes/0479-docx-tail-append-baseline.md) records 720 samples of the existing one-paragraph tail-copy lifecycle.
The 41,793,870-byte incremental heap peak at 131,072 source paragraphs confirms
that this materialized transaction is not a bounded-window append solution.
Its current one-operation and section-property restrictions also leave the
proposed 64/256 repeated-append workloads uncovered. Production changes follow
measured ownership/CPU attribution. Broader CRUD, native/source variants,
parallel scaling and the full non-iWork objective remain open.

## Explicit public PPTX metadata window implemented (0478)

[0478](changes/0478-pptx-generated-metadata-spool.md) integrates a bounded
checked name plan and caller-supplied central-directory scratch through the
public fresh PPTX writer. The final evidence bundle records the matched policy
comparison and the operation allocator window gate. Caller scratch storage is
separate and grows with serialized directory bytes; this is not a constant-RSS
or all-PPTX claim. Ordinary constructors retain their metadata indexes.
The full non-iWork objective remains open, including logical append, broader
CRUD and native/source coverage, arbitrary repackaging and measured parallel
scaling. The [next batch](results/change-0478/next-work.md) starts with measurement
of the existing DOCX append path before proposing a bounded implementation.

The final 24-process, 720-sample matrix passes byte and reopen verification.
All 180 spool allocator samples have a 432,436-byte operation peak, versus
8,875,252 bytes for the 8,192-slide control. Scratch extents are 4,074, 40,842
and 1,237,692 bytes at 8, 256 and 8,192 slides. Small-deck normal mean latency
increases 3.11% and 3.31% in the two repeats; no registered 5% review threshold
is crossed. This is an operation-heap result with separate caller storage.

## Central-directory scratch implemented; name indexes remain (0477)

[0477](changes/0477-zip-central-directory-spool.md) removes growing central
header/name retention in an explicit low-level spool mode. Its measured heap
peak is flat across three tested member counts, while caller scratch and
per-member allocations/I/O remain real costs. ZIP Office and OPC validation
indexes still grow; the public PPTX route has not met its total-memory
requirement. A checked generated-name plan and semantic scratch integration are
next. The full non-iWork goal, including broader CRUD, source/native variants
and scaling, remains open.

## Repeated compressor allocation removed; total-memory gap remains (0476)

[0476](changes/0476-zip-deflate-state-reuse.md) implements the allocation owner
selected by 0475. All 720 main samples preserve output and the large operation
reduces requested bytes by 99.538473%. Peak heap remains approximately
8.88 MB for 8,192 slides, with retained name/directory metadata. This advances
implementation without closing the explicit total-memory requirement, logical
append, native breadth, source variants, repackaging or scaling.

## PPTX allocation ownership investigated (0475)

[0475](changes/0475-pptx-streaming-attribution.md) adds repeated attribution
to the 0474 memory-growth finding. Exact writer context is separated from
materialized preflight, raw traces and the failed demangled Heaptrack filter
are retained, and a separately declared mangled-symbol export corrects that
filter. This advances diagnosis; it does not close bounded total memory,
logical append, native breadth, arbitrary repackaging or scaling. The full
non-iWork goal remains open.

## PPTX fresh streaming memory requirement remains open (0474)

[0474](changes/0474-pptx-streaming-operation-memory.md) adds a missing public
fresh-creation baseline and rejects constant total memory for the tested path:
operation peak rises from 435,541 to 8,875,092 bytes over 8/256/8,192 slides.
All 360 samples and authored-slide oracles pass, with zero allocator live exit
delta. This advances measurement, not closure of the explicit-window requirement.
ZIP/OPC metadata and repeated allocation work need profiling and implementation.
Fresh creation remains distinct from logical append, Part addition and edits
followed by repackaging. Native breadth, source variants, scaling and the full
non-iWork goal remain open.

## DOCX fresh-creation evidence added (0473)

[0473](changes/0473-docx-streaming-operation-memory.md) advances the explicit-window
streaming requirement with 360 samples and full paragraph/run/package oracles.
The observed operation allocation peak is constant at 414,732 bytes over the
three tested sizes, separately from the 64-byte scratch reservation. Production
and the default matrix are unchanged. This adds fresh plain DOCX coverage;
logical append, Part addition, arbitrary repackaging, PPTX streaming, native
breadth, source variants and parallel scaling remain open. The full non-iWork
goal is not complete.

## Current scoped progress (0472)

[0472](changes/0472-xlsx-plain-cell-tags.md) retains plain-cell tag elision after
a 14.330% whole-process allocation reduction and 1,263 passing XLSX tests.
The full guard retains 86 latency flags; there is no registered latency or
peak-memory claim. Independent audit identifies public DOCX fresh streaming
creation as a missing allocator/scaling measurement, distinct from buffered
creation and the other three append meanings. Native, source-variant, CRUD,
streaming and scaling requirements remain open; the full goal is not complete.

## Latest measured rejection (0471)

[0471](changes/0471-xlsx-rewrite-buffer-lifetime.md) rejects a safe explicit
release of obsolete worksheet rewrite bytes because matched measurements do
not demonstrate the required peak-memory benefit. Seven-row diagnostic ABBA,
full-guard results and Heaptrack totals remain reproducible. This closes one
hypothesis, not a goal requirement or the broader non-iWork program. Larger
snapshot allocation and duplicate parsing work, CRUD coverage, source variants,
bounded streaming and measured scaling remain open.

## Current audit: 0470 bounded XLSX pass reuse (2026-09-08)

The [0470 record](changes/0470-xlsx-empty-web-proof.md) advances the measured
ordinary XLSX commit bottleneck: compaction can establish empty web bindings
without another worksheet traversal, while every unproven input retains the
original reader and error phase. All 1,257 XLSX tests and six scoped gates
pass. The protocol retains six-row ABBA, full-default guard and whole-process
allocation evidence, with its exact results and limitations recorded separately.

This changes neither the default scenario matrix nor the taxonomy's remaining
correctness-only mappings. Eager parsing and lossless snapshot scans remain
substantial leads. Full native-producer, physical-cold/range, bounded-streaming,
parallel-scaling and comprehensive CRUD evidence is not established by this
batch. Existing caches and publication checks retain their bounds. The full
non-iWork goal remains open; no completion or general speedup claim follows.

## Current audit: 0466 dense XLSX investigation (2026-09-08)

[0466](changes/0466-xlsx-dense-commit-profile.md) advances the required
profile-before-optimization work for the default dense one-percent XLSX
commit/save lead. It retains repeated normal timings, whole-process counters,
allocation stacks, initial incomplete CPU callchains and a same-source diagnostic
frame-pointer build. Supplemental normal p50 is 399.470314/398.785007 ms.
Wrapper recovery and postprocessing overlap remain explicit. Production,
37-case/201-row default coverage, 11 measured/22 correctness-only mappings,
and all accepted ADR contracts are unchanged. This is actionable profiling
evidence, not a production speedup or closure of the program's largest-bottleneck,
native, cold/remote, bounded-streaming or scaling requirements. The goal remains open.

All 18 focused Python tests, exact cross-process summary replay, fresh-copy
portable verification and postcleanup verification pass. The 103-artifact seal
is complete; both temporary executable copies and generated bytecode are removed.

## Current audit: 0465 checked default coverage (2026-09-07)

0465 completes the measured capture for the existing materialized
`odp_existing_append_lifecycle` case. Preflight preserves all prior 198 row
identities and the checked catalog now binds 37 default cases, 201 rows and 31
corpora. The mixed append-incremental category has exactly one measured ODP
row and two correctness-only fresh-streaming rows, enforced by the validator;
the taxonomy is 15 categories, 33 mappings, 11 measured and 22
correctness-only.

The formal lanes retain 6,030 normal samples and 90 ODP-only allocator samples.
ODP normal p50 is 1.691887/1.687518 ms (tiny), 67.786424/67.745553 ms
(medium) and 136.334131/136.843521 ms (large) in R1/R2. All four lanes,
353 harness tests (one ignored), warning-denied Clippy, rustdoc, scoped format,
167 latest Python tests, both full-report CRUD validators and boundaries pass.
Initial stale Python hash/count-pin failures remain retained. The sealed
precleanup verifier, five resealed negative probes, finalize precleanup,
fresh-copy flagless portable verification and owned cleanup pass; the portable
seal is unchanged and temporary directories are absent. This evidence makes no
regression, speedup, independent native-producer, bounded-memory streaming or
scaling claim; the broader non-iWork goal remains open.

## Current audit: 0464 evidence (2026-09-07)

0464 adds descriptive harness evidence for a generic PPTX source/destination
pair and changes no production implementation. The pair is a same-source-derived
positive control: its destination is a Litchi self-copy publication of the
source archive, not an independently authored native package. The frozen matrix
contains eight reports and 240 samples across serialized R1 forward/R2 reverse
normal and operation-scoped allocator bytes/range lanes. API-sum p50 values in
R1/R2 are 1.9536/1.9380 ms (normal bytes), 1.9872/1.9865 ms (normal range),
2.0768/2.0705 ms (allocator bytes), and 2.1182/2.1213 ms (allocator range).

The logical-range adapter returns at most 256 bytes with zero fixed delay. Each
formal output is three slides and 55,891 bytes and passes the independent
package oracle. The API sum covers four public calls and excludes input/setup,
sink/adapter construction, artifact/oracle work and teardown; it is not a
contiguous end-to-end measure. Normal allocator totals are unavailable by
design. No native Office acceptance, independent-producer, network, physical-I/O,
cold-cache, scaling or copied-byte/compression claim follows.

Across all 120 allocator rows, summing the four allocator regions gives exactly
8,623,012 allocated bytes, 9,616 allocation calls and 1,455 reallocation
calls. This does not infer a lifecycle peak by summing phase peaks. Bound GNU
time resource logs report maximum RSS in kB for R1 normal-bytes/normal-range/
allocator-bytes/allocator-range as 16,912/16,908/17,068/17,152 and for R2 as
16,900/16,684/16,612/16,584. The overall 16,584–17,152 KiB (~16.2–16.75 MiB)
range is process-level and includes setup, oracle work and teardown. The
summary does not derive this GNU-time resource field; raw
`captures/R1/*/resource.log` and `captures/R2/*/resource.log` files are bound
in the bundle.

The final source review and harness Clippy receipt are clean at custody epoch
`a35f4507a74e7678f29f91dd6857d07a87579ba872785a1eaf19e296540debbe`; harness
validation records 391 passing tests with one ignored, and all 240 formal
samples plus focused smoke (four positive and seven negative cases) pass.
LibreOffice 26.2.5.2 saves the formal three-slide output. Native-r2 passes the
scoped application-save plus source-backed/eager validation of all three slide
counts, slide sizes and ordered text; its saved output is 39,830 bytes with
SHA-256 `d2ca3109448f5a0797f00561beff4eafa84616f7dbf897ff2456667fbf555915`.
All three post-save source-backed image inventories are unavailable because
`UnsafeEdit` / `source-backed picture inventory` refuses markup compatibility.
The failed native-roundtrip-r1 receipt and raw saved output remain preserved.
This does not establish full Office acceptance; image equivalence and rendering
compatibility remain unproven. A supplemental inventory-only source epoch
changes only that diagnostic binary; its build and warning-denied Clippy receipt
pass. Source compatibility records 7,031
unchanged files, while the measured copy Rust/binary/captures remain immutable
and no retiming occurs. The documentation
receipt and owned performance-crate scoped format check pass. A broad `cargo
fmt --all --check` receipt remains failed only at pre-existing formatting in the
out-of-scope iWork `crates/litchi-keynote/src/document.rs`; no full-workspace
format pass is claimed. Boundary checks pass. The corrected `profiling-r1`
receipt supersedes the initial UID-0-unavailable diagnostic. It passes a
whole-process user-counter and DWARF profile over three warmups and 30 retained
samples (33 checked iterations); separate counter and sampling runs add 60
diagnostic samples outside the 240 formal samples. Totals are instructions
1,758,215,198, cycles 631,274,012, branches 327,372,796, branch misses
3,514,541, cache misses 2,295,690, and page faults 1,709. PMU counters ran at
83% because of multiplexing, and the raw zero cache-reference value is
uninterpreted. The 20-sample cycle symbol view places
`sha2::sha256::x86_sha::compress` at 19.96% and
`zlib_rs::inflate::inflate_fast_help_avx2` at 14.25%; the whole-process scope
includes binary hashing, setup and oracle work, so this is diagnostic only and
does not establish operation hotspots or stable ranking. Sealed precleanup, fresh-copy portable verification and resealed-summary
tamper rejection pass. Owned cleanup removed 12 files totaling 252,885,687
bytes plus two archived audit temporaries. The full non-iWork goal remains open, counts remain 439
selectors / 36 defaults,
0463 remains retained, and iWork is untouched. See the [0464 summary](results/change-0464/summary.json),
[0464 protocol](results/change-0464/protocol.json),
[pair](results/change-0464/pair.json), [fixture provenance](results/change-0464/fixture-provenance.md),
[source review](results/change-0464/source-review.md), [native-r1 receipt](results/change-0464/native-roundtrip/receipt.json),
[native-r2 receipt](results/change-0464/native-roundtrip-r2/receipt.json),
[profile summary](results/change-0464/profiling-r1/profile-summary.json),
[top symbols](results/change-0464/profiling-r1/samples/top-symbols.txt), and
[source compatibility](results/change-0464/source-compatibility.json).

## Current audit: 0463 evidence (2026-09-07)

0463 retains a private writer-origin proof for the ordinary ODP publication
path after the frozen 3% gate passes. The private ODP serializer records proof
accounting around the existing common `PackageWriter` audit path, which is
unchanged. The A1/B1/B2/A2 evidence contains 24
reports and 720 samples. Normal p50 candidate-minus-baseline deltas are
-10.0491% / -8.9678% for tiny, -6.9642% / -7.3050% for medium, and
-7.3657% / -6.8199% for large in R1/R2. All normal and allocator bootstrap
upper bounds are below zero; no adverse >5% elapsed or RSS flag or allocation
increase review flag is present.

Allocator bytes fall by 2,087,682 / 11,072,106 / 20,202,090 for
tiny/medium/large in both repeats. Every lane reduces allocation calls by
1,039, reallocations by 93 and deallocations by 946; peak above entry changes
are 0 / -49,674 / -15,438 bytes and retained-live deltas remain zero. The proof
is eligible only for the exact writer/source owners and conservative audit
bounds; candidate reopen/readback, source precheck, media/domain checks,
no-op and patch behavior remain. The source review reports 381 ODP tests and
warning-denied Clippy passing.

Supplementary phase clocks exclude setup, warmups and checks; whole-process
counters include them. The proof changes the commit validation path, whose p50
phase delta is -12.6863% / -13.0797%; other phase movement is diagnostic. The
full non-iWork goal remains open, counts remain 439 selectors / 36 defaults,
0460 remains accepted, and iWork is untouched. The harness and final gates are
complete: 387 harness tests pass with one ignored, for 768 passed ODP/harness
tests in total. Warning-denied rustdoc, scoped formatting, boundaries,
precleanup and source replay, fresh-copy portable replay, resealed +1ns tamper
rejection, and owned cleanup of four executables totaling 233,058,712 bytes
pass. See the [0463
comparison summary](results/change-0463/summary.json), [phase summary](results/change-0463/phase-summary.json),
[source review](results/change-0463/source-review.md), and [proof design](results/change-0463/proof-design.md).

## Current audit: 0462 evidence (2026-09-07)

0462 records partial normal-lifecycle improvements from a private seventeen-key
shape-attribute index, but rejects the candidate after the predeclared 3%
medium/large gate. The frozen A1/B1/B2/A2 evidence contains 24 reports and 720
samples. Normal p50 candidate-minus-baseline deltas are -2.3623% / -3.0049%
for tiny, -3.4649% / -3.1917% for medium, and -2.6602% / -2.6279% for large
in R1/R2. Both large rows miss the threshold while every normal p50 bootstrap
interval remains below zero. Allocation bytes, calls, reallocations,
deallocations, regional peak and retained-live metrics are exactly unchanged;
the R1 tiny allocator interval crosses zero, and no adverse >5% elapsed or RSS
flag is present. The index adds 280 bytes per element, so no retained speed or
memory benefit claim follows.

Manual assembly review finds `shape_builder` changing from 17 generic getter
static call sites and a 1,400-byte frame to 17 indexed typed getter call sites and a 1,688-byte
frame. Generic `ElementAttrs::get` remains 328 bytes. Indexed known-key hits
bypass the cached loop, but the defensive fallback remains; the automatic
elimination field is a flawed direct-call heuristic and is excluded from the
evidence. Supplementary phase clocks cover public API calls and exclude setup,
warmups and checks; whole-process counters include them and remain diagnostic.

Candidate validation records 379 ODP tests and warning-denied owner Clippy as
passing; the initial compilation failure remains preserved. All 387 harness
tests pass (one ignored). Both Rust files are restored byte-exact to baseline
revision `dbd2f8ece`; final Clippy/docs/format/boundaries, portable replay, tamper
rejection and owned temporary cleanup pass.
No selector or corpus coverage is added: the registry remains 439 selectors /
36 defaults, 0460 remains accepted, the full non-iWork goal remains open, and
iWork is untouched. See the [0462 comparison summary](results/change-0462/summary.json),
[phase summary](results/change-0462/phase-summary.json), [source review](results/change-0462/source-review.md),
and [assembly review](results/change-0462/assembly-review.md).

## Current audit: 0461 evidence (2026-09-07)

0461 records a partial normal-lifecycle improvement from splitting ODP
attribute matching from value decoding, but rejects the candidate after the
predeclared 3% practical gate. The frozen A1/B1/B2/A2 evidence contains 24
reports and 720 samples. Normal p50 candidate-minus-baseline deltas are
-2.0578% / -2.2940% for tiny, -2.0604% / -3.8754% for medium, and
-2.1547% / -3.1281% for large in R1/R2. R1 medium and large fail the gated
threshold; all normal p50 bootstrap upper bounds remain below zero. Allocation bytes,
allocation calls, reallocations, deallocations, regional peak and retained-live
metrics are exactly unchanged, and no adverse >5% elapsed or process-RSS flag
is present. No retained speedup claim follows.

The assembly receipt confirms that the out-of-line `ElementAttrs::lookup` body
is removed, namespace/local-name checks are inlined into `get`, and the `get`
stack frame moves from `0x148` to `0x128`. Supplementary phase clocks measure
public API calls separately from the primary matrix. Whole-process counters
include setup, warmups and checks and are not operation-only totals. Neither
diagnostic establishes causal attribution.
The candidate records 372 ODP tests, warning-denied all-target Clippy and
scoped formatting as passing, plus 387 harness tests with one ignored. Source
restoration to `05f432d48`, final Clippy/docs/formatting/boundaries, portable
verification, tamper rejection and owned temporary cleanup pass.

The experiment adds no selector or corpus coverage; the registry remains 439
selectors / 36 defaults. 0460's fused staging optimization remains accepted,
the full non-iWork goal remains open, and iWork is untouched. See the
[0461 comparison summary](results/change-0461/summary.json), [phase summary](results/change-0461/phase-summary.json),
[source review](results/change-0461/source-review.md), and [assembly receipts](results/change-0461/candidate-assembly.json).

## Current audit: 0460 evidence (2026-09-07)

0460 retains the private ODP fused staging/source-scanning optimization. The
authoritative A1/B1/B2/A2 lifecycle matrix contains 24 reports and 720 samples.
Normal p50 candidate-minus-baseline deltas are -3.5430% / -3.8257% for tiny,
-6.0849% / -5.0679% for medium, and -5.4732% / -5.8268% for large in R1/R2.
All four medium/large rows pass the predeclared 3% improvement gate with
negative independent bootstrap upper bounds; no adverse >5% elapsed or RSS
flags are present.

Allocator p50 deltas are -4.7840% / -4.1309% for tiny, -7.3100% / -6.7785%
for medium, and -7.3197% / -6.9282% for large. Every allocator lane reduces
allocated bytes by 4,642, allocation calls by 16, reallocations by 12 and
deallocations by 4, while regional peak and retained-live deltas are unchanged.
Supplementary phase clocks show transaction p50 reductions of -24.2681% /
-24.0697%; whole-process counters show instructions -6.1959%, cycles -4.9100%
and branch misses +0.5503%. The phase clocks cover public API calls without
setup, warmups or checks; the whole-process counters include those activities.
Neither scope establishes causal API attribution.

The source review finds the fused namespace-aware traversal preserves state
machines, error precedence, BOM-relative spans, limits, source ownership and
readback/patch/no-op contracts. The corrected owner retry passes 371 tests and
all-target warning-denied Clippy passes. The optimization is retained under
this scoped evidence; no selector or corpus coverage is added, counts remain
439 selectors / 36 defaults, the full non-iWork goal remains open, and iWork is
untouched. All builds, 387 harness tests (one ignored), documentation, boundaries,
portable verification and owned temporary cleanup pass. See the [0460 comparison
summary](results/change-0460/summary.json), [phase summary](results/change-0460/phase-summary.json),
and [source review](results/change-0460/source-review.md).

## Current audit: 0459 evidence (2026-09-07)

0459 makes measurement progress: it repairs sampled phase ancestry and rejects
a local-name-first attribute lookup experiment after 720 matched operations.
No tested production change is retained. Candidate readback and repeated staging
scans now have stronger profile attribution, while caching family reopen is
rejected as low impact. Every regression flag remains visible.

The full non-iWork objective remains open. This batch adds neither a CRUD
selector nor source/output/native/scaling coverage; counts remain 439/36.
See the [0459 evidence](results/change-0459/README.md). Historical audits below
retain their original scoped conclusions.

## Current audit: 0458 evidence (2026-09-07)

0458 completes a supplementary ordinary ODP append phase diagnostic: 24 lanes,
720 samples, separate profiles, exact allocation-volume conservation, and
recomputable sealed evidence. Commit is the largest individual large-input
phase; opening plus transaction setup consumes about half the phase time.
Sampled stacks do not resolve phase-marker ancestry, so internal cost ranking
needs another diagnostic. This batch adds no production optimization or new
CRUD/source/output coverage. The registry remains 439 selectors / 36 defaults.

The full non-iWork goal remains open: wider selective CRUD, source/input/output
matrices, native application roundtrips, cold/range behavior and measured
bounded-worker scaling still need completion. The historical 0457 audit below
retains its original scope. See the [0458 evidence](results/change-0458/README.md).

## Current audit: 0457 evidence (2026-09-07)

The full non-iWork goal remains open, but the bounded existing-ODP evidence
now has a formal current-revision ordinary control. The
`odp_existing_append_lifecycle` baseline retains two repeats, normal and
allocator lanes, 64/4,096/8,192 source slides, three warmups and 30 samples
per lane (360 samples across 12 reports); the final candidate summary retains
the same 360-sample, 12-report matrix. The paired
`odp_source_tail_append_lifecycle` path validates bounded source XML and
replays the changed ZIP member through a sequential sink. It is a specialized
publication plan with a different retained-result contract and does not
integrate the ordinary `edit::Snapshot`/`edit::Transaction`/`edit::Patch`/
`edit::Commit` lifecycle. Those owned ordinary APIs are implemented; source-tail
integration with them is the remaining gap.

Normal p50 candidate-minus-control deltas are -33.352% / -33.185% for the
64-slide shape, +3.262% / -2.461% for 4,096 slides, and -1.940% / -4.221% for
8,192 slides in R1/R2. The allocator evidence reports control versus
source-tail regional peak above entry of 781,342 vs 620,381 bytes (tiny),
18,027,568 vs 620,385 (medium), and 35,958,388 vs 620,385 (large). Allocated
bytes are 10,613,383 vs 2,607,922, 110,226,105 vs 36,593,937, and 211,442,207
vs 71,164,177. These are operation-scoped allocator observations; source and
fixture ownership at entry are excluded, and candidate allocation volume
continues to grow with document size.

The comparison receipt explicitly withholds ordinary Commit/Patch speedup or
regression, general CRUD or retained-result equivalence, causal, bounded-memory,
scaling, cancellation and physical-I/O claims. R1 normal tiny is +8.301% and
R2 normal tiny is +6.517%; allocator tiny is +10.184% in R1, while allocator
large is -5.113% / -5.740% in R1/R2. The remaining process-peak rows are
within the five-percent threshold. The final identity-bound set has ten native
records and three synthetic source/output pairs passing the independent ZIP/XML
oracle, with no Office application launch. ZIP, ODF-common and ODP suites pass
1,348 tests with three ignored (ODF-common 499, ZIP 481, and ODP 368); the
harness passes 383 with one ignored, OPC passes 497 with one ignored, both
ZIP/XML sanitizer fuzz lanes retain 1,000 runs, and candidate derivation and
comparison receipts pass. Initial pre-fix large-normal profiles retain 1,000 sampled
`cycles:u` observations: candidate SHA-256 compression is 11.00%, XML
`validate_name` 9.08%, `memcmp` 6.01% and `validate_start_element` 5.42%;
control `memcmp` is 9.67%, Quick-XML attribute iteration 6.00%, `memmove`
5.76% and namespace-prefix resolution 5.24%. These are whole-process sampled
stacks, not API-attributed CPU percentages or causal hotspots, and no hard
counter totals were collected. These whole-process samples remain initial
pre-fix diagnostics, not final candidate attribution; no sampling profile was
captured for the final candidate epoch. Candidate profile checking passes for
this initial capture; the unchanged control report/raw profile passes the
existing amended oracle after the original frame-pointer expectation failed.

The source mapping verifies 439 selectable selectors and 36 defaults. The
coverage index keeps 15 categories, 33 representative mappings, 10 measured
mappings and 23 correctness-only mappings; its ODP append row remains
correctness-only because generated-per-run 0457 corpora do not satisfy the
index's checked-catalog/default measured-status contract. The ordinary control
is nevertheless recorded as the formal baseline in the 0457 evidence and
reports. The [final source read review](results/change-0457/final-code-review.md)
records the retained-fragment lease fix as resolved and found no additional
verified blocker. The sealed batch passes precleanup, portable-copy replay,
and three altered-copy rejection checks. Owned temporary artifacts are removed
with a retained inventory; see the [bundle](results/change-0457/README.md). See the [final candidate summary](results/change-0457/candidate-final/summary.json)
and [final comparison](results/change-0457/comparison-final.json).

## Current audit: 0456 evidence (2026-09-07)

The full non-iWork goal remains open. [0456](changes/0456-zip-shared-payload-framing.md)
retains verified/shared ZIP payload storage during publication, removing a
payload-sized preparation allocation and copy. Media publication allocates
4,243,083 bytes (-79.85%) and peaks 593,892 bytes above entry (-96.59%); output,
read/write work and managed budgets remain unchanged. Plain publication metadata
costs 3,360 more allocated bytes. The complete document lifecycle is not bounded
by this change, and managed admissions still include the old writer allowance.

The 720 formal observations retain all 15 timing flags. Ordinary bytes/media API
medians regress 15.411%/16.347% with higher page faults. Separate 240-sample fixed
allocator-policy diagnostics improve API medians 2.626–10.266%. Acceptance is for
the memory reduction with allocator-sensitive latency disclosed, not a general
speedup. Release checks pass 2,654 tests (7 ignored), strict lint, documentation,
formatting, workspace/boundaries and 1,000 ASAN fuzz iterations. Native self-pair
output remains exact; six distinct source/destination probes refuse shared-graph
incompatibility. An evidence SHA transcription failure and corrected retry are
retained alongside the initial compile correction.

Coverage remains 438 selectors, 36 defaults, 15 categories, 33 representative
mappings, 10 measured mappings and 23 correctness-only mappings. No coverage row
is promoted. Broader CRUD, native application roundtrips, distinct-package pairs,
physical cold I/O, bounded existing-document append/repackaging and representative
worker scaling remain incomplete. [Next-work notes](results/change-0456/next-work.md)
identify the streaming XML/ODP append dependency. See the
[0456 bundle](results/change-0456/README.md) for replay and owned cleanup.
User-owned `docs/GOAL.md` remains unchanged.

## Prior audit: 0455 evidence (2026-09-07)

The full non-iWork goal remains open. [0455](changes/0455-zip-preservation-transfer-chunks.md)
retains a 64 KiB ZIP preservation buffer: 256 fewer media publication reads and
sink writes, identical byte counts/output, and simulated range API median
improvements 4.775%/4.449%. It adds 32 KiB fixed stack per active publication;
operation allocation counts and bytes do not change.

The evidence contains 720 formal samples, 240 separate ordinary confirmation
samples and 240 fixed allocator-policy diagnostic samples. Original adverse
whole-process CPU counters and all 19 formal timing flags remain disclosed.
Slow/high-page-fault processes appear in both builds; controlled allocator
policy supports paging variability without reconstructing historical mapping
decisions. No CPU or large bytes-only gain is claimed. Release checks pass
2,188 tests (6 ignored), strict lint, documentation, formatting, minimal workspace,
boundaries and 1,000 sanitizer fuzz runs. The native self-pair produces its
previous exact output via bytes/range; this adds no native-application roundtrip
or distinct-package evidence.

Coverage remains 438 selectors, 36 defaults, 15 categories, 33 representative
mappings, 10 measured mappings and 23 correctness-only mappings. This transfer
optimization promotes no coverage rows. Broader semantic CRUD, native application
roundtrips, distinct-package pairs, physical cold I/O, bounded existing-document
append/repackaging and representative worker scaling remain incomplete.
See the [0455 bundle](results/change-0455/README.md) for capture custody,
verification and owned cleanup. User-owned `docs/GOAL.md` remains unchanged.

## Prior audit: 0454 evidence (2026-09-07)

The full non-iWork goal remains open. The 0454 production source epoch
implements three recorded boundary fixes: the baseline optional
producer-visible-name refusal, the name-only historical intermediate
candidate's noncanonical relationship XML refusal, and the ordinary
authored-XML compactness refusal through OPC source XML proof. The opaque
`source_xml`, `checked_range`, `AuthoredXmlFragment`, and splice publication
path preserves source whitespace and literal namespace grammar while retaining
source identity, limits, resource accounting, cancellation, and exact no-op
copying. Ordinary authored XML remains strict. This capability evidence makes
no baseline-refusal speedup claim.

The final native inventory covers 588 files and 189 image-bearing direct-picture
rows: four self-pair cases publish and 185 outcomes remain unchanged. Each
case reopens the same unmodified archive independently as source and
destination. The pinned LibreOffice QA fixture additionally passes the exact
ZIP/XML preservation checks in its external bytes/range captures. This is a
self-pair proof, not an independent package pair or a native-application
roundtrip ([inventory](results/change-0454/final-native-inventory.json),
[outcome comparison](results/change-0454/final-outcome-comparison.json)).

The formal capture passes all 18 lanes, 16 provider plus 2 external, with 30
samples each: 540 retained samples
([measurements](results/change-0454/measurements.md),
[machine-readable rows](results/change-0454/measurements.json)). The matched
whole-API p50 movement stays within about 1.3%, no RSS comparison exceeds 5%,
and all 15 absolute review flags remain retained, including the three positive
range/media-rich R2 open p99 flags (+5.84%, +9.93%, +7.87%). The
range/media-rich R1 open-source p99 flag is -8.97% and does not repeat.
Logical read/work counters remain consistent; scheduler or sleep variability is
a plausible inference for the tail behavior, not a proved cause.

The core release epoch passes build, format, strict, harness, OPC, PPTX,
documentation, workspace, boundary, oracle, and native checks: 497 OPC tests,
854 PPTX tests, and 381 aggregate harness tests pass. The standalone amended
ASAN fuzzer also passes its 1,000-run seed-454 smoke. The original fuzzer
build failure is historical; `fuzz-source-amendment.json` records the only
post-measurement source change, an isolated `parse_opc` harness bridge. The
amended fuzzer source manifest is separate from the measured production
manifest, which remains unchanged
([amendment](results/change-0454/fuzz-source-amendment.json),
[validation history](results/change-0454/validation-notes.md)).

Custody corrections are disclosed. The Markdown-renderer correction temporarily
rewrote `protocol_sha256` in all 18 formal receipts; the renderer-r1 restoration
reconstructs those receipt bindings and retains the intermediate hashes, so
custody is reconstructed rather than uninterrupted
([restoration](results/change-0454/custody-corrections/renderer-r1/restoration.json)).
Raw timing reports, sampling, statistics, and timing inputs remain unchanged.
The separate derivation amendment records the corrected `derive-final.py`
display lookup apart from the original `derive.py`; it does not create a new
measurement ([derivation amendment](results/change-0454/derivation-amendment.json)).

The coverage taxonomy remains 438 selectors, 36 defaults, 15 categories, 33
representative mappings, 10 measured mappings, and 23 correctness-only
mappings ([coverage index](crud-coverage-index-v1.json)). The 540 formal samples
are one matched workload control and do not promote those coverage rows. Native
application roundtrip, distinct-package, cold-cache/physical-I/O, allocator,
scaling, and the broader semantic CRUD rows remain unmeasured or incomplete.

Final precleanup and post-cleanup verification pass. Owned temporary binaries,
fuzz build files, and all seven generated PPTX outputs have been removed; the
retained evidence records their identities. No future measured hotspot is asserted.
The next measurement scope is the 64 KiB destination pass-through hypothesis:
compare a bounded 32-to-64 KiB or adaptive candidate with exact request
histograms, timing, resource and sink gates. It remains a measurement scope,
not an established bottleneck or performance claim. The full non-iWork goal
remains open.

The sections below retain historical audit records. Their older “current” labels
refer to their original audit dates and must not override this section.

## Historical evidence through 0453

[0453](changes/0453-pptx-shared-decoded-payload.md) removes the duplicate staged decoded payload after successful PPTX image/chart
capture. Both allocator repeats save 16,777,408 planning allocated/retained bytes;
bytes/media API medians improve about 3%. Full fallback allowance remains and
media destination staging grows 128 bytes. All 1,702 final tests and 1,000 existing
OPC fuzz runs pass. Primary plain p99 increases remain visible; a separate fixed
240-sample ABBA investigation does not reproduce them.

Registry/default counts remain 438/36; representative coverage remains 15
categories/33 mappings/10 measured/23 correctness-only. Native application
breadth, cold I/O, bounded existing append, repackaging and scaling remain
required. This is progress; the full non-iWork goal remains active and uncompleted.

## Earlier evidence through 0452

[0452](changes/0452-pptx-retained-capture.md) integrates retained OPC captures into PPTX image/chart plans with
independent publication reservations. Complete simulated-range API p50 improves
31.324%/31.048%; separate balanced bytes/media confirmation supports about 8%
improvement with about 6% extra planning time. Publication source data reads
fall to zero. Plan-held source reservations increase 16,815,144 bytes until drop;
existing decoded staging remains. All 1,700 final tests and 1,000 fuzz runs pass.

Registry/default counts remain 438/36; representative coverage remains 15
categories/33 mappings/10 measured/23 correctness-only. Native application
breadth, cold I/O, bounded existing append, repackaging and scaling remain
required. Shared staged decoded ownership is the next local memory opportunity.
This is progress; the full non-iWork goal remains active and uncompleted.

## Earlier evidence through 0451

[0451](changes/0451-opc-combined-capture.md) implements combined OPC read/authorization under cache, source identity,
work and memory/object budgets. Eight deterministic I/O cases preserve whole
publication output with a compressed source pass removed. Tests cover ordinary
and combined loader coordination, rollback, short reads, budget boundaries and
native-input compressed transfer after package/data drop. The API is opt-in;
semantic PPTX plans have not adopted it and no complete latency gain is claimed.

Registry/default counts remain 438/36; representative coverage remains 15
categories/33 mappings/10 measured/23 correctness-only. Reusable PPTX publication
reservations, matched timing, native breadth, cold I/O, bounded existing append,
repackaging and scaling remain required. The full non-iWork goal remains active.

## Earlier evidence through 0450

[0450](changes/0450-zip-combined-capture-decode.md) adds the low-level combined ZIP capture/decode primitive needed for OPC
first-read transfer authorization. Deterministic I/O evidence removes a compressed
source pass while preserving decoded bytes and writer token output. All 455 ZIP,
436 OPC and 59 PPTX targeted tests pass, as do strict lint and the instrumented
1,000-run fuzz smoke. This is a measured enabler; OPC/PPTX has not adopted it yet.

No timed or native coverage is promoted. Registry/default counts remain 438/36;
representative coverage remains 15 categories/33 mappings/10 measured/23
correctness-only. Combined OPC reservations, token/cache lifetime, matched timing,
native breadth, cold I/O, bounded existing append, repackaging and scaling remain
required. The full user-owned non-iWork goal remains active and uncompleted.

## Earlier evidence through 0449

[0449](changes/0449-pptx-caller-source-attribution.md) corrects the attribution used to rank the next PPTX optimization: untimed
harness hashing is about half of lifecycle SHA period, and publication read totals
combine source compressed capture with destination passthrough. Replaying all
240 existing samples confirms source-cache hits rather than cold rereads during
publication. This is diagnostic evidence; no new workload or native test ran.

Registry/default counts remain 438/36 and representative coverage remains
15 categories/33 mappings/10 measured/23 correctness-only. Native breadth, cold
I/O, bounded existing append, repackaging and scaling remain incomplete. The full
user-owned non-iWork goal remains active and uncompleted.

## Earlier evidence through 0448

[0448](changes/0448-pptx-minimum-service-pacing.md) completes the scoped minimum-service timer calibration identified in 0447.
Eight reports/240 samples and four profiles support retention of the opt-in
policy; both plain median gates and all service floors pass. All 381 harness
tests pass and strict lint adds zero diagnostics. No repeat trigger exceeds 5%.
This is a measured tooling enabler, with no production or physical-network claim.

Registry/default counts remain 438/36; the representative index remains
15 categories/33 mappings/10 measured/23 correctness-only. Its default full-run
contract is not promoted by opt-in captures. Native breadth, cold I/O, shared-link
concurrency, bounded existing append, repackaging and full scaling evidence are
incomplete. The full user-owned non-iWork goal remains active and uncompleted.

## Earlier evidence through 0447

[0447](changes/0447-pptx-range-transfer-pacing.md) fills a range-simulation tooling gap with optional configured transfer rate,
checked requested-delay counters and matched managed PPTX lifecycles. Eight
reports/240 samples and four profiles retain exact source-work/output equality.
The 378 harness tests pass; strict lint adds zero diagnostics. Twelve repeat
tail flags and sleep-granularity limitations remain explicit. This is a measured
enabler, with no production speedup, actual network-bandwidth or allocator claim.

The representative index intentionally reserves measured status for its default
full-run contract; later opt-in timing bundles do not automatically promote
those rows. Counts remain 438 selectors/36 defaults and 15 categories/33 mappings,
10 measured/23 correctness-only. Shared-link concurrency, cold I/O, native
breadth, bounded existing append, repackaging and full scaling evidence remain
incomplete. The full user-owned non-iWork goal stays active and uncompleted.

## Earlier evidence through 0446

[0446](changes/0446-opc-owned-content-type-name.md) retains a one-line ownership optimization in content-type parsing.
The 720-sample matrix passes the allocation gate: 5.963%/5.990% fewer calls at
medium/large in both repeats. The separate normal latency gate fails; peak
memory is unchanged and no paired latency/RSS or repeat trigger crosses 5%.
All 458 OPC/373 harness tests pass, owner strict lint is clean and harness lint
adds no diagnostics. Independent oracles and corruption probes remain retained.

This improves the existing synthetic package-addition slice without promoting
semantic or native coverage. There remain 438 selectors/36 defaults and 15
categories/33 representative mappings, 10 measured and 23 correctness-only.
Native breadth, repackaging, semantic dependency closures, bounded existing
append, cold/range and scaling are incomplete. The full user-owned non-iWork
goal remains active and uncompleted.

## Earlier evidence through 0445

[0445](changes/0445-opc-part-add-plain-source.md) adds a matched plain-source Part-addition lifecycle and a single-build
24-report/720-sample calibration. Plain large normal p50 is about 18.9 ms versus
62.1 ms observed; this is observer overhead, not a production speedup. Allocation
calls/requested bytes/above-entry peaks match. One instrumented p99 repeat flag
is disclosed. The 373 harness tests, 37 topology tests and 53 corruption probes
pass; strict harness Clippy retains inherited debt with zero new diagnostics.

Plain stack evidence points to content-type map construction. Source inspection
finds an avoidable owned-string clone as the next small measured candidate;
required validation, freshness and memory accounting must remain. The standalone
selector count is 438, with unchanged defaults and semantic representative
coverage. Native breadth, repackaging, semantic dependency closures, bounded
existing append, cold/range and scaling remain open. The full user-owned
non-iWork goal is active and uncompleted.

## Earlier evidence through 0444

[0444](changes/0444-opc-part-add-baseline.md) adds a low-level source-backed OPC Part-addition baseline:
12 reports/360 samples, three sizes, normal and allocator repeats, actual
independently verified ZIP fixtures, 29 corruption probes, and complete source
and sink observations. The 372 harness and 458 OPC tests pass; strict harness
Clippy retains inherited debt with zero new diagnostics. No production code
changes or comparative performance claim are made.

The observed source reader accounts for 55.24% of whole-process sampled self
time and scans all ordinary ranges per read. A matched plain-source baseline
must precede production attribution for this path. This partially fills package
Part-addition evidence only: semantic owner creation, broader dependency closures,
repackaging, bounded existing append, native breadth, cold/range and scaling
remain open. The user-owned non-iWork goal remains active and uncompleted.

## Earlier evidence through 0443

[0443](changes/0443-odp-compact-fragment-frames.md) replaces copied namespace/local-name frame data
in ODP source-fragment scanning with exact element kinds. The 720-sample matrix
passes the frozen allocation-call gate at 14.886%/14.943% medium/large reductions.
The normal latency gate fails; timing costs, two instrumented tail flags and
eight repeat flags remain disclosed. Peak and retained bytes are unchanged.
The original scanner remains a byte-identical differential reference.

This removes one source of temporary allocation work. Repeated validation,
one-shot attribute lookups, Part addition, repackaging, bounded existing append,
native breadth, cold/range and scaling remain open. The full user-owned
non-iWork goal is unchanged and uncompleted.

## Earlier evidence through 0442

[0442](changes/0442-odp-shared-staging-traversal.md) combines three ODP staging XML traversals while
preserving complete-pass error priority. The 720-sample matrix passes the frozen
normal p50 gate: medium/large improve 9.830–13.307% across both repeats.
All five repeat flags remain disclosed. Peak and retained bytes are unchanged;
no general tail, RSS or bounded-append benefit is claimed.

The three auxiliary scans are consolidated. Source-fragment scanning, one-shot
cache costs, Part addition, repackaging, bounded existing append, native breadth,
cold/range and scaling remain open. This batch does not complete or narrow the
user-owned non-iWork goal.

## Earlier evidence through 0441

[0441](changes/0441-odp-shared-preservation-projection.md) shares the immutable
ODP preservation comparison projection and keeps the mutable draft detached.
The 720-sample comparison shows 9.834%/9.857% lower medium/large operation peak,
with unchanged retained live bytes. Its original practical gate failed;
retention is explicitly based on a post-hoc peak-memory review. All raw timing
costs and the baseline tiny p99 repeat flag remain visible. No normal latency,
RSS or bounded existing-append claim is made.

Repeated XML staging scans remain a larger CPU target. Part addition,
repackaging, bounded existing append, native breadth, cold/range and scaling
remain open. This measured ownership improvement does not complete or narrow
the user-owned non-iWork goal.

## Earlier evidence through 0440

[0440](changes/0440-odp-borrowed-attribute-namespaces.md) reduces temporary
namespace allocation work in owned ODP append: about 20% fewer medium/large
calls and 6.3–6.6% fewer requested bytes across both repeats. Exact archive,
semantic, preservation and patch gates pass. The main 720 samples, additional
120-sample confirmation, four selected profiles, 352 ODP tests and 368 harness
tests are retained. Peak and retained live bytes are unchanged; no normal
latency or RSS benefit is claimed. Original adverse tails and two excluded
overlapping attempts remain documented.

This closes one private ownership experiment. One-shot cache cost and repeated
staging/validation are separate next hypotheses. Bounded existing append,
Part addition, repackaging, native/producer breadth, cold/range input, scaling
and the remaining non-iWork goal stay open. The user-owned goal is unchanged.

## Earlier evidence through 0439

[0439](changes/0439-odp-existing-append-lifecycle.md) establishes the full
owned existing-ODP append interval that the older one-edit timer omitted.
Twelve reports/360 samples and two accepted profiles cover 64/4,096/8,192
source slides plus an opaque package member. The 368-test harness suite,
independent fixture/resource gates, and no-new-Clippy-debt comparison pass.
Large normal p50 is 170.742/175.416 ms, with 236,704,188 requested bytes and
a 39,890,548-byte region peak above entry. This is a baseline addition, with
no production optimization or bounded-commit-memory claim.

The next evidence question is operation-specific attribution of owned open
and commit validation, including repeated attribute/namespace work. Part
addition, arbitrary repackaging, native/producer breadth, cold/range input,
scaling and the remaining non-iWork goal stay open. The user-owned goal is
unchanged; this batch does not narrow or complete it.

## Earlier evidence through 0438

[0438](changes/0438-odp-markup-batching-negative.md) tested the proposed ODP
fixed-markup accounting change and rejected it under the predeclared 5%
practical-gain gate. The 24-report/720-sample comparison found only
0.703–1.982% normal p50 improvements. Production was restored; candidate
source, measurements, four profiles, and validation remain reproducible.

The next coverage target is logical append to an existing ODP through the full
owned open/transaction/commit/sequential-output lifecycle. The existing
`odp_semantic_one_edit_save` timer excludes opening. This is distinct from
fresh streaming creation, adding a package Part, and arbitrary repackaging.
The broader non-iWork goal remains open, including rich/native producer
coverage, cold/range I/O, and scaling. `docs/GOAL.md` remains user-owned.

## Earlier evidence through 0437

[0437](changes/0437-odp-bounded-plain-slides.md) implements and measures fresh
plain titled-slide ODP streaming with explicit source/sink/resource ownership.
The common fixed-prelude constructor is opt-in and preserves the old strict
contract. Full release suites, resource/refusal/sink tests, independent
semantic/package gates, and the 36-report/1,080-sample/six-profile matrix pass.
The API is kept for its measured 420,352-byte operation peak across three
sizes, 98.247% below large Builder, with the 1.574/1.588 large p50 ratios and
all other regression flags disclosed. Native visual compatibility is unproven.

This closes the current plain fresh ODP creation slice, not the program goal.
Fixed-markup execution-accounting cost was the next measured-hypothesis target;
existing append, richer creation, Part addition/repackaging, real-producer
breadth, cold/range/scaling, and other taxonomy gaps remain. The user-owned
`docs/GOAL.md` stays unchanged and excluded from commits. The batch retains
reproducible raw evidence and portable cleanup proof.

## Earlier evidence through 0436

[0436](changes/0436-odt-bounded-text-spans.md) retains measured private ODT
ordinary-text batching: 24 formal reports, 720 samples, four profiles and
1,002 passing ODT tests. Normal p50 is 19.382–30.156% lower across three sizes
and two repeats, with exact archive/sink identities and identical aligned
allocator vectors. Region peak above entry remains 420,091 bytes; RSS is
unchanged in practice. No matched comparison crosses 5%; candidate tiny p99
repeat drift of −5.489% remains visible. Formal whole-process consume-self
share falls 44.27% → 19.24%; the scope includes setup and oracle work.

The immediate ODT accounting hypothesis is now measured. Next establish ODP
fresh creation evidence. Rich authoring, logical append, Part addition,
repackaging, native compatibility breadth, cold/range I/O, scalable parallelism
and the other original completion criteria remain open. The original user goal
is unmodified, and this batch neither completes nor narrows it.

## Earlier evidence through 0435

[0435](changes/0435-odt-bounded-plain-paragraphs.md) implements bounded fresh
plaintext ODT publication and retains 36 formal reports, 1,080 samples, six
profiles, 1,445 passing ODT/common tests, and 315 passing harness tests.
The new API lowers large-case operation allocator peak from 22.45 MB to
0.42 MB and requested bytes by 86.159%, with roughly 3.35–3.37 times candidate
Builder latency. Whole-process RSS is essentially unchanged. Every regression
and repeat flag remains visible; this is a measured sequential-output enabler.
The next hypothesis targets ODT per-text Work charging, sampled at 45.74% self
in the whole-process profile, before ODP fresh creation. Logical append, Part
addition, repackaging, rich authoring, native compatibility breadth, cold/range
I/O, scaling, and the other original completion criteria remain open.
`docs/GOAL.md` is unmodified. This batch does not complete or narrow that goal.

## Earlier evidence through 0434

[0434](changes/0434-ods-bounded-text-spans.md) retains a matched ODS
streaming comparison with 24 formal reports, 720 samples, and four passing
whole-process profiles. The same selector and deterministic scalar rows are
used before and after the private ordinary-text span batching. Descriptive
normal p50 deltas (after versus before) are tiny **−12.499% / −16.211%**,
medium **−13.731% / −13.121%**, and large **−13.562% / −13.492%** for R1/R2.
No matched comparison exceeds the 5% regression trigger, including RSS; the
only repeat flag is baseline tiny p99 at **+9.844%**.

Allocator calls, requested bytes, and regional peak-above-entry vectors are identical
before and after for every shape/repeat, with a 419,347-byte regional peak
above entry. The four profiles are whole-process samples that include setup
and the untimed oracle; `ExecutionContext::consume` self samples move from
25.01% to 10.57%. Ten before and eleven after addr2line warnings remain in the
retained diagnostics, zero L1 readings are not a zero-miss proof, and LLC was
not captured. The derived `claims[]` is empty: no broad 10x, RSS, physical-copy,
total-memory, or scaling claim follows. Copied-bundle replay and six mutation probes pass before and after task cleanup.

The ODS release receipt reports 454 tests passed; scoped Clippy retains the
known common `ArchiveReaderKind` large-enum debt. This is descriptive evidence
for one fresh ODS scalar-creation path. The next implementation slice is
bounded ODT paragraph creation, then bounded ODP creation; logical append,
package-Part addition, arbitrary repackaging, native breadth, and the overall
non-iWork goal remain open.

## Prior evidence through 0433

[0433](changes/0433-ods-bounded-fresh-scalar-creation.md) records 36 formal
reports and 1,080 samples for fresh one-sheet ODS scalar creation. The new
sequential writer's normal p50 is 59.017–62.137% below the before-buffered
role across the retained shapes and repeats; the after-buffered control retains
a +7.037% tiny R1 p99 regression flag. In allocator mode, the large
operation-region peak is observed at 71,050,076 bytes for the buffered role and
419,347 bytes for streaming. These are descriptive harness observations, not
an accepted speedup, total-memory, RSS, or allocator-internal claim. The
[result bundle](results/change-0433/README.md) retains the protocol, summary,
source/build custody, profiles, and checks. Copied-bundle replay and mutation
checks pass before and after removal of the task temporaries. Fresh creation
is the only covered semantic;
logical append, package-Part addition, arbitrary repackaging, and broader
native/I/O/scaling coverage remain open.

## Prior evidence through 0432

[0432](changes/0432-xlsx-streaming-operation-memory.md) closes the missing
operation-allocation observation for the existing XLSX scalar streaming writer.
Across 64/8,192/131,072 rows, all 180 allocator samples retain the same
420,110-byte incremental peak and zero exit live-byte change. This supports
the tested creation path; it does not close richer authoring, logical append,
all-feature memory bounds, cold/remote I/O or explicit scaling. Medium normal
latency repeat drift remains flagged. The [next-work audit](results/change-0432/next-work.md)
identifies ODS bounded scalar-row creation; the [native audit](results/change-0432/native-gap.md)
records the PPTX identity/fixture gap. The non-iWork goal remains active.

## Prior evidence through 0431

[0431](changes/0431-verified-compressed-source-transfer.md) implements and
measures verified compressed source-part transfer in the source-backed PPTX/OPC
workflow. Matched synthetic media API medians improve 86.5–89.0% for bytes,
warm files and short ranges, and 19.4–20.1% for simulated delayed range. The
initial delayed-range regression and its bounded-read refinement remain
reviewable. Refined validation includes 1,755 tests, strict checks and ASAN
smoke. This closes one measured publication hotspot; representative CRUD,
broader native/size coverage, true cold/remote I/O, semantic streaming/append,
allocator/physical-copy attribution, concurrent scaling and remaining strict
gates stay open. See [batch scope](results/change-0431/goal-scope.md). The
broader non-iWork goal remains active.

## Prior evidence through 0430

[0430](changes/0430-pptx-publication-cpu-attribution.md) closes the missing
publication CPU attribution in the synthetic media-rich provider experiment.
Frame-pointer capture on the unchanged binary resolves the Deflate callers;
its measured iteration share is 83.22% / 83.38% for bytes/warm files. The
[source and transfer audit](results/change-0430/transfer-design.md) identifies
compressed-media reuse as the next optimization, subject to preserved
validation, source authority, budgets and archive ownership. This batch makes
no production change or performance-improvement claim. Normal matched
captures, transfer implementation and adversarial validation remain required;
the broader non-iWork goal remains active.

## Prior evidence through 0429

[0429](changes/0429-pptx-provider-native-baselines.md) completes the 32-process/960-sample provider and native selected-
image baseline. It adds a bounded ZIP short-read correctness fix, not a measured
speedup. All final managed gauges are zero, and non-RSS phase observations match
exactly across samples and repeats. All 27 repeat flags are RSS points already
different at entry; comparative RSS conclusions remain withheld. Files are warm,
ranges are explicit simulations, and native selected-image ownership does not
establish native cross-copy. The 100-sample CPU-profile validator correction is
explicitly amended with original validators and unchanged captures retained.
Broader producer/size and CRUD coverage, true cold I/O, allocator/physical-copy
attribution, semantic streaming and explicit scaling remain open. The global
goal is still active; see [batch scope](results/change-0429/goal-scope.md).

## Prior evidence through 0428

[0428](changes/0428-managed-pptx-cache-lifetimes.md) completes the fixed synthetic
managed-cache/budget matrix: 16 processes, 480 samples, exact admission and
refusal, pinning/eviction, oversized bypass, and three publications at an exact
cumulative output ceiling. Every final releasable caller budget is zero. It
adds no production optimization and does not close the global goal. All 27
repeat flags are RSS points with differing entry values; comparative RSS
claims require a new setup/warmup attribution protocol. Native-producer and
cold/range-source lifecycles, the full CRUD taxonomy, streaming, and explicit
scaling remain open. See the [next-work record](results/change-0428/next-work.md).

## Prior evidence through 0427

[0427](changes/0427-pptx-allocator-drop-checkpoints.md) adds explicit allocator
phase/drop observations for the current plain/media-rich PPTX APIs. Eight
fresh processes retain 240 samples, all returning to their entry callback
live-byte value after the final sink drop, with identical repeat phase changes.
This closes the narrow missing caller-drop callback observation. It does not
close the full retention/cache/managed-budget/near-limit gate or establish RSS
release, object-owned bytes or a leak finding. The
[next-work record](results/change-0427/next-work.md) identifies additive
fallible PPTX cache-diagnostic forwarding using existing managed constructors.
Composite harness coverage is 288 passes and one ignored after a baseline-
reproduced stale selector assertion is corrected; the full failed command is
retained. Existing harness strict debt remains the same 29 findings.

An earlier retained allocation optimization is
[0424](changes/0424-staged-pptx-payload-reuse.md), based on production revision
`d18bf7db4` and evidence revision `340cc91ae`. Its matched source-backed media
lifecycle records 14.266% fewer requested allocation bytes and a 7.317% lower
mean V3 operation-region peak. Logical reads and endpoint retention remain
essentially unchanged; normal timing is diagnostic. The standalone bundle
replays 16 reports, four traces and 104 mutation probes after original
worktree/binary cleanup. It establishes no release latency, physical-copy,
managed-budget or post-drop claim.

The intervening [0419](changes/0419-pptx-bounded-archive-growth.md),
[0420](changes/0420-opc-owned-payload-reuse.md),
[0421](changes/0421-allocator-peak-counter.md),
[0422](changes/0422-operation-region-allocator-peak.md), and
[0423](changes/0423-matched-source-backed-pptx-lifecycles.md) records distinguish
allocation requests, retained endpoints, corrected lifetime peaks, V3 region
peaks and aligned lifecycle boundaries. Historical high-water values affected
by the 0421 correction remain unsuitable for current comparisons. There are
nine strict registered claim replays and 15 index categories with 32
representative selectors in the retained 0424 verification; coverage remains
representative rather than complete.

The 0425 audit finds 64 workspace members: 45 non-iWork leaf/shared packages,
the separately configured `litchi` facade, 17 iWork owners, and the Python
binding that unconditionally enables iWork. Shared ZIP and XML owners remain
in scope. The current strict layout finding in ODF's `ArchiveReaderKind`
requires a measured ownership decision; its borrowed reader is retained inline.
Mechanical constant-chunk and test setup findings are closed in the 45 leaf/shared
packages, with strict Clippy passing for 44 and the ODF layout gate open.
[0425](changes/0425-non-iwork-verification-maintenance.md) retains composite
coverage of 15,778 passing tests and 98 ignored tests. The separate facade had
six baseline-reproduced test failures, now closed by
[0426](changes/0426-xls-formula-ancillary-preservation.md): five fixture/assertion
corrections and bounded standard BIFF8 Formula ancillary preservation. The
final facade suite passes 461 tests with 11 ignored; XLS passes 1,341 tests
with one ignored doctest and its scoped strict gate passes. This is correctness
evidence, with static metadata layout cost documented separately and no
performance claim. The facade strict rerun retains the same 18 findings from
the prior audit; no allowance is added.
The standalone harness has 29 lint findings outside the modified code, and
native-resave requires a lockfile refresh before its locked gate can run.
These strict-gate debts remain open; no blanket strict pass is claimed.

Resolve the remaining scoped strict-gate debt and return to the performance
evidence program with facade correctness restored.
Explicit caller-drop snapshots, near-limit memory/cache evidence, broad native
producer matrices, cold/range sources, bounded semantic streaming and append,
scaling/CPU counters, full CRUD coverage and final strict gates remain open.
The full non-iWork goal is not achieved. Sections below retain earlier
revision-specific audits and their historical counts; this section supersedes
only the current facts explicitly listed above.

Change [0418](changes/0418-pptx-cross-copy-candidate-reuse.md) retains a scoped
owned PPTX candidate-reuse optimization. Media lifecycle p50 improves
38.27–38.69%, with an explicitly reviewed approximately 8% whole-process RSS
increase. Two opt-in lifecycle selectors bring the registry to 427; the default
36-case/198-row matrix and representative index statuses remain unchanged.
The batch retains 6,400 normal observations, 480 separate allocator observations,
paired profiles and source/output checks. Retained-memory, broader corpus,
source-backed, native/cold/remote and scaling work remains open.

Change [0417](changes/0417-representative-crud-baseline.md) broadens current
descriptive evidence to all 30 representative selectors across 14 executable
categories at clean revision `b10d6c25a13242ca260a8c897946f4d80ae06c61`.
It retains 30,000 normal observations, 1,800 separate allocator observations,
30 preflight checks, explicit timer boundaries, repeat uncertainty and portable
verification. Five selectors have >5% repeat drift on at least one quantile;
28 lack operation allocation attribution. This is progress on baseline coverage,
with no speedup or completion claim.

The index's default/full-run status contract is unchanged, and several cases
time only an already-open query or commit. Aligned end-to-end workflows,
operation allocation, broad corpus/producer matrices, cold/remote, native,
failure and scaling evidence remain open. The media-rich PPTX path provides
a concrete investigation target with an existing source-backed counterpart
capability; matched timing boundaries and preservation checks are required.

Change [0415](changes/0415-zip64-streaming-deflate.md) adds explicit streamed
ZIP64 Deflate framing, automatic Office transport selection from raised limits,
and strict support for descriptor-backed zero local-size placeholders. Its
transport tests and independent large Python OPC corpus narrow the ZIP64 gap;
semantic streaming, broader append/failure matrices and the performance program
remain open.

Change [0416](changes/0416-local-zip64-read-preservation.md) adds local-only
forced-ZIP64 read/preserve compatibility: 13 deterministic Python fixtures, 10
focused ZIP integration cases, and 7 OPC no-op/edit/failure cases, with strict
path fuzz coverage. The fixtures cover signed and unsigned descriptors,
descriptor-signature CRC collision, seekable no-descriptor output, empty and
many-small members, and central/local or ZIP32-tail combinations. This records
capability coverage only; the source identity, performance guards, and final
gate status belong to the 0416 change record. The overall non-iWork goal
remains open.

The prepared ODF catalog's existing `has_zip64_metadata` observation is based
on central-directory/tail metadata. It does not classify a seekable local-header
ZIP64 sentinel when the central and tail metadata remain ordinary. ODF catalog
parity for that input form, together with semantic, cold/remote, native,
concurrency/scaling, and broader failure-matrix evidence, remains open.

Change [0414](changes/0414-zip64-output-promotion.md) implements preservation
output promotion at ZIP32 local-offset and member-count boundaries, including
OPC provenance for reopened packages above 65,535 members. Generated ZIP64
Deflate through that revision's streaming local-header path remained a typed
refusal; 0415 replaces that refusal with explicit framing. These capability
changes do not close the performance program.

Change [0413](changes/0413-cfb-chain-scratch-reservation.md) retains a scoped
production CFB optimization with nine XLS ABBA comparisons, exact allocation
and logical-I/O checks, a reviewed ~3% CFB guard cost, paired CPU/PMU evidence
and eight strict registered claims. This is verified progress, not completion
of the non-iWork program. Broad scenario, native/cold/remote and scaling gaps
below remain open.

**Audit date:** 2026-09-05
**Audit basis:** the 0418 candidate `f8f9e6667` against control `79dfee502`,
the 0417 representative baseline at `b10d6c25a` and retained evidence,
with the 0416 local-framing candidate `d18cd04a2` and retained evidence,
with the 0415 streaming source batch and its retained evidence,
with the 0414 source batch and its retained verification evidence,
with the 0413 committed control `6b632726b` and the 0412 captured candidate at
`63c95bc22d5883c8ecab0872030757e5584254f7`, with the verified 0411 baseline at
`44edf790669a0aa4dc0aff73af6f7b5f5e709b6d` and the earlier
`9c6742c5212dd0e7ff2367da585abe357aae8975` ZIP64 control retained for the
historical comparison. iWork is outside this audit by the user's instruction.

**Disposition: OPEN.** The repository has substantial correctness and scoped
performance work, but the definition of done in `docs/GOAL.md` is not met. The
coverage index explicitly describes itself as representative, and the current
records do not provide a complete, independently reproducible baseline and
optimization result for the non-iWork CRUD matrix.

## 0412 current implementation

Change [0412](changes/0412-xls-observer-isolation.md) at committed revision
`b8f61970d` corrects the XLS observer's category range catalog by coalescing
only exactly adjacent spans, preserving overlap multiplicity, and disabling
the unused generic repeated-read union. It adds three explicit opt-in plain
`OwnedSource` lifecycle selectors: `xls_owned_source_open`,
`xls_owned_source_open_list_worksheets`, and `xls_owned_source_open_one_cell`.
Their timed reports expose operation/allocation metrics with `source: None`;
separate instrumented observations retain the logical locality evidence.

Eleven focused XLS observer/owned/allocator tests plus the registry test pass,
along with scoped formatting, crate boundaries, coverage-index validation, the
seven-claim strict checker, and report classification. The clean candidate
`63c95bc22` now has 18 schema/corpus-verified timing, allocation, profile and PMU
reports. Plain one-cell p50 is 0.166–0.169 ms; separate allocator captures record
126 calls / 223,774 bytes. Instrumented locality remains unchanged across the
observer correction. All 91 comparator tests pass, including mismatched observer
rejection. This is a measurement enabler and baseline, with no production
speedup claim. The registry is now 425 selectors; the default remains 36 cases / 198 rows and the
index remains 15 categories / 30 representative selectors. The non-iWork goal
remains open.

## 0411 current update

Change [0411](changes/0411-xls-read-allocation-baseline.md) now retains a
verified, descriptive baseline for six explicit opt-in XLS/CFB lifecycle
selectors: eager and source-backed open, open plus worksheet listing, and open
plus one-cell selection. The fixed generated corpus is
`xls-comments-opaque-heavy` (`litchi-xls-comments-opaque-heavy-v1`): two sheets
(`Comments`, `Untouched`), selected `Untouched!E21 = 42.0`, 257 logical entries,
10 archive members, 16,995,840 archive bytes, an 80,946-byte `Workbook` stream,
and eight 2 MiB opaque streams. The eager and source-backed paths use the same
two-sheet semantic oracle; this does not assert parity with the separate real
producer corpus discussed in older records.

The [0411 evidence bundle](results/change-0411/) contains four fresh CPU-2,
one-worker normal processes with 20 warmups and 500 samples per selector, plus
two allocator processes with 3 warmups and 30 samples per selector. Samples
share each process; the protocol is warm in-memory corpus evidence, not cold,
remote, concurrent, or per-sample process-isolated evidence. Source-backed
reports retain classified `ReadAt` counters and prove zero worksheet/opaque
payload reads for open/list and selected-only worksheet reads for one-cell.
Allocator reports retain operation-scoped allocation vectors; their peak values
are process-lifetime snapshots, and neither allocation nor timing is an A/B
speedup claim.

The capture and corpus verifier passed on clean revision
`44edf790669a0aa4dc0aff73af6f7b5f5e709b6d`. All-feature, all-target
`litchi-xlsx` Clippy passed; the full XLSX test run passed 1,238 tests with no
failures, and the combined harness/allocator test selection passed 9 tests with
four test threads. This evidence expands the current descriptive baseline but
does not promote the selectors into the default matrix or claim completion of
the non-iWork goal. The index remains 15 categories and 30 representative
selectors; the default remains 36 cases / 198 rows and the selector registry
remains 422.

## 0410 current update

Change [0410](changes/0410-mce-attribute-name-reuse.md) is a bounded MCE
expanded-attribute ownership experiment on the XLSX selected-cell path. It
uses candidate `e4d477466718a8fad38cd55b9babe0b826e7f3a7` and control
`972dc25be0dbd6690c74429839a48288d637e2d5`, the fixed four-sheet 9,216-cell,
17-member corpus with 4 MiB of media, and Rust/Cargo/Rustdoc 1.98.1 release
builds on CPU 2. The primary ABBA p50 candidate-minus-control changes are
`-3.9447%` and `-4.1373%`; operation-local allocator calls move from 81,918
to 77,212 and allocated bytes from 10,690,444 to 10,309,094. The seventh
strict-registry claim entry has been added, and the strict checker passes all
seven claims.

The initial eager edit/save guard is adverse at +0.53% / +5.66% p50 elapsed
time. A diagnostic repeat is lower by 1.106% / 0.505%, while same-role p50
drift is roughly 5–6% for both roles. The edit claim is therefore withheld and
the initial adverse result retained. Review found that the eager fixture does
not execute the changed MCE stream; this limits causal interpretation and does
not prove an eager no-regression result. The residual selected-path profile
attributes 10.91% of leaf weight to `clone_bounded_name_part` and reports
`parse_element` at 5.89% self / 23.42% inclusive. These are sampled profile
figures, not paired CPU evidence.

The final all-feature run passes 1,918 tests across the three crates. The
OPC/common warning-denied Clippy passes after two preexisting test-lint fixes;
XLSX all-target Clippy remains blocked by 9 library diagnostics and 28
library-test diagnostics, including 19 additional test diagnostics. This
update records scoped progress only; the non-iWork goal remains open.
The [0410 evidence bundle](results/change-0410/) retains the build identities,
ABBA reports, allocator captures, profile attribution, and gate logs. Rustdoc,
crate-boundary, and scoped-format checks pass.

## Authorities and reading boundary

This audit uses the current goal, the accepted ADRs, and the current
non-iWork performance records:

- [`docs/GOAL.md`](../GOAL.md), especially the mission, measurement contract,
  CRUD matrix, deliverables, and definition of done.
- ADR 0001 (public layers and typed refusals), ADR 0002 (downward crate
  ownership), ADR 0003 (immutable snapshots and atomic edits), ADR 0005 (the
  `ReadAt`/budget/output/measurement contract), ADR 0006 (preserve-by-default
  validation and security), ADR 0008 (migration gates), ADR 0010/0011 (facade
  and OPC physical ownership), and ADR 0024 (current topology).
- [`REPORT.md`](REPORT.md), [`CRUD_COVERAGE.md`](CRUD_COVERAGE.md),
  [`crud-coverage-index-v1.json`](crud-coverage-index-v1.json), and the
  claim-registry policy.

The decisive constraints for the open ZIP work are: preserve is the default;
validation must not mutate; an owned changed OPC source may publish only
through a proven preservation plan; unsupported framing must return a typed
refusal before output; and physical ZIP ownership stays in `soapberry-zip`.
Those constraints come from ADR 0005's exact-source amendment, ADR 0006, and
ADRs 0010/0011.

## Current evidence through 0502

| Goal requirement | Current evidence | Assessment |
| --- | --- | --- |
| Reproducible CRUD baseline | The current refresh covers 15 categories, 33 selector-backed mappings, 37 cases and 31 corpora with two serial 201-row lanes and 6,030 samples; generated report/catalog validators and 38 static coverage tests pass. | Timing-report baseline gate is closed for the 11 measured mappings (48 rows); 22 mappings remain correctness-only. The complete checklist, missing metrics, native producers, and correctness-only rows remain open. |
| Scoped claims | The strict registry currently validates 10 claims. 0501 retains 208 comparison rows and 52 favorable flags; 0502 retains a matched timing summary and an explicit regression review. | Claims are scoped to named workflows and evidence; they do not establish program completion. |
| Correctness and boundaries | Recent bundles retain source/build manifests, exact output/source oracles, budget/cancellation tests, preservation checks, and ADR matrices; 0495–0497 cover managed edit and an atomic logical-tail capability. | Strong scoped DOCX/OPC and ODG correctness/custody evidence; general atomic save, all formats, and full CRUD remain unproven. |
| Provider and cache states | 0491 and 0494 cover owned, file, instrumented, short, simulated delayed/range, and verified-cold DOCX lanes; 0492/0493 measure and integrate bounded read-ahead. | Descriptive and opt-in evidence only. Genuine borrowed lifetimes, physical-device cold behavior, native producers, and cross-format intersections remain open. |
| Hardware/resource profiling | 0490 sync traces, 0496 phase clocks, 0501 whole-child profiles, 0502 heaptrack/layout observations, and allocator/RSS reports are retained with scope limits. Hardware counters are unavailable for 0502. | Attribution is partial: whole-child/setup-inclusive evidence is not operation-local CPU, allocation, or RSS proof; lock-wait and full scaling evidence remain open. |
| Parallelism and batching | 0498/0499 provide explicit bounded ordered Part batches, and 0500 measures managed paragraph batching; adverse local rows and serial controls remain visible. | Partial capability evidence. Full lifecycle 1/2/4/8 scaling, serial-fraction/Amdahl analysis, and stable lock-contention evidence remain open. |

The retained 0486–0502 evidence establishes several bounded capabilities and
descriptive baselines, including managed edit ownership, atomic logical-tail
publication, source read-ahead, Part batching, and PPTX/ODG scoped experiments.
It still does not establish the `docs/GOAL.md` requirements for complete
p50/p95/p99 coverage, throughput, allocations, peak RSS,
copied/decompressed/recompressed bytes, physical I/O, lock wait, borrowed
inputs, native producers, all cold/warm intersections, or Amdahl scaling. The
0502 open regressions, 0499/0500 local adverse rows, and 0501 missing replay
binary remain explicit rather than being absorbed into a completion claim.

## ZIP64 passthrough audit

The low-level foundation is now materially further along than the last
committed control. `soapberry-zip` retains ZIP64 field origins and tail bytes;
its preservation plan preflights the complete output, patches only movable
offset fields, retains central/local metadata and ZIP64 extensible data, and
has malformed, multi-disk, truncation, descriptor, limit, offset, sink, and
no-output tests.

The earlier integration changed public preservation construction to
`AllowZip64` and admitted already-ZIP64 sources. Change 0414 also promotes
generated/copied central offsets and synthesizes ZIP64 tails for ZIP32 sources.
OPC size/count capability guards and the provenance count ceiling are removed;
arithmetic, limits and structural refusals remain. Earlier synthetic tests
continue to cover:

- `zip64_source_targeted_save_preserves_untouched_records_and_tail` changes
  one Part in a ZIP64 package, retains a different Part's raw local and central
  records, retains ZIP64 tail extensible bytes and the comment, reopens the
  output, and compares `to_bytes` with `write_to_stream`.
- `projected_zip64_descriptor_targeted_save_preserves_untouched_member`
  changes one Part while a different member uses ZIP64 data-descriptor fields,
  then checks raw preservation, reopen, and stream output.

Focused source-backed coverage now also includes
`zip64_one_part_overlay_preserves_raw_unknown_members_and_cold_work` and
`topology_zip64_add_remove_preserves_untouched_records_and_tail`. Both pass in the final integrated test run. The
package-writer suite also has a ZIP64 partial-sink failure case.

These tests are meaningful end-to-end package-writer correctness evidence, but
they are synthetic in-memory fixtures. They leave the following gates open:

1. Extend the source-backed selected-Part and topology matrix to signed/encrypted inputs,
   both descriptor forms, non-seek sinks, cancellation, configured metadata
   limits, and atomic filesystem finalization. Each refusal must leave the
   sink/destination untouched where the API promises that property.
2. Exercise ZIP64 add/remove/topology plans with more than the current narrow
   synthetic graph, including unknown physical members and dependency closures.
   The existing focused add/remove test is valuable correctness evidence but
   does not certify every topology operation.
3. Complete generated-size and streaming creation coverage. Change 0414
   implements offset/count promotion and known-size Store/precompressed ZIP64
   headers. Public sparse tests cover generated offsets immediately below, at
   and above `u32::MAX`; OPC tests cover count promotion and repeated owned
   publication beyond 65,535 members. Copied-offset promotion has focused
   layout/metadata tests. Change 0415 adds explicit one-pass streamed Deflate
   framing and raised-limit Office transport admission, including generated
   metadata charging and descriptor-backed local placeholders. Preservation
   regeneration still buffers payloads; transport-level bounded-memory evidence
   does not certify semantic row/paragraph/slide creation or append.
4. Add at least one real-producer or independently generated large ZIP64 OPC
   corpus and validate all semantic Part bytes, raw member identity, archive
   layout, and reopen behavior. Change 0415 supplies an independently generated
   Python OPC source with a real 4 GiB logical member and a small XML overlay
   preservation check. The 0416 batch adds deterministic Python forced-ZIP64
   local-framing fixtures and focused ZIP/OPC no-op, edit, and refusal cases;
   these remain synthetic compatibility evidence, and the final gate status is
   deferred to its change record. Native Office producer coverage, ODF
   prepared-catalog classification of local-only sentinels, and the broader
   dependency/topology matrix remain open.
5. Expand the scoped gates in
   [`changes/0404-zip64-preservation-integration.md`](changes/0404-zip64-preservation-integration.md)
   to the full non-iWork feature and native-producer matrix. The scoped final
   commands, source bindings, and pass/fail logs are now retained.

The control behavior is also important: at the control revision, a changed
owned ZIP64 source was refused rather than normalized. That refusal remains the
safe fallback for unsupported framing, suffixes, opaque topology, and
unrepresentable generated output. A successful targeted test does not authorize
a normalizing fallback for a different unsupported source.

## Prioritized remaining work

OLE2 and OOXML take precedence until their optimization goal is complete.
ODF work, including the retained ODG regression investigation, is deferred.
The priorities below apply within that ordering; historical ODF evidence
remains retained without scheduling it ahead of OLE2/OOXML.

| Priority | Requirement from `docs/GOAL.md` | Next reviewable evidence |
| --- | --- | --- |
| Deferred ODF | Resolve the 0502 ODG open regressions | Profile parser and per-shape metadata work with operation-local attribution, then capture a same-protocol before/after pair for plain and metadata corpora. Keep the boxed-layout/heaptrack result separate from latency and allocation-count claims. |
| P0 | Restore replayable 0501 custody or recapture it | The retained 0501 reports and historical receipt are usable, but `/tmp/litchi-goal-0501` no longer contains the frozen before executable. Rebuild from the recorded source/build manifest or retain a new verified replay bundle before claiming current independent replay. |
| P0 | Complete the Phase-1 metric and CRUD baseline | The 6,030-sample default refresh closes the timing-report gate for 11 measured mappings (48 rows) within the 33-selector index, but the 15-category checklist still has correctness-only/unsupported rows and lacks complete throughput, copy/decompression, lock-wait, cold/warm, and scaling coverage. Promote only rows with validated timing evidence. |
| P0 | Measure provider and publication intersections | Extend the 0491–0494 DOCX provider/cold baselines and 0497 atomic logical-tail capability to genuine borrowed input, sequential non-seek sinks, filesystem atomic save, source-version/cancellation failures, and physical cold behavior across representative formats. |
| P1 | Address the retained local batch and lifecycle adverse rows | Revisit 0499/0500 serial-versus-batch choices and their flagged latency/RSS rows with operation-local CPU/allocation evidence. Publish full explicit 1/2/4/8-worker curves, lock-wait proxy, and serial-fraction/Amdahl analysis before widening concurrency. |
| P1 | Finish source-backed CRUD adoption across formats | Extend selective open/read/edit/save and dependency-closure publication through DOCX, XLSX, PPTX, XLSB, DOC, XLS, and PPT owners, followed by ODF after the OLE2/OOXML optimization goal completes; measure physical I/O, decompression/recompression, copies, allocations, RSS, and semantic phase boundaries. |
| P1 | Cover the high-impact CRUD categories and real producers | Add or explicitly classify conversion, creation, append variants, structural/deletion/sanitization, cross-document copy, merge/split, patch/inverse/three-way merge, repair/normalize, dynamic content, security, malformed, signed/encrypted/macro-enabled, and independently produced Office corpora. |
| P1 | Close ZIP64/CFB and output-source preservation intersections | Exercise ZIP64 and CFB topology/stream cases with non-seek sinks, cancellation, configured limits, physical cold/high-latency sources, unchanged-member passthrough, and atomic finalization while retaining typed refusals and exact untouched bytes. |
| P2 | Apply layout, cache, or SIMD tuning only from measured hot loops | Require operation-local profiles, scalar fallbacks, differential malformed-input tests, and material end-to-end benefit before any low-level change. |

No row in this audit should be read as a completion claim. The current evidence
supports targeted capability progress, including the 0491–0494 provider/cold
baselines, 0495–0497 managed publication capabilities, 0498–0500 bounded batch
work, the 0501 PPTX comparison, and the 0502 ODG audit. The retained
regressions, missing replay executable, incomplete CRUD metrics and provider
intersections, and unmeasured scaling keep the full non-iWork performance goal
open.

## 0546 — retained XLSX shared traversal

Retained exact-admission shared traversal: planning p50 improves 24.77–26.51%, workflow p50 7.09–10.80%, and planning Ir 21.81–22.36% on both primary shapes/repeats. All frozen gates and nine final quality commands pass (1,306 tests). Same-invalid refusal and clustered scanner regressions remain explicitly reviewed; no universal speedup, RSS, or cold-cache claim is made. Accepted ADR/index hashes remain unchanged. OLE2/OOXML optimization remains active; ODF is deferred until completion and iWork excluded. [Change and limitations](changes/0546-xlsx-shared-traversal-retained.md); [sealed evidence](results/change-0546/README.md).

## 0547 — OLE2 collector sub-operation attribution

Fresh baseline profiles attribute visited-map per-step checks/update to 20.37% of XLS-owned and 22.35% of few-large CFB constructor Ir (tiny 0.69%, many-small 2.05%). A 376,264-case semantic model validates terminal proof and exact-error replay, but exposes declared-length short-cycle amplification in the terminal-only design. A bounded checkpoint/replay candidate is the next measured hypothesis; no production change or speedup is claimed. Exact-source 0546 quality evidence is reused, not rerun. [Detailed boundaries and next gates](changes/0547-ole2-collector-attribution.md). OLE2/OOXML remain active; ODF is deferred until completion and iWork excluded.
