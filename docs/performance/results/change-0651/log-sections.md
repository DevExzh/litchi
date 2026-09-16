# Log sections for change 0651

The four paragraphs the coordinator inserted at the top of `HOTSPOTS.md`,
`GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`.

## For `HOTSPOTS.md`

## 0651 — the second wave closed: every measurable row landed, every contract row froze with a price, and the codec row moved to the top

Record: [0651](0651-queue-refresh-after-the-second-wave.md). Nineteen records,
0632 to 0650, worked the queue change 0630 left. Ten landed in production code,
all value-identical and each measured beside its own floor: the XLS whole-sheet
walk reads its string table once (16,105 → 37 positional reads on `54016.xls`,
`all-cells` over a file source −80.1% at p50, 0648) after the validation walk
took a bounded cursor window (82,727 → 60 reads, 0636); a selected-cell XLS
query validates its non-target cells without building them (−28% to −41%
cycles, 0641); the XLS commit's second complete parse is gone (commit −38.5%
at p50, 0633); the XLSX range visitor stops building the range (a warm
whole-sheet visit allocates nothing, 6.1× at p50, 0642) and its fact builder
lost a third (−35% Ir, 0635); the DOCX sink text parser borrows its events
(−88.8% allocations, −15.6% at p50, 0643); the ZIP locator reads the central
directory once (one request and a 64 KiB scratch less on every one of 533
containers, 0632); the eager PPTX slide catalog is parsed once per borrowed
presentation (a 200-slide by-index walk −95% Ir, 0637); the PPT per-slide
re-parse spans the retained stream (−5% Ir, 0634); and a relationship call that
establishes nothing keeps its pristine proof (0647). Three designs froze with
their price and their gates: the OLE2 snapshot fence (four of six complete
reads load-bearing, the adopted variant modelled at −13% of the DOC open, 0644),
the memoized PPTX revision proof (−26% of an opened lifecycle's instructions
against a durable format bump, 0645) and the cross-copy candidate retention
(apply phase −72% against a retained whole package, 0646). Three correctness
findings were fixed (0639, 0640, 0650) and one attribution closed 0592's open
regression (0643). The finding that reorders the queue is 0649's: the 133.61 ms
shape-text edit 0638 measured on a real deck is six whole-slide MCE rewrites
amplified 16× by the namespace emission 0588 withdrew, and every prior PPTX
record measured on corpora that never exercised that branch. The "Ranked work
queue" section below is replaced by the twenty-row table of this record: rows
1 to 10 each wait on an owner or an ADR; rows 11 to 20 are the follow-ons and
gaps the wave named.

## For `GOAL_AUDIT.md`

## 0651 — the second wave against the goal: unnecessary work removed on ten paths, three contracts priced instead of moved, three refusals corrected

Record: [0651](0651-queue-refresh-after-the-second-wave.md). Against the
optimization order of `docs/GOAL.md`, every landed change of the wave sits in
the first tier: unnecessary reads (0632, 0636, 0648), unnecessary parsing (0633,
0634, 0637, 0641), unnecessary allocation and copying (0635, 0642, 0643) and an
unnecessarily discarded proof (0647); no change in the wave used parallelism,
SIMD, a weakened limit or `unsafe`. Against the definition of done, 0638 gives
the facade's `.doc` and `.ppt` routes and the documented ordinary OOXML save
their first selectors and finds the save publication-bound, which bounds what
any compression work can be worth on that route. Against "refusals must be
correct, not merely conservative", 0640 removes two refusals that read a
redundant flag bit as structure and confirms five, 0650 fixes an editor refusal
caused by a three-byte offset error and freezes the four refusals it was hiding,
and 0639 closes the demo route that could not succeed. Against "never trade a
typed refusal for a partial result", 0642 declined to make the scanner yield and
0644, 0645 and 0646 stop at frozen designs where the saving would move a fence,
a durable format or a budget. The wave's method finding is 0648's: a count of
removed reads is not a saving until bytes are counted, the second time this
programme has had to say that a removed count trades into something (0612 said
it of time). The queue is refreshed; its first ten rows are decisions.

## For `REPORT.md`

## 0651 — what the second wave delivered, in the numbers of the records that measured it

Record: [0651](0651-queue-refresh-after-the-second-wave.md). This is a
coordination record: no measurement of its own, `performance_claim: none`, and
no claim-registry entry for any record of the wave. What a reader can take from
the wave, each figure scoped as its record scopes it: an XLS whole-sheet walk
over a file source is 5× faster (0648) and a selected-cell query 1.4× to 1.7×
(0641); an XLS source-backed numeric commit is 1.6× faster (0633); a warm XLSX
whole-sheet visit is 6× faster and allocation-free (0642); an XLSX one-edit
plan's fact builder is a third cheaper (0635); a DOCX text export to a sink is
15% faster with a tenth of its allocations (0643); every OOXML open over a
range source costs one request less (0632); a 200-slide PPTX by-index walk costs
a twentieth of its instructions (0637). What a reader should not take: any of
the three frozen designs' savings (0644 modelled, 0645 and 0646 measured on
scratch code that is not on the branch), the documented save's 133.61 ms real
deck edit as fixed (0649 attributed it to the codec row, which waits on an
owner), or any cold-cache, physical-device, peak-RSS or cross-platform result,
none of which the wave measured. The integration gate on the merged head and
the merge log are in the packet; the rebase onto the upstream branch that
follows this record is recorded there with a commit map.

## For `ADR_COMPLIANCE.md`

## 0651 — no boundary moved by any landed change; three designs name the boundary they would move

Record: [0651](0651-queue-refresh-after-the-second-wave.md). Every production
change of the wave is value-identical under the accepted ADRs and says so with a
corpus differential (0632's 22,875 inputs, 0633, 0634, 0635, 0636, 0637, 0641's
67,422 cell records, 0642's 397 worksheets, 0643's 333 documents, 0647's 6,281
publications, 0648's 252 rows, 0650's 63 fixtures, 0640's 57). The three designs
each state the boundary they would move and stop there: 0644 an ADR 0006 fence
(and refuses the `SourceVersion` substitution with a runnable witness), 0645 an
ADR 0003 durable patch format plus an ADR 0005 question about a facade-carried
memo, 0646 an ADR 0005 budget for a retained package. 0642 declined an ADR 0006
refusal-order change on its own; 0647 found that the ADR 0006 byte question
0628 froze does not arise; 0650's fix is ADR 0006's preserve-by-default applied
to three bytes. Proposed ADRs 0030 and 0031 were cited by no record as
authority; their rows stay blocked on acceptance. The coordinator's own changes
are documentation: this record, the four log insertions per merged record, and
the queue replacement in `HOTSPOTS.md`.
