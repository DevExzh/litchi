# Change 0635 log paragraphs

The coordinator merges these into the four repository logs. Written in the style
of each log's newest section.

## For `HOTSPOTS.md`

## 0635 — the XLSX fact builder, the stylesheet count, and a chain that nothing reaches

**The hotspot change 0622 named is reduced (change 0635).** 0622 closed by pointing
at its own successor: "with the second scan gone, the planning traversal is the whole
cost of an XLSX value commit, and the fact builder is 15-18% of it; the next
reduction on this route is the builder's per-cell `r` parse and per-element ampersand
probe, not the writer." Both are gone. The `<c r="…">` and `<row r="…">` values are
now parsed in one pass against the row the builder already holds, in place of an
alphabet sweep, a UTF-8 validation and `parse_a1`/`parse_one_based_row` with their
`format!` diagnostics; the ampersand question is answered once per part with one
`memchr` instead of once per start tag; and two name comparisons that compared a
value with itself are gone. Measured on `taskset -c 11`, callgrind isolation pairs
(N=1, N=11): the builder falls **35.0-36.2%** on every one of eight scenarios (`FactsBuilder::element` −41.8 to −43.4%, `FactsBuilder::cell` −36.9 to −37.7%, `raw_attribute` −24.6 to −27.4%), planning falls **4.29-5.26%**, and the whole harness iteration falls **0.13-2.19%** — the largest on the 0601 producer-shaped edit selectors, which are not dominated by the cell-CRUD corpora's 4 MB of media. The builder's share of planning goes from 13.2-14.3% to 8.9-9.7%. No scenario got worse. The builder's decline set is unchanged — change
0622's corpus funnel is identical at 391 worksheet parts, 207 admitted, 1 accepted, 1
publishing facts — and two new tests compare the fused parsers against the shared
ones over 484 byte strings. **XLSX-5 of the 0587 survey is answered**:
`raw::styles::parse` already retained only a count, so what the survey measured was
the traversal, and that traversal copied every event twice (`Event::into_owned` plus
a `NamespaceResolver` clone) and discarded both copies. Removing them takes exactly
**80 allocation calls and 6,457 allocated bytes out of every planning**, on every
case and shape, with the same events, checks and messages; on the 603-byte harness
stylesheet the instruction saving is inside the symbol's own variation. **The
remaining XLSX hotspot on this route is now the shared traversal itself**, not the
builder: the fused traversal `worksheet_xml_and_parse_source` is 96.7-97.4% of
planning after this change, and the builder is now 8.9-9.7% of it rather than
13.2-14.3%. **One hotspot is closed by
measurement rather than by code**: the "snapshot chains drop facts after the first
commit" item is *not* worth closing — see `GOAL_AUDIT.md`.

## For `GOAL_AUDIT.md`

## 0635 — queue item 17 and survey item XLSX-5 closed, one by code and one by evidence

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

## For `REPORT.md`

## 0635 — a smaller note-taker on the XLSX edit path, and a shortcut that leads nowhere

`litchi-xlsx`: the note the source-backed value editor takes while it reads a
worksheet — the sixteen bytes per cell that let change 0622 delete the commit's
second read — now costs about a third less to take. Reading a cell's address used to
sweep its `r="B7"` for legal characters, validate it as UTF-8, parse the column
letters and the row digits, and then check the row against the one already known;
it is now one pass that stops at the first disagreement. The question "does this tag
contain an ampersand?" used to be asked of every tag and is now asked once of the
whole worksheet. Separately, counting the workbook's cell formats — which the editor
does on every plan, and which needs only a number — stopped copying every XML event
twice on its way past: exactly 80 heap allocations and 6.5 KB leave every plan.
Output is byte-identical: 34 of 34 published packages hash the same on both legs, and
the differential oracle that compares the two commit routes over every worksheet part
of every `.xlsx` in the repository reports exactly the numbers it reported before.
A third idea was built and thrown away: carrying the note forward so a *second* edit
on the same worksheet would not have to re-read it. It works, and it is useless —
nothing in the public API can start a second edit from the result of the first, so
the note would be taken on every save and read on none. The measurement said so — it costs
0.2% to 4.1% of a save — and the code was reverted; it is kept as a patch for the day
that changes.

## For `ADR_COMPLIANCE.md`

## 0635 — compliant; two contract questions raised and both answered by declining

Change 0635 (`litchi-xlsx`: fact-builder reduction, stylesheet-count traversal) is
compliant and adds no ADR question. **ADR 0003 (bounded resources):** no retained
state changed — `change_0622_retained_records_stay_compact` still pins 16 bytes per
cell and 32 per row — and the styles traversal now allocates strictly less (−80
calls, −6,457 bytes per planning). No new `unsafe`, no new dependency, no weakened
limit; `MAX_XML_DEPTH`, `MAX_XML_EVENTS`, `MAX_CELL_FORMATS` and the 256-action
commit cap are untouched. **ADR 0005 (validation placement):** no validation moved,
none was added and none was removed. Every replaced test is provably the same test —
the fused reference parsers are compared against `parse_a1` and `parse_one_based_row`
over 484 byte strings covering every branch of both; the whole-part ampersand probe
is the same predicate as the per-tag probe; the two dropped name comparisons compared
a value with itself. The styles parser reads the same events with the same reader
configuration and raises the same refusals with the same messages in the same order.
**ADR 0006 (lossless preservation):** output bytes are unchanged, proved by 34 of 34
published-package hashes and by change 0622's oracle reporting an identical funnel
and an identical comparison count. **Contract movement: none — and two proposals to
move one were declined.** Substituting `raw::styles::stream_count` for
`raw::styles::parse` would move refusals (trailing text, DTDs, processing
instructions, custom entities, a late XML declaration, unbound prefixes, depth, event
count, several stream-only bounds) and would move the error type and message as well;
it is frozen as a design. Carrying facts across a snapshot chain moves no contract at
all, and was rejected on cost against an unreachable benefit rather than on
compliance.
