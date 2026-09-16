# Log paragraphs for change 0641

Four blocks, one per merged log. Each is written to be prepended as the newest
section of its file.

## For `HOTSPOTS.md`

## 0641 — the XLS query hotspot is the cell record itself, and the corpus mix says which part

Change 0587's item XLS-5 named `drop_in_place<CellRecord>` at 2.05% of the
`54016.xls` one-cell query and left the corpus `Formula`/`Label` mix as
**unknown**, because change 0584's census counted `LabelSst` only. The census is
now taken — 126 fixtures, 372 worksheet substreams, 106,689 records, **127,072
cell values** — and it moves the hotspot twice. First, `54016.xls` carries
**zero `Formula` and zero `Label` records**: its 37,929 records are `LabelSst`,
`Blank`, `MulBlank` and `RK`, so the 2.05% the survey priced is the 88-byte
enum's *drop glue and moves*, not heap frees, and a measure-only path limited to
the two allocating kinds would have saved nothing on the fixture that motivated
it. Second, the allocating kinds are concentrated somewhere else entirely:
`Label` exists **6 times in the whole corpus** (42 bytes, two fixtures), while
`Formula` is 14,854 records carrying 224,754 `Rgce` bytes, of which
**`15228.xls` alone carries 175,054** across 11,240 records — 73.8% of that
fixture's worksheet records, a formula density no fixture in the program's
measured set had. Routing all seven single-cell kinds through a measure-only
instantiation takes `CellRecord::parse` from 893,405 Ir to **40** on the
`54016.xls` query and `drop_in_place<CellRecord>` from 662,150 to 158,831; on
`15228.xls` it takes `parse_record_with` to **0** and `malloc` plus `free` from
234,890 to 17,551. Natively that is **−27.99%** and **−36.73%** cycles on the two
one-cell queries and −31.70% / −41.00% on the two-query scenario, against an A/A
floor of at most 2.24%. What remains on this axis is `MulRk` and `MulBlank`:
**60,575 of the corpus's 127,072 cell values, 47.7%**, each still built and
dropped as an 88-byte record by a query that does not want it — no heap traffic,
only the move and the drop glue. That is the named next item.
[Change and limitations](0641-xls-scan-measure-only-cells-and-sheet-hint.md);
[evidence](results/change-0641/README.md).

## For `GOAL_AUDIT.md`

## 0641 — two survey items closed, one design frozen, and a prior record's claim retired

`docs/GOAL.md` ranks eliminating unnecessary work first, and both halves of this
change are that: validation that must happen still happens, and only the
materialization that nothing reads is removed. Items **XLS-5** and **XLS-8** of
change 0587 move from open to implemented. XLS-5's precondition — "the corpus
`Formula`/`Label` record mix is unknown", one of the survey's named measurement
blockers — is retired with a retained census that reproduces the survey's own
37,929-record figure for `54016.xls` exactly. XLS-8 closes the per-sheet cursor
construction term change 0585 recorded as "left open", and closes it with the
second `StreamChainHint` that record said was required. It also retires a claim
0585 made and could not measure: that one hint serving both the shared-string
resolver and the worksheet cursor would "save nothing on either path". A third
instrumented leg now prices that arrangement, and the claim holds in direction
and was overstated in degree — sharing is **1.61% worse** than the disjoint pair
over the 96 multi-sheet fixtures and worse on 16 of them, while still better than
no worksheet hint at all. One thing is deliberately **not** done: the brief asked
for the hint to live in the snapshot so that repeated queries resume it, and that
is frozen as a design rather than implemented, because `StreamChainHint` borrows
the reader it came from and every way to store one in an immutable snapshot
either weakens the documented "it cannot cross readers" property or puts a lock
on the selected-cell path of a `Clone`, cross-thread snapshot — the ADR 0005
clean-value-cache gate XLS-3 is already behind. No coverage row changes status:
reads, bytes, `version()` calls and the harness observation are identical on all
16 measured scenarios. OLE2 and OOXML remain the active priority; ODF stays
deferred and iWork excluded.
[Change and limitations](0641-xls-scan-measure-only-cells-and-sheet-hint.md);
[evidence](results/change-0641/README.md).

## For `REPORT.md`

## 0641 — measure-only cell validation for a query's non-target cells, and a second chain hint

`CellRecord::measure` runs every check `CellRecord::parse` runs and builds
nothing, returning a six-byte `MeasuredCell` in place of an 88-byte enum that may
own a `String` and a `Vec<u8>`; the frame loop chooses between them from the
`rw`/`col` header both parses read identically, and the whole-sheet walk, which
wants every cell, never reaches the measure path. It is not a second parser: the
five fixed-width kinds share one `cell_head` length check, `Label` shares one
`string_record_parts` framing, and `Formula` shares one `frame_record` and
`check_token_stream`, so the two paths cannot drift. Alongside it,
`WorksheetScan::new` takes a `StreamChainHint` and a document-wide text
extraction carries one worksheet-region position across its sheets, disjoint from
the shared-string resolver's. **Value identity:** 67,422 cell records of 124
fixtures compared through both paths, including 142 refusals — **0 divergences**
— plus 32,884 truncation and bit-flip mutations reaching 9,240 refusals with 0
divergences, and identical text digests on all 119 fixtures of the chain-link
sweep. **Counts:** reads, bytes and observations identical on 16 scenarios;
text-extraction chain links fall **6.14%** over the 96 multi-sheet fixtures
(`15228.xls` −65.5%, `ConditionalFormattingSamples.xls` −10.4%) and **0.00%**
over the 23 single-sheet ones, with the open untouched to the link. **Cycles**
(`perf stat` isolation pairs, median of five, CPU 24, A/A floor beside each):
`15228` second-cell **−41.00%** (floor −1.25%), `15228` one-cell **−36.73%**
(+2.24%), `54016` second-cell **−31.70%** (−0.16%), `54016` one-cell **−27.99%**
(−0.06%), `WithCustomViews` one-cell −10.07% (+0.48%). **Paired timing**, order
A1 B1 B2 A2 A3 as three interleaved rounds pooled per leg, 60 samples per leg, in
**two windows**: −40.45%/−39.43%, −35.34%/−35.46%, −32.36%/−32.23%,
−28.01%/−28.57% and −11.46%/−12.40% at p50 against a measured floor of ±1.6% in
the quiet window, with the earlier window at load 11-30 agreeing in sign on every
row and within 2.0 percentage points in magnitude. **One scenario got worse and
is reported rather than averaged away:** `15228.xls` full text is +0.84% in
cycles and +0.28%/+1.24% at p50 against a floor of −0.03%, because
splitting `parse_record_with` made LLVM outline the `CellRecord` drop glue that
had been inlined into its callers — 412,197 instructions moving into a shared
symbol against the 399,641 the chain hint removes from the two chain symbols;
its p95 and p99 fall in both windows. `#[inline(always)]` was measured as the
fix and rejected, because it helps that scenario and hurts the whole-sheet walk. No claim-registry entry;
`performance_claim: none`.
[Change and limitations](0641-xls-scan-measure-only-cells-and-sheet-hint.md);
[evidence](results/change-0641/README.md).

## For `ADR_COMPLIANCE.md`

## 0641 — a second validating instantiation, not a second validator

The compliance argument is structural rather than testimonial: **no check was
rewritten**. Every refusal the measure path can produce is produced by the same
function the materializing parse calls — one `cell_head` for the fixed-width
kinds, one `string_record_parts` for `Label`, one `frame_record` and
`check_token_stream` for `Formula` — so ADR 0005's mandatory validation is met by
construction, not by a promise that two copies agree, and the corpus differential
over 67,422 records with 142 refusals is the audit rather than the argument. Two
divergences exist and both are stated rather than absorbed. First,
`Ancillary::new`'s `Error::Allocation("retaining Formula RgbExtra")` is
unreachable on the measure path, because that path makes no reservation; it stays
reachable, in the same position and with the same text, on the materializing
path, and it needs an 8,224-byte allocation to fail. This is exactly the
disposition change 0576 recorded for `cannot allocate shared string characters`,
and it is the only error whose reachability moves. Second, the UTF-16 refusal is
*decided* by `char::decode_utf16` and *constructed* by re-running
`parse_string_record`, so `String::from_utf16`'s wording — the trap 0576 named —
cannot leak into a `DecodeUtf16Error` message; a test asserts the exact string
and asserts the absence of the other, and a mutation confirms it. The XF index of
every record is still validated in the position `accept` validated it, whether or
not the sink keeps the cell, and a test pins that a reserved style-XF slot on a
cell no query asked for is still refused. ADR 0003's bounded resources are
strengthened, not relaxed: the change only removes allocations. ADR 0006 is
untouched — no execution context, worker pool, ambient I/O or lock is added, and
the chain hint is a stack value that borrows its reader, so change 0579's three
properties (it cannot cross readers, cannot cross streams, cannot move a read
backwards) hold unchanged for the second position exactly as for the first.
`#![forbid(unsafe_code)]` is unchanged and this change adds no `unsafe` anywhere,
test code included. No public API, typed error, limit, fence or output byte
changes; no ADR is amended and no ADR clarification is proposed.
[Change 0641](0641-xls-scan-measure-only-cells-and-sheet-hint.md);
`performance_claim: none`.
