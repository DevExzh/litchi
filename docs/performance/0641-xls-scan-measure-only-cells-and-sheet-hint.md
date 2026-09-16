# 0641: validate the cells a query is not looking for without building them, and resume the per-sheet cursor walk

Status: retained. `performance_claim: none` — this record carries a corpus
census, a corpus-wide differential over every cell record, exact chain-link
counts on three legs, callgrind isolation pairs, native cycles with a measured
A/A floor and paired wall-clock medians in both directions. None of it is
registered as a claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements items **XLS-5** and **XLS-8** of change
[0587](0587-remaining-opportunity-survey.md). XLS-5 is change
[0576](0576-xls-sst-scan-without-materialization.md)'s measure-only pattern
applied to worksheet cell records instead of shared strings. XLS-8 is the
per-sheet cursor construction term that change
[0585](0585-cfb-resumable-cursor-construction.md) priced and left open, and it
closes 0585's open question about whether one chain hint could serve both
consumers — by measuring it, where 0585 only argued it.

## What was changed

### XLS-5: a measure-only instantiation for the cells a query discards

`query_cell` frames every cell record on the worksheet — 37,929 of them on
`54016.xls` — parses each one into an 88-byte `CellRecord`, and keeps one. The
parse allocates for two kinds: `Label` transcodes its characters into a `String`
(`utils::parse_string_record`) and `Formula` copies its `Rgce` token stream with
`.to_vec()` and builds a `Metadata` (`formula_metadata::codec`). Everything else
is moved through the enum and dropped.

`CellRecord::measure` runs **every check** `CellRecord::parse` runs and builds
nothing. It returns a six-byte `MeasuredCell` — row, column, XF index — which is
all a scan needs from a record it is not keeping, because the XF index is
validated for every record whether or not the cell is the target.

It is not a second parser. Each kind's checks are the *same code*:

- the five fixed-width kinds (`Blank`, `Number`, `BoolErr`, `RK`, `LabelSst`)
  go through a new shared `CellRecord::cell_head(data, expected)`, which is the
  one length check that makes their three header reads infallible; every field
  after it is an in-bounds read of a fixed offset, so `cell_head` *is* their
  complete validation;
- `Label` goes through a new shared `utils::string_record_parts`, which owns the
  `XLUnicodeString` header, the declared character count, the `fHighByte` flag
  and the byte extent. `parse_string_record` transcodes what it returns;
  `measure_string_record` decides UTF-16 well-formedness with
  `char::decode_utf16`, which allocates nothing;
- `Formula` goes through a new shared `codec::frame_record` and
  `codec::check_token_stream`. `measure_record` calls `validate_formula_extra`
  where `parse_record_with` calls `retain_formula_extra` — the same scan of the
  same `RgbExtra` structures, stopping before the copy that retains them — and
  calls the same `decode_flags`, discarding the `Metadata` it returns.

The frame loop decides which to run from the four header bytes every BIFF8 cell
record opens with (`rw` and `col`, `[MS-XLS]` 2.5.19), which both parses read
identically, so the decision can never change which bytes are checked. A payload
too short to carry them is handed to the materializing parse, and therefore to
exactly the refusal it always produced. `CellSink` gains `wants(row, col)` —
`false` only for a selected-cell query's non-target cells — and
`accept_measured`, which validates the XF index in the same position `accept`
does. The whole-sheet walk answers `true` for every position and never reaches
the measure path at all.

`MulRk` and `MulBlank` are **not** routed through it. They already expand through
`visit_mul_*` without allocating, and a per-packed-cell variant would have to
rebuild the record for the target inside the visitor; that is left open below.

Six `#[inline]` attributes are part of the change, not decoration, and **four**
of them are load-bearing: `frame_record`, `check_token_stream`, `cell_head` and
`measure_fixed`, the functions the split created. `frame_record` returns a
`Framed` that carries a `FormulaValue`, which owns a `String` in one variant;
left out of line it returns a droppable temporary through memory, and the first
measurement of this change — retained below — showed that costing **+595,720
instructions** on one `15228.xls` text extraction and turning a whole-sheet walk
from −0.35% into **+0.95%**. The other two, on `TargetCell::wants` and
`wants_record`, sit on a predicate the frame loop evaluates once per record and
were not separately measured.

### XLS-8: a worksheet-region chain hint with a lifetime disjoint from the resolver's

`WorksheetScan::new` reached its sheet's first byte through
`stream_cursor_at`, which walks the workbook stream's allocation chain **from
its first sector, every time**. Change 0584 priced that at 28,143 links over a
16-sheet workbook's text extraction, about 91% of them re-walks of a prefix an
earlier sheet had already walked, and 0585 left it open because taking it needs a
*second* hint: a hint retains one position and is discarded when it sits past the
offset asked for, so one hint alternating between the string table and a
worksheet region would be discarded on every resolve and again on every sheet.

`WorksheetScan::new` now takes a `&mut StreamChainHint<'_>` and uses
`stream_cursor_at_hinted`. A document-wide text extraction creates one such
position beside the `SharedStringResolver`'s and carries it across the
document's sheets; a selected-cell query and a whole-sheet walk each create a
fresh one, which is exactly what `stream_cursor_at` did, because a single cursor
construction has no earlier position to resume from.

**What is not implemented, and why.** The brief also asked for the hint to live
in the *snapshot*, so that repeated queries resume it. That is a contract
change and is frozen as a design rather than implemented; see *The snapshot-scoped
form is not taken* below.

## Why it is sound

**The measure path refuses exactly what the parse refuses.** Not an appeal to
care: every check lives in one function reached from both paths. The three
shared framings above are the whole of it, and the corpus differential below is
the oracle.

**Error identity is preserved, including the trap change 0576 named.**
`String::from_utf16` fails with `FromUtf16Error`, whose `Display` is
`invalid utf-16: lone surrogate found`; `char::decode_utf16` fails with
`DecodeUtf16Error`, whose `Display` is `unpaired surrogate found: d800`.
`measure_string_record` uses the second only as a **boolean**, and when it says
malformed it re-runs `parse_string_record` on the same bytes, so the message is
produced by the identical code from the identical type. Cost on that path does
not matter; it is reached only by an input that is about to be refused.
`a_lone_surrogate_label_keeps_the_from_utf16_message` asserts the exact string
*and* asserts the message does not contain `unpaired surrogate`; mutation M2
confirms the assertion has teeth. `String::from_utf16` is
`decode_utf16(..).collect::<Result<_, _>>()`, so the two agree on exactly which
inputs fail, and both short-circuit at the first one.

**Refusal order is preserved.** `frame_record` performs the size, cached-value
and token-bound checks in the order `parse_record_with` performed them;
`measure_record` then runs `validate_formula_extra` where the parse runs
`retain_formula_extra`, then `check_token_stream`, then `decode_flags` — the
same four steps in the same sequence. `cell_head`'s length check is the first
thing every fixed-width kind did before and still is.
`formula_refusals_keep_their_position_on_the_measure_path` pins six distinct
`Formula` refusals and the two payload shapes that are accepted — a one-token
stream, and the empty token stream a string-valued `Formula` is allowed.

**One error becomes unreachable on the measure path, and it is stated rather
than hidden.** `Ancillary::new`'s `Error::Allocation("retaining Formula
RgbExtra")` is raised when the suffix reservation fails. The measure path makes
no such reservation, so on that path the error is unreachable; it stays
reachable, in the same position and with the same text, on the materializing
path. The reservation is at most `MAX_FORMULA_PAYLOAD` = 8,224 bytes, so the
divergence needs an 8 KiB allocation to fail. No test can reach it and none is
claimed. This is the same disposition change 0576 recorded for
`cannot allocate shared string characters`.

**The XF index is still validated for every record.** `accept_measured`'s only
job is `validate_cell_xf`, in the position `accept` validated it, so a cell that
names a reserved style-XF slot is refused wherever it sits on the sheet.
`a_query_refuses_a_bad_cell_format_it_is_not_looking_for` pins it and mutation N2
confirms it.

**The chain hint cannot serve the wrong bytes.** Nothing new is argued here:
`StreamChainHint`'s three properties (it cannot cross readers, cannot cross
streams, cannot move a read backwards) are change 0579's and change 0585's, the
per-link checks inside `next_chain_sector` still run on every link the walk
takes, and a hint that does not apply costs exactly what an unhinted call costs.
What is new is that this change uses a **second** position whose lifetime is
disjoint from the resolver's, which is the arrangement 0585's doc comment
requires, and that the counterfactual is now measured rather than asserted.

**Contracts untouched.** No public API, no typed error, no limit, no fence, no
ADR-governed boundary and no dependency edge changes. `MeasuredCell`,
`CellRecord::measure`, `cell_head`, `measure_string_record`, `string_record_parts`,
`frame_record`, `check_token_stream` and `measure_record` are all crate-private.
No `unsafe`: `crates/litchi-xls/src/lib.rs` keeps `#![forbid(unsafe_code)]` and
this change adds no `unsafe` anywhere, test code included. Fallible allocation is
unchanged — the change only removes allocations, it does not add or relax one.
ADR 0005's mandatory validation is met by construction: the scan still validates
every record to EOF and still refuses at the same record.

## Measured

Environment: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc
1.95.0, valgrind 3.26.0. Both legs are `--release --locked` builds of
`tools/perf-baseline`'s `xls_source_attribution`, staged outside any Cargo target
directory (0627), pinned to CPU 24. **Host quiescence is not established:** seven other agents were
building on this host, and the load average ran from 11 to 30 on 32 cores for
most of the capture, falling to 6.2 for the second timing window. That is why
every timing figure below is reported against an A/A floor measured in its own
window, and why the timing is reported in two windows rather than one.

Binaries, base and branch hashes: [`README.md`](results/change-0641/README.md).

### The census XLS-5 asked for first

The survey recorded the corpus `Formula`/`Label` mix as **unknown**: change
0584's census counted `LabelSst` only.
[`cell_record_census.py`](results/change-0641/probe/cell_record_census.py) walks
the CFB container, the BIFF8 framing and every worksheet substream of all 126
`.xls`/`.xlt` fixtures with pure arithmetic, sharing no code with `litchi-xls`.
It reproduces the survey's independently recorded **37,929 records** for
`54016.xls`'s single sheet exactly, which is the evidence that it is counting the
thing the scan frames.

372 worksheet substreams, 106,689 records, **127,072 cell values**:

| kind | records | share of cell values |
| --- | ---: | ---: |
| `LabelSst` | 30,316 | 23.86% |
| `MulBlank` (33,758 packed cells) | 5,699 | 26.57% |
| `MulRk` (26,817 packed cells) | 2,185 | 21.10% |
| **`Formula`** | **14,854** | **11.69%** |
| `Blank` | 13,938 | 10.97% |
| `RK` | 4,359 | 3.43% |
| `Number` | 2,524 | 1.99% |
| `BoolErr` | 500 | 0.39% |
| **`Label`** | **6** | **0.005%** |

Three findings the item did not anticipate, and they changed what was built:

1. **`Label` is effectively extinct.** Six records, 42 bytes of characters, in
   two fixtures (`embedded-chart.xls` and `WithFormattedGraphTitle.xls`). BIFF8
   producers write `LabelSst`. The `Label` measure path is a correctness
   obligation, not a saving, and this record does not claim otherwise.
2. **`Formula` is the allocating kind that matters**, on 36 of 126 fixtures:
   224,754 `Rgce` token bytes corpus-wide, of which `15228.xls` alone carries
   175,054 across 11,240 records — 73.8% of that fixture's 15,240 worksheet
   records. Only **5** `Formula` records in the whole corpus carry a nonempty
   `RgbExtra`, so `Ancillary` is essentially never built and
   `validate_formula_extra` essentially always returns at its first line.
3. **`54016.xls`, the fixture the survey priced `drop_in_place<CellRecord>` at
   2.05% on, carries zero `Formula` and zero `Label` records.** Its 37,929
   records are `LabelSst`, `Blank`, `MulBlank` and `RK`. That 2.05% is therefore
   the enum's *drop glue and 88-byte moves*, not heap frees — which is why the
   measure path covers all seven single-cell kinds and not only the two that
   allocate. `15228.xls` was added to the measured set because of finding 2; the
   survey named neither it nor the mix.

Full per-fixture and per-sheet output:
[`probe/census.jsonl`](results/change-0641/probe/census.jsonl).

### The control: not one byte of I/O moved

Logical reads, read bytes, `version()` calls and the harness's observation, over
16 scenarios and 5 samples each, are **identical** between the legs in all 16
cells. Six of them:

| scenario | reads | read bytes | `version()` |
| --- | ---: | ---: | ---: |
| `54016` one-cell | 65 | 932,993 | 40 |
| `54016` second-cell | 90 | 1,548,815 | 59 |
| `54016` all-cells | 16,145 | 1,256,139 | 16,117 |
| flagship one-cell (sheet 10) | 62 | 633,337 | 37 |
| `15228` one-cell (sheet 8) | 50 | 179,952 | 28 |
| `15228` full text | 595 | 631,201 | 4,571 |

Full table: [`counts/`](results/change-0641/counts/). This change removes CPU
work and chain walks only.

### XLS-8: chain links, three legs, one instrumented binary

Method and switches: [`chain/instrumentation.md`](results/change-0641/chain/instrumentation.md).
`base` reproduces the cold per-sheet construction, `hint` is this change, and
`shared` is the counterfactual in which one position serves both the resolver and
the worksheet cursor. Scenario: full text extraction, every `.xls`/`.xlt` fixture.

| | `base` | `hint` | `shared` |
| --- | ---: | ---: | ---: |
| **text-extraction links, 119 fixtures** | 2,453,391 | **2,405,381** (−1.96%) | 2,417,222 (−1.47%) |
| **the 96 multi-sheet fixtures** | 781,916 | **733,906 (−6.14%)** | 745,747 (+1.61% vs `hint`) |
| the 23 single-sheet fixtures | 1,671,475 | **1,671,475 (0.00%)** | 1,671,475 |
| open links | 14,806 | 14,806 | 14,806 |
| text digest mismatches | — | **0** | 0 |
| fixtures that got worse | — | **0** | 16 vs `hint` |

**The setup validated itself before it was believed.** The `base` leg reproduces
change 0585's recorded open figure for `ConditionalFormattingSamples.xls` —
**2,099** links — exactly, and the open is identical on all three legs for every
fixture, which is the control: this change operates entirely inside the scan and
the counter proves it rather than the prose asserting it.

Per fixture, the movers:

| fixture | worksheets | `base` | `hint` | removed | `shared` |
| --- | ---: | ---: | ---: | ---: | ---: |
| `15228.xls` | 18 | 15,643 | **5,391** | **65.5%** | 5,806 |
| `SimpleWithImages-mac.xls` | 3 | 508 | **173** | **65.9%** | 173 |
| `HyperlinksOnManySheets.xls` | 3 | 22 | **14** | **36.4%** | 17 |
| `50939.xls` | 2 | 363 | **218** | 39.9% | 218 |
| `29942.xls` | 3 | 722 | **464** | 35.7% | 464 |
| `ConditionalFormattingSamples.xls` | 16 | 75,263 | **67,458** | 10.4% | 70,664 |
| `external_name.xls` | 4 | 1,478 | **1,242** | 16.0% | 1,242 |
| `FormulaEvalTestData.xls` | 4 | 12,034 | **11,388** | 5.4% | 11,420 |
| `59858.xls` | 7 | 429,630 | **420,158** | 2.2% | 421,644 |
| `WithCustomViews.xls` | 3 | 55,215 | **54,714** | 0.9% | 54,714 |
| `54016.xls` | 1 | 1,669,938 | 1,669,938 | **0.0%** | 1,669,938 |

**Change 0585's unmeasured claim is now measured, and it holds.** That record
asserted in prose — and in `SharedStringResolver`'s doc comment — that one hint
serving both consumers "would be discarded as a backward step by every resolve
*and* again by every sheet, saving nothing on either path". On the multi-sheet
corpus the shared arrangement is **1.61% worse than the disjoint pair** and worse
on 16 fixtures, while still better than no worksheet hint at all. So the claim is
right in direction and overstated in degree: sharing does not save *nothing*, it
gives back a quarter of what the disjoint pair gives.

**Selected-cell queries and whole-sheet walks are unchanged to the link**, on all
four fixtures measured, which is the scoping claim stated as a number rather than
an intention:

| fixture | one-cell `base` → `hint` | all-cells `base` → `hint` |
| --- | --- | --- |
| `54016.xls` | 1,810 → 1,810 | 1,669,938 → 1,669,938 |
| `ConditionalFormattingSamples.xls` | 3,187 → 3,187 | 13,887 → 13,887 |
| `WithCustomViews.xls` | 344 → 344 | 54,629 → 54,629 |
| `15228.xls` | 679 → 679 | 877 → 877 |

Raw: [`chain/corpus.tsv`](results/change-0641/chain/corpus.tsv),
[`chain/query.tsv`](results/change-0641/chain/query.tsv).

### Instructions: callgrind isolation pairs

Profiles at 2 and 12 harness samples, differenced and divided by 10, so the
staging, oracle and reporting the harness performs once are removed.

| scenario | before | after | change |
| --- | ---: | ---: | ---: |
| `54016` one-cell | 18,587,922 | **16,206,514** | **−12.81%** |
| `54016` second-cell | 31,806,392 | **27,031,160** | **−15.01%** |
| `15228` one-cell (sheet 8) | 3,811,712 | **2,887,629** | **−24.24%** |
| `15228` second-cell | 6,682,372 | **4,815,914** | **−27.93%** |
| `WithCustomViews` one-cell | 1,696,209 | **1,599,314** | −5.71% |
| flagship one-cell (sheet 10) | 2,760,177 | **2,722,473** | −1.37% |
| `15228` all-cells | 32,896,484 | 32,782,851 | −0.35% |
| `15228` full text | 76,284,529 | 76,388,454 | **+0.14%** |

Per symbol, on `54016.xls`'s one-cell query — the fixture with no allocating cell
kinds at all, so this is the enum itself:

| symbol | before | after |
| --- | ---: | ---: |
| `TargetCell::accept` | 1,675,037 | **401,831** |
| `CellRecord::parse` | 893,405 | **40** |
| `drop_in_place<CellRecord>` | 662,150 | **158,831** |
| `query_cell` self | 2,983,554 | 2,450,644 |
| `number_format::codec::parse_xf` | 297,066 | 237,737 |
| `CellRecord::measure` | 0 | **806,196** |
| `alignment::CellAlignment::parse` | 0 | **95,480** |

The last two rows are the cost side and are not netted away. `measure` is the
work the query still does — it validates every record — and the 806,196 it
carries is what `parse` and the record's drop glue used to carry. The
`CellAlignment::parse` row is an **inlining shift in the open**, not new work:
`parse_xf` loses 59,329 in the same profile, and the whole-operation figure is
the column that settles it.

and on `15228.xls`'s, where the allocator moves instead:

| symbol | before | after |
| --- | ---: | ---: |
| `codec::parse_record_with` | 443,949 | **0** |
| `CellRecord::parse` | 222,264 | **30** |
| `malloc` + `free` | 234,890 | **17,551** |
| `drop_in_place<CellRecord>` | 109,875 | **9,622** |
| `__memcpy_avx_unaligned_erms` | 296,884 | 239,460 |
| `codec::measure_record` | 0 | 326,541 |
| `extra::validate_formula_extra` | 0 | 77,049 |

The `memcpy` figure carries change 0574's standing caveat — callgrind instruments
ERMS string loops per iteration and 0604 measured a 35× overstatement — so it is
an upper bound and the native cycles below are what price it.

### Cycles: `perf stat`, five repetitions, with the A/A floor beside each row

Isolation pairs at 20 and 120 samples, `-r 5`, median of five repetitions, pinned
to CPU 24. The `A/A floor` column is the **before binary against itself** in the
same window: a delta inside it is not a result.

| scenario | before cycles | after cycles | change | A/A floor |
| --- | ---: | ---: | ---: | ---: |
| `15228` second-cell | 1,546,420 | **912,396** | **−41.00%** | −1.25% |
| `15228` one-cell | 894,507 | **565,954** | **−36.73%** | +2.24% |
| `54016` second-cell | 6,498,129 | **4,438,293** | **−31.70%** | −0.16% |
| `54016` one-cell | 3,678,522 | **2,649,083** | **−27.99%** | −0.06% |
| `HyperlinksOnManySheets` full text | 112,524 | **96,146** | −14.56% | −3.10% |
| `WithCustomViews` one-cell | 331,451 | **298,061** | **−10.07%** | +0.48% |
| flagship full text | 2,156,870 | **2,070,203** | −4.02% | −0.28% |
| `WithCustomViews` full text | 5,710,906 | 5,582,339 | −2.25% | +0.02% |
| flagship one-cell | 433,893 | 424,723 | −2.11% | −1.63% |
| `15228` all-cells | 6,545,500 | 6,472,548 | −1.11% | −0.89% |
| **`15228` full text** | 19,186,474 | 19,346,884 | **+0.84%** | −0.09% |

Instructions from the same runs move the same way and are much quieter: −31.77%,
−30.10%, −17.04%, −14.84%, −0.12%, −7.28%, −3.74%, +0.29%, −2.61%, +0.06%,
+0.07%, against an A/A floor that is at most 1.56%, is 0.00% to two decimal places on
five of the eleven rows, and is below 0.25% on ten of them.

**Two rows are inside their own floor and are not claimed:** the flagship's
one-cell query (−2.11% against −1.63%) and `HyperlinksOnManySheets`'s full text
(−14.56% against −3.10%, on an operation of 112,524 cycles, which is too small to
measure on a loaded host). `15228` all-cells (−1.11% against −0.89%) is the
whole-sheet walk, which by construction never enters the measure path; that it
does not move is the result.

### The one regression, chased rather than hidden

**`15228.xls` full text is 0.84% slower in cycles and 0.14% more instructions**,
against an A/A floor of −0.09%. It is real and it is reported. Its cause is
named by the symbol difference: the text path is not otherwise touched by this
change, and it *gains* the chain hint —

| symbol | before | after |
| --- | ---: | ---: |
| `next_chain_sector` | 377,496 | **131,448** |
| `stream_cursor_at_hinted` | 272,670 | **119,077** |
| `scan_text_sheet` self | 1,783,477 | 1,663,035 |
| `drop_in_place<CellRecord>` | 2,624 | **412,197** |
| `codec::parse_record_with` | 1,360,040 | 1,427,480 |

— but splitting `parse_record_with` around `frame_record` made LLVM **outline**
the `CellRecord` drop glue that had been inlined into its callers. The work did
not appear: `drop_in_place<CellRecord>` goes from 2,624 to 412,197, a move of
**409,573** instructions out of the callers that had inlined it, and an outlined
call per `Formula` record is not free. The arithmetic, to the instruction: the two chain
symbols give up **399,641** and `scan_text_sheet` a further 120,442, against
**409,573** for the outlined drop glue, 70,444 for the `CellValue` drop that
follows it and 67,440 for `parse_record_with` — a net **+27,374** across the
named symbols, with the rest of the whole-operation **+103,925** spread below the
reporting threshold. The chain hint is paying for a codegen move, not for
itself.

`#[inline(always)]` on `frame_record` was measured as a candidate fix and
**rejected**: it improves full text (+0.42% → +0.08% cycles in a three-repetition
window) and makes the whole-sheet walk worse (−0.89% → +1.83% cycles, −0.06% →
+0.03% instructions). `#[inline]` is better or equal on the deterministic metric
on both scenarios, so it is what landed. Both codegen measurements — the first
pass with no attributes at all and this one — are retained in
[`callgrind/inline-trial.md`](results/change-0641/callgrind/inline-trial.md)
rather than discarded.

### Paired wall clock, in two windows

`p50` of the measured operation, nanoseconds, ordered A1 B1 B2 A2 A3 on CPU 24,
**three interleaved rounds** of that order pooled per leg so that a drift inside
the window is charged to every leg equally. `dir1` is the first before→after
pair, `dir2` the second, and `A/A` is the **before binary against itself** in the
same window.

**Window 2 is the table**, taken when the host had quieted to a load average of
6.2 on 32 cores. 60 samples per leg, all eleven scenarios by one method.

| scenario | before p50 | after p50 | `dir1` | `dir2` | A/A | window 1 |
| --- | ---: | ---: | ---: | ---: | ---: | --- |
| `15228` second-cell | 332,721 | **199,381** | **−40.45%** | **−39.43%** | −0.72% | −40.35% / −40.18% |
| `15228` one-cell | 191,821 | **123,951** | **−35.34%** | **−35.46%** | +1.56% | −33.38% / −35.01% |
| `54016` second-cell | 1,437,597 | **974,049** | **−32.36%** | **−32.23%** | −0.61% | −32.15% / −31.97% |
| `54016` one-cell | 818,074 | **586,068** | **−28.01%** | **−28.57%** | −0.10% | −27.51% / −27.56% |
| `WithCustomViews` one-cell | 70,930 | **62,445** | **−11.46%** | **−12.40%** | +0.54% | −12.39% / −11.30% |
| flagship one-cell | 88,650 | **85,875** | −3.55% | −2.68% | +0.59% | −2.82% / −3.30% |
| flagship full text | 470,887 | **457,427** | −2.92% | −2.98% | −0.14% | −2.99% / −2.81% |
| `WithCustomViews` full text | 1,262,916 | 1,247,216 | −1.80% | −1.07% | −0.66% | −1.63% / −1.15% |
| `15228` all-cells | 1,439,962 | 1,429,337 | −0.86% | −0.93% | −1.03% | +0.57% / +0.35% |
| `HyperlinksOnManySheets` full text | 23,730 | 23,550 | −0.67% | −0.84% | +0.30% | −1.27% / −1.06% |
| **`15228` full text** | 4,268,035 | 4,298,485 | **+0.28%** | **+1.24%** | −0.03% | +0.67% / +1.09% |

**The floor, measured in this window.** Same-binary excursions ran from −1.03%
to +1.56% at p50 across the eleven scenarios, so the floor is **±1.6%**. Five
scenarios beat it by 7 to 25 times and agree between their two directions to
within 1.1 percentage points. Three more sit just outside it — the flagship's
one-cell query and its full text, both around −3%, and `WithCustomViews`'s full
text, whose two directions straddle it at −1.80% and −1.07%. The `perf stat`
table agrees with the flagship's full text (−4.02%) and with
`WithCustomViews`'s (−2.25%), and puts the flagship's one-cell query inside its
own cycles floor, so that row is reported and not claimed. The three remaining
rows — `15228` all-cells, `HyperlinksOnManySheets` full text and `15228` full
text — are inside the floor in both directions.

**Window 1** was taken earlier, with the host at a load average of 11 to 30. Its
floor was ±2.5% and one scenario's single window put a 15% monotone drift on one
leg before it was re-measured as interleaved rounds. It is retained in full
([`timing/summary-window1.txt`](results/change-0641/timing/summary-window1.txt),
[`timing/window1-raw/`](results/change-0641/timing/window1-raw/)) and its
agreement with window 2 is the last column above: every row agrees in sign, and
the largest disagreement in magnitude is 2.0 percentage points on `15228`
one-cell, where both windows put the effect above 33%. The two windows were run
from the same two staged binaries.

Tails move with the medians on the scenarios that move: p99 falls from 832,674
to 601,783 and from 1,452,467 to 993,605 on `54016.xls`'s two queries, and from
214,621 to 131,610 and 383,352 to 207,311 on `15228.xls`'s. On `15228.xls`'s
full text, where p50 rises, **p95 and p99 fall** in both windows —
4,339,281 → 4,335,551 and 4,376,601 → 4,346,311 here — which is consistent with
the cause named below being a steady per-record cost traded against a chain walk
that is also a tail source.

Raw per-sample nanoseconds for every leg of both windows:
[`timing/`](results/change-0641/timing/).

## Correctness evidence

### The corpus differential, which is the oracle

`every_fixture_cell_record_measures_identically` frames every substream of every
fixture **by hand**, sharing no code with the scan it checks, and runs both
`CellRecord::parse` and `CellRecord::measure` over every cell record.

| | count |
| --- | ---: |
| `.xls` and `.xlt` fixtures under `test-data` | 126 |
| refused by the CFB container before any record is reachable | 2 |
| walked | **124** |
| **cell records compared** | **67,422** |
| of which refused by both paths, identically | **142** |
| **divergences** | **0** |

All seven kinds are covered and the test asserts they are, so the differential
cannot silently narrow: `Formula` 14,854, `LabelSst` 30,318, `Blank` 14,150,
`Number` 3,235, `RK` 4,359, `BoolErr` 500, `Label` 6. The four kinds that appear
only in worksheet substreams agree exactly with the independent Python census
above — 14,854, 4,359, 500 and 6 — which is two counts of the same corpus taken
by code that shares nothing.

`mutated_cell_payloads_are_refused_identically` reaches the refusals the corpus
does not contain: over the first 64 cell records of 40 fixtures it truncates each
payload at every length up to 28 bytes and flips three bit patterns at every
offset up to 28 — **32,884 mutations**, of which **9,240** are refused — and
compares the two paths on each. 0 divergences.

Two bounds on what these two sweeps establish, stated rather than left implicit.
Both compare the two paths' `Display` strings, not their `Error` **variants**,
so two variants that rendered identically would pass; every refusal this change
can reach is one of `InvalidLength`, `InvalidData`, `InvalidRecord`, `Encoding`
and `Allocation`, whose renderings are prefixed differently, but the comparison
is of text. And the mutation sweep stops at payload byte 28, so a deep
`RgbExtra` structure is exercised by the real corpus and not by mutation — the
corpus carries five such records.

### Tests added

Ten, in two groups. In `records.rs` (`cell_measure_tests`): the two corpus sweeps
above; `a_lone_surrogate_label_keeps_the_from_utf16_message`;
`every_two_unit_label_agrees_with_the_materializing_parse` (81 two-code-unit
labels drawn from both surrogate halves, both extremes of each, the units
immediately outside the surrogate block and two plain units);
`a_compressed_label_measures_without_decoding`;
`formula_refusals_keep_their_position_on_the_measure_path` (six refusals and the
two accepted shapes); `an_unknown_kind_is_refused_identically`. In
`tests/source_backed.rs`, three end-to-end:
`a_query_refuses_a_malformed_record_it_is_not_looking_for`,
`a_query_refuses_a_bad_cell_format_it_is_not_looking_for` and
`every_cell_the_walk_reports_is_still_found_by_a_query`.

This change is behaviour-preserving by construction, so no behavioural test can
flip against the pre-change tree, and none is claimed to. The tests are justified
by what they catch in the **new** code.

### Mutation checks

Twelve mutations were applied to the merged tree and the suites re-run. M1-M8
are re-run by [`mutate_records.py`](results/change-0641/mutate_records.py)
against `cell_measure_tests` and their output is
[`mutations.txt`](results/change-0641/mutations.txt); N1-N4 by
[`mutate_scan.py`](results/change-0641/mutate_scan.py) against the whole
`litchi-xls` suite, output
[`mutations-scan.txt`](results/change-0641/mutations-scan.txt).

| mutation | caught by |
| --- | --- |
| M1: `measure_label` drops the string refusal instead of reporting it | 4 tests |
| M2: the refusal carries `decode_utf16`'s wording, not `from_utf16`'s | `a_lone_surrogate_label_keeps_the_from_utf16_message`, `every_two_unit_label_…` |
| M3: `measure_formula` skips `validate_formula_extra` | `formula_refusals_keep_their_position…`, `mutated_cell_payloads_…` |
| M4: `measure_formula` skips `decode_flags` | 3 tests, including the corpus differential |
| M5: `measure_formula` skips the empty-token checks | `formula_refusals_keep_their_position…` |
| M6: `LabelSst` measures against the wrong minimum length | `mutated_cell_payloads_are_refused_identically` |
| M7: `wants_record` answers `false` for a payload too short to carry a position | **nothing — and it is equivalent** |
| M8: `measure_label` reports the wrong column | 4 tests |
| N1: the position peek reads the XF index where the column is | 10 tests |
| N2: `accept_measured` skips the XF validation | `a_query_refuses_a_bad_cell_format_it_is_not_looking_for` |
| N3: the measure path skips validation entirely | `the_whole_sheet_walk_refuses_what_a_selected_cell_query_refuses` |
| N4: one chain hint serves both the resolver and the worksheet cursor | **nothing — it is a performance property** |

**M7 survives because it is genuinely equivalent**, and that is worth stating
rather than counting as a gap: a payload shorter than four bytes routed to the
measure path meets the same length check it would have met on the parse path and
produces the same refusal, which is exactly what the corpus differential
establishes for every payload. The `is_none_or` is kept because it makes the
intent — *a record whose position cannot be read is not a record this scan may
skip building* — legible without depending on that equivalence.

**N4 survives because no test can catch it**, and that is the point of the
chain-link counter: sharing one hint is behaviour-preserving (a position that
does not apply is discarded), so it is invisible to every correctness test and
visible only as the 1.61% the `shared` leg above gives back. That is why the
constraint lives in `SharedStringResolver`'s doc comment and now in
`WorksheetScan::new`'s.

### Suites and gates

| slice | result |
| --- | --- |
| `cargo fmt --all --check` | clean |
| `cargo clippy -p litchi-xls --all-features --all-targets` | clean (workspace lints deny) |
| `cargo doc -p litchi-xls --no-deps` | clean (rustdoc lints deny) |
| `litchi-xls`, all features | **1,409 tests, 0 failures**, 72 binaries (1,399 before this change) |
| `cargo test -p litchi --features docx,xlsx,pptx,xls` | **266 tests, 0 failures**, 26 binaries |
| `litchi-xlsx`, `litchi-xlsb`, `litchi-cfb` | **2,400 tests, 0 failures**, 82 binaries |
| `tools/perf-baseline`, all features | 513 pass in the library suite; **2 fail** — see below |
| `tools/perf-baseline`, `--bins --no-fail-fast` | 19 binary targets, **71 tests, 0 failures** |
| `tools/check_crate_boundaries.py` | exit 0 |
| `tools/check_report_claim_classification.py` | exit 0, 167 rows, `strict_claim=0` |
| `tools/check_perf_claims.py --mode structural` | one complaint — **pre-existing**, see below |

**Two harness failures, pre-existing, flaky, and reproduced on the untouched
base.**
`docx_bounded_tail_append_compare`'s
`allocator::tests::global_allocator_records_successful_alloc_and_dealloc_with_live_peak`
and `…_records_zeroed_alloc_and_zeroes_memory` assert on a **process-global**
allocator counter and observe allocations made by other tests running in
parallel in the same process. They reproduce on the untouched before checkout at
`c7326f680`, they pass on both trees under `--test-threads=1`, and they are
**flaky rather than deterministic** in either tree — an isolated parallel run of
that binary fails on the base and passes on this branch, which is the signature
of a timing-dependent counter and not of an assertion about the code under test.
`gates.txt` records all six runs, including one in which that binary passes.
`cargo` stops on the failure, so the combined invocation does not reach the
remaining binary targets; run on their own with `--bins --no-fail-fast` all 19
of them pass, that binary included. No file this change touches is an input to
it.

**One claim-registry complaint, pre-existing.**
`check_perf_claims.py --mode structural` reports that
`claim-0251-xlsx-xml-borrowed` "requires strict evidence verification". The
identical line comes out of the before checkout. This change adds no
claim-registry entry and does not touch the registry.

The two corpus sweeps this record cites were run with `--nocapture` and their
output is in [`gates.txt`](results/change-0641/gates.txt).

## Validation preserved

Every record a scan frames is still parsed to the same depth, refused on the same
inputs, with the same message, in the same order. Nothing is skipped and no
refusal moves. The scan still validates the worksheet substream to its EOF; this
change does not touch change 0574's opportunity 6 and nothing here should be read
as bounding the cost of a query. The XF index of every cell is still validated.
The freshness fences, the execution checks, the worksheet byte and record limits,
and the window fills are untouched — the read counters above are the evidence.

## Limitations

**What is not claimed.** No cold-cache, physical-device, remote or range-source,
peak-RSS, concurrency-scaling, real-producer or cross-platform result. The timing
figures are warm-cache, page-cached, on staged immutable copies, on one host,
from two binaries that differ in five production files, and **host quiescence is
not established** — the load average was 11 to 30 on 32 cores for the counts,
callgrind, cycle and window-1 captures, and 6.2 for window 2. The A/A columns
are the defence against that and are reported for every row.

**`Label` savings are not claimed at all.** Six records exist corpus-wide. The
`Label` measure path is there so that the measure instantiation is complete, not
because it pays.

**`MulRk` and `MulBlank` are not covered.** They are 60,575 of the corpus's
127,072 cell values — 47.7% — and each packed cell still builds and drops an
88-byte `CellRecord` on a query that does not want it. They allocate nothing, so
what is left on the table is the enum move and its drop glue, not heap traffic.
Routing them needs `visit_mul_*` to yield positions and rebuild the record inside
the visitor when the target is reached; it is a separate change.

**The snapshot-scoped form of XLS-8 is not taken, and it is a contract change.**
The brief asked for the hint to live in the snapshot so that repeated queries
resume it. That cannot be done as written: `StreamChainHint<'a>` borrows the
`SharedOleFile` it came from, and `SourceInner` owns that reader, so storing one
in the snapshot is a self-reference no safe Rust expresses. The two ways around
it each break something this program does not break on a measurement:

- exposing the retained position as a plain reader-independent value that a
  `SharedOleFile` can be re-seeded from **weakens the first of the three
  properties** `StreamChainHint`'s documentation states — "it cannot cross
  readers", enforced today by a borrow and an address comparison. A position
  carrying a matching SID and first sector but taken from a different reader's
  FAT would resume at a sector that reader's cold walk would not reach, which is
  the one thing the type promises cannot happen;
- keeping the borrow and putting the position behind interior mutability adds a
  lock to an immutable, `Clone`, cross-thread snapshot on the selected-cell path,
  which is a parallelism cost this record has not measured and ADR 0005 treats as
  a clean-value cache needing its own design record — the same gate XLS-3 is
  behind.

The measured value is also smaller than it looks. A second query would resume
only the *construction* walk, and a whole one-cell query costs **1,810** chain
links on `54016.xls` and **3,187** on the flagship — the construction is a part
of those, and change 0584 put it at 2,238 of a flagship query's 4,428 on its own
base — against the **1,669,938** that fixture's whole-sheet walk performs. The
ceiling on the snapshot-scoped form is therefore a few thousand dependent loads
per repeated query, which is not obviously worth a lock on a shared snapshot.
The design is frozen here; it is not implemented and it is not rejected.

**The hint records the sheet's start, not where the scan ended.** A sheet's
cursor resumes from the *previous sheet's first sector*, not from where the
previous sheet's scan finished, because only construction records into the hint.
That is the 91%-resumable shape change 0584 modelled, and closing the remaining
gap needs a way to publish a live cursor's position into a hint, which
`litchi-cfb` does not have.

**The corpus chain-link figure is dominated by two fixtures.** `54016.xls`
(1,669,938 links, one sheet) and `59858.xls` (429,630) are 85% of the corpus
total, and neither is a multi-sheet case. The corpus-wide −1.96% is reported
because it is the honest aggregate; the −6.14% over the 96 multi-sheet fixtures is
the figure that describes what the change does.

**Chain links are a count of dependent loads, not a latency**, and change 0585's
caveat applies unchanged: the relationship between −6.14% of links and any
wall-clock figure is not established by the counter. What relates them here is
the `perf stat` row for full text on the fixtures that move.

**Four measurement gaps.** No allocation counter was run for this change
(change 0576's `sst_scan_allocations.rs` pattern would apply directly, and the
callgrind `malloc`/`free` figures are the only allocation evidence here). No
`from_path` or range-source leg was taken, so the `fstat` cost of the text path's
per-string fences (XLS-4) is still invisible. The flagship's text extraction is
still refused by this reader, so its 16 sheets are only partly walked before the
refusal. And `second-cell` in the harness queries two cells of the **same**
worksheet; the brief's "second cell on another sheet" has no selector, and would
not have moved anyway, since XLS-8 does not serve repeated queries.

## Retained evidence

[Census, differential output, chain-link legs, callgrind, cycles, timing, gates
and the instrumentation note](results/change-0641/README.md).
