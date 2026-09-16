# The three `#[inline]` attributes, and the two measurements that placed them

Change 0641 splits `CellRecord::parse`'s per-kind checks into functions the
measure-only path also reaches: `CellRecord::cell_head`,
`utils::string_record_parts` and `formula_metadata::codec::frame_record` with
`check_token_stream`. Splitting a function that is called once per worksheet
record is a codegen decision as much as a structural one, and this note records
the two measurements that settled it, because neither is visible in the final
numbers.

## The first measurement, taken without any `#[inline]`

The first callgrind pass over the eight scenarios, on the split as first written,
found the **scan paths this change does not touch getting slower**:

| scenario | before | after (no `#[inline]`) | delta |
| --- | ---: | ---: | ---: |
| `54016` one-cell | 18,585,700 | 16,206,517 | −12.80% |
| `54016` second-cell | 31,806,931 | 27,031,143 | −15.01% |
| `15228` one-cell | 3,974,597 | 3,009,335 | −24.29% |
| `15228` second-cell | 6,671,272 | 5,035,313 | −24.52% |
| **`15228` all-cells** | 32,866,827 | 33,178,924 | **+0.95%** |
| **`15228` full text** | 76,279,594 | 76,734,923 | **+0.60%** |

The per-symbol difference named the cause immediately: `frame_record` appeared as
its own symbol at **+194,457** instructions per whole-sheet walk and
**+595,720** per text extraction, with `_int_malloc` and `_int_free_merge_chunk`
following it. `Framed` carries a `FormulaValue`, which owns a `String` in one
variant, so an out-of-line `frame_record` returns a **droppable temporary through
memory** on a path that runs once per `Formula` record.

`#[inline]` on `frame_record` and `check_token_stream`, and on `cell_head` and
`measure_fixed`, moved those two rows to **−0.35%** and **+0.14%** and took
`15228` second-cell from −24.52% to **−27.93%**. That is the configuration this
change ships, and it is why the attributes are load-bearing rather than
decorative.

The raw profiles of the no-`#[inline]` pass were not retained: they are callgrind
`.out` files of tens of megabytes each and the packet keeps folded output only,
as it does for the shipped pass. The table above is the folded result.

## The second measurement: `#[inline(always)]`, offered and rejected

`#[inline(always)]` on `frame_record` was then measured against the shipped
`#[inline]`, three repetitions per leg, `perf stat` isolation pairs at 20 and 120
samples on CPU 24. Raw output: [`../cycles/inline-trial/`](../cycles/inline-trial/).

| scenario | metric | `#[inline]` (shipped) | `#[inline(always)]` |
| --- | --- | ---: | ---: |
| `15228` full text | cycles | +0.42% | **+0.08%** |
| `15228` full text | instructions | **+0.09%** | +0.16% |
| `15228` one-cell | cycles | **−35.93%** | −34.80% |
| `15228` one-cell | instructions | −30.82% | −30.82% |
| `15228` all-cells | cycles | **−0.89%** | +1.83% |
| `15228` all-cells | instructions | **−0.06%** | +0.03% |

It buys back most of the full-text regression and pays for it on the whole-sheet
walk. On the deterministic metric — instructions — `#[inline]` is better or equal
on both scenarios, so `#[inline]` is what landed and the +0.84% full-text cycle
regression is reported rather than traded for a worse one elsewhere.

## A third artifact retained here: the flagship one-cell repeat

[`../cycles/flagship-repeat/`](../cycles/flagship-repeat/) holds five
repetitions of the flagship one-cell perf pair on three legs (`before`, `after`,
and `before` again). A single-shot pair in the first pass scored that scenario
**+11.17% cycles**; five repetitions with the A/A leg beside them scored it
**−3.31% against a −1.51% floor**. It is retained as the reason this record's
cycle table is a median of five repetitions rather than one pair: on a host
carrying a load average of 11 to 30, one `perf stat` pair on a 400,000-cycle
operation is not a measurement.
