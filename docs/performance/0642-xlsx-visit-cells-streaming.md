# 0642: the XLSX range visitor stops building the range it is about to visit

Status: retained, implemented in `litchi-xlsx`. `performance_claim: none` — the
counts and paired medians below are reported as evidence, not registered as a
claim. The change is **value-identical and refusal-identical**: a four-way
differential over 397 worksheets of 182 `.xlsx` files produces a byte-identical
table on the before and the after binary, including all thirteen refusals.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements item **XLSX-7** of
[`0587-remaining-opportunity-survey.md`](0587-remaining-opportunity-survey.md)
(unranked, low risk; `results/change-0587/survey/xlsx.md` item R7), which
observed that `SourceBackedWorksheet::visit_cells` "is documented as a visitor
but materializes the whole range first", modelling the vector it builds at over
5.2 MB for a 65,536-cell range. It is the OOXML twin of change
[0605](0605-xls-retained-sheet-index.md), which gave the XLS source-backed
worksheet a one-scan whole-sheet walk. The selected-range scan it sits on was
introduced by changes
[0365](changes/0365-xlsx-source-worksheet-range-streaming.md) and
[0366](changes/0366-xlsx-selected-general-references.md).

## What was changed

One file: `crates/litchi-xlsx/src/workbook/source.rs`.

`visit_cells` used to call `cells`, take the `Vec<SourceCell>` it returns, and
iterate it. `SourceCell` is 80 bytes, so a whole-sheet visit built a second
whole-range vector on top of the one the scan had already retained — and on the
materialized-store route it also **cloned every `Cell`** out of the store into
that vector, only to hand the caller a reference to the copy.

The route decision and every step that reads the source moved into a new private
`select_cells`, which returns a private `Selection`:

```rust
enum Selection<'a> {
    /// The materialized worksheet store is authoritative for this read.
    Stored(&'a Store),
    /// Physical records retained by the bounded selected scan, in source
    /// order, with the shared-string payloads they depend on resolved and
    /// every record validated.
    Selected {
        cells: Vec<raw::selected_worksheet::SelectedRecord>,
        shared_text: Option<Vec<(usize, Text)>>,
    },
}
```

`cells` converts a `Selection` into exactly the `Vec<SourceCell>` it returned
before; `visit_cells` walks the same `Selection` and produces one cell at a
time. On the stored route it now hands the callback the store's own `&Cell`, so
that route allocates nothing at all. The former `eager_cells` became the free
function `collect_stored_cells`, unchanged; the conversion arms of the former
`stream_cells` became `resolve_selected_record`, unchanged, shared by both
readers; the per-cell cancellation fence, callback and bounded count became
`visit_one`, unchanged.

**No scanner change.** `raw/worksheet/selected.rs` is untouched, and the shorter
route — making the scan yield to the visitor — was not taken. See below.

One addition that is not a refactor: `stream_selection` now walks the retained
records once, before it returns them, and refuses a record that carries both a
semantic cell and a shared-string dependency, or neither. That check used to sit
inside the conversion loop.

## Why it is sound

**The order-of-refusal contract is preserved, structurally.** The scan's own
documentation states that "the returned eligible value is published only after
the shared MCE/XML stream reaches EOF": `scan_stream` feeds the whole worksheet
through `x14ac::capture_stream_with_active` and only then calls
`Scanner::finish`. So today every refusal the worksheet can raise — including
one in a row *after* the requested rectangle — arrives before any cell is
visited. That is unchanged, because the scan still runs to EOF before
`select_cells` returns. What the change had to add is the tail: after the scan,
`stream_selection` resolves the shared-string and style dependencies, decides
the fallback, runs both source/execution fences, and now validates every
retained record — so when the first callback runs, every refusal a whole-range
read of that rectangle can raise has already been raised. The two record
refusals hoisted into that pass are **defensive, not reachable**: `SelectedCells`
is only ever built by `Scanner::finish` from records pushed by `retain_selected`,
whose two call sites each set exactly one of `cell` and `shared_string_index`.
The other two arms of `resolve_selected_record` are unreachable for a different
reason — the existing fallback pre-passes already convert every index with
`usize::try_from` and binary-search every index in the resolved shared-string
table, falling back to the store rather than refusing if either fails.

**Error identity is unchanged.** Every `invalid(...)` message and every fence is
the same text in the same order. `cells` is all-or-nothing, so moving a check
from its conversion loop into a pass before it cannot be observed; `visit_cells`
is the caller that can observe it, and for it the pass is the point.

**The fences did not move.** `cells` publishes through `finish_result` before a
caller can observe anything; `visit_cells` now runs `finish_result` on the
selection before the first callback, exactly where the inner `cells` call used
to run it, and runs it again at the end, so a source mutation or a cancellation
remains primary over a callback error. `Selection` holds no source reader in
either variant: the `Stored` arm borrows the published `OnceCell` store, and the
`Selected` arm owns records produced after `with_verified_decoded_reader`
returned and after the dependency readers were released. The documented promise
that "no callback runs while a source reader is active" therefore still holds,
with the same fence that held it before.

**ADR reading.** ADR 0005 makes semantic payloads lazy and cache behavior
"semantically invisible", and this change keeps both routes and their
laziness assertions intact — a cold `visit_cells` still leaves
`data.cells` unpublished on an eligible worksheet. ADR 0003's borrowed-view rule
is what the visitor was always for: the callback receives a borrowed `&Cell`,
and conversion to owned storage stays explicit in `cells`. Nothing in ADR 0006
is touched: no limit, no defence, no validation and no output byte changed. No
new `unsafe`, no new dependency, no public API or error type change.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0
(the repository's pinned toolchain), valgrind/callgrind 3.26.0, `taskset -c 17`,
with seven other agents building on the host throughout. Both legs built
`--release` with identical flags from a probe whose only difference is the path
dependency. Binaries were staged outside every Cargo target directory before
being run. Corpora: `no_drawing_patriarch.xlsx` (a 672,414-byte POI
fixture whose single worksheet part is 3,382,556 bytes uncompressed and holds
75,770 stored cells, backed by a 3,440,972-byte shared-string table) and a probe
workbook whose shape follows the harness `dense-wide` corpus — two sheets of
256 × 256 integer cells, 384,231 bytes against the harness corpus's 384,525,
built by the probe itself so the packet can reproduce it.

### Allocations and bytes

A whole-sheet visit over an already-materialized store. This is the operation
`visit_cells` exists for, and it becomes allocation-free:

| corpus | metric | before | after |
| --- | --- | ---: | ---: |
| dense-wide-probe `Sheet1`, 65,536 cells | allocation calls | 65,551 | **0** |
| | allocated bytes | 5,608,448 | **0** |
| | peak live bytes | 5,608,448 | **0** |
| `no_drawing_patriarch` `Лист 1`, 75,770 cells | allocation calls | 16 | **0** |
| | allocated bytes | 10,485,760 | **0** |
| | peak live bytes | 10,485,760 | **0** |

The two corpora remove different things. The dense probe's cells are numeric, so
`Cell::clone` allocated one `Box<str>` numeral per cell (65,536 of the 65,551
calls) on top of the vector; the POI worksheet's cells are shared-string text,
whose `Arc` clone allocates nothing, so all sixteen calls were the vector's own
geometric growth to a 131,072-slot capacity — 10,485,760 bytes for 75,770 cells,
a 42% over-allocation that a `try_reserve(1)` loop cannot avoid.

A cold whole-sheet read, open included. The two corpora take different routes:
the probe corpus is eligible for the bounded selected scan, the POI worksheet
falls back to the materialized store.

| corpus (route) | metric | before | after | delta |
| --- | --- | ---: | ---: | ---: |
| dense-wide-probe (selected scan) | allocation calls | 2,234,372 | 2,234,371 | −1 |
| | allocated bytes | 72,410,333 | 67,167,453 | −5,242,880 (−7.24%) |
| | peak live bytes | 14,615,454 | 14,615,454 | 0 |
| `no_drawing_patriarch` (stored) | allocation calls | 5,570,626 | 5,570,610 | −16 |
| | allocated bytes | 296,248,621 | 285,762,861 | −10,485,760 (−3.54%) |
| | peak live bytes | 57,025,552 | 57,025,552 | 0 |

The −5,242,880 is exactly 80 bytes × 65,536: on the scan route the conversion
moved cells rather than cloning them, so the whole saving is the vector itself,
in one `try_reserve_exact` call.

**Peak live bytes do not move on a cold read, and that is the honest limit of
this change.** The peak of a cold whole-sheet read is set earlier — by the
worksheet payload plus the scan's own record vector on one route, and by the
eager parse on the other — and the vector this change removes was allocated
after that peak had passed. The peak only moves when the store is already
warm, which is the table above it.

`cells` is unchanged on every corpus, route and column, cold and warm: the same
2,234,372 / 72,410,333 / 14,615,454 and 5,570,626 / 296,248,621 / 57,025,552.

### Instructions

Callgrind isolation pairs, cache and branch simulation off, N and N+M samples of
the operation differenced and divided by M. The probe's counts are deterministic
to within about 100 instructions across repeats.

| scenario | M | before | after | delta |
| --- | ---: | ---: | ---: | ---: |
| warm whole-sheet visit, dense-wide-probe | 4 | 40,796,855 | 7,483,791 | **−81.66%** |
| warm whole-sheet visit, `no_drawing_patriarch` | 4 | 16,823,824 | 8,929,044 | **−46.93%** |
| cold whole-sheet visit, dense-wide-probe | 2 | 1,435,103,151 | 1,437,768,943 | +0.186% |
| cold whole-sheet `cells`, dense-wide-probe | 2 | 1,431,667,393 | 1,434,574,716 | +0.203% |

Per visited cell the warm walk falls from 622.5 to 114.2 instructions on the
dense probe and from 222.0 to 117.8 on the POI worksheet; the two after figures
agreeing to within 4% is what one expects once the per-cell work is a bounds
check and a callback rather than a copy. Callgrind prices `rep movsb`/`rep stosb`
per byte, so the copy-heavy before leg is over-counted relative to native cycles;
the paired timing below is the native check.

The two cold rows are the change's cost, and they are reported because they are
adverse. Per-symbol differencing attributes **+852,085 instructions per open** —
about 29% of the +0.20% — to `select_cells`'s new record-validation pass, which
is 13 instructions for each of 65,536 records. The remainder is drop-glue and
inlining attribution churn from routing both readers through one conversion
helper; no litchi symbol gains identifiable work.

### Paired timing

Order A1 B1 B2 A2, one process per leg, 40 samples per leg on the warm
scenarios and 30 on the cold ones, `taskset -c 17`, with an A/A floor measured in
the same window. Three windows were taken because the host is shared with seven
other agents and its floor moves; every window is retained.

| window | scenario | A/A floor p50 | A1→B1 p50 | B2←A2 p50 | speedup |
| --- | --- | ---: | ---: | ---: | ---: |
| 1 | warm whole-sheet visit, dense-wide-probe | −2.74% / +0.24% | **+83.52%** | **+83.59%** | 6.07× / 6.10× |
| 1 | warm whole-sheet visit, `no_drawing_patriarch` | (same) | **+75.69%** | **+75.35%** | 4.11× / 4.06× |
| 3 | warm whole-sheet visit, dense-wide-probe | −3.87% / −0.73% | **+83.75%** | **+83.89%** | 6.15× / 6.21× |
| 3 | warm whole-sheet visit, `no_drawing_patriarch` | (same) | **+77.74%** | **+78.30%** | 4.49× / 4.61× |

Full quartiles for window 1's dense probe: before p50 1,928,086 ns, mean
1,930,191, p95 1,949,290, p99 1,980,860; after p50 317,692, mean 320,495,
p95 331,162, p99 392,852. For the POI worksheet: before p50 1,780,715 /
mean 1,785,868 / p95 1,834,119 / p99 1,962,981; after p50 432,962 /
mean 434,204 / p95 440,713 / p99 508,253. The reduction holds at every quartile,
in both directions, in both windows, and is about twenty to thirty times the
floor.

**The cold scenarios are not resolvable at this host's floor and no timing
result is claimed for them.**

| window | scenario | A/A floor p50 | A1→B1 p50 | B2←A2 p50 |
| --- | --- | ---: | ---: | ---: |
| 1 | cold whole-sheet visit, dense-wide-probe | −3.07% / +14.57% | +1.02% | −1.24% |
| 1 | cold whole-sheet `cells`, dense-wide-probe | (same) | +1.22% | +2.75% |
| 1 | cold whole-sheet visit, `no_drawing_patriarch` | (same) | +1.76% | +2.98% |
| 2 | cold whole-sheet visit, dense-wide-probe | +1.51% / −0.97% | +1.70% | +0.13% |
| 2 | cold whole-sheet `cells`, dense-wide-probe | (same) | +3.80% | −0.04% |
| 2 | cold whole-sheet visit, `no_drawing_patriarch` | (same) | +2.18% | +3.21% |

Window 1's cold floor reached 14.57% at p50, far past the 5% threshold this
program treats as the point where timing stops resolving anything, so the
deterministic counts above carry the cold path. Window 2's
floor was ±1.5%, and in it the only delta that clears the floor in both
directions is the cold POI visit at +2.18% / +3.21% in the change's favour — the
stored route's removed 10.5 MB vector and its memory traffic. That is consistent
with window 1's +1.76% / +2.98% but is not claimed: two windows on one corpus
is not a result.

The harness selectors were run for the same reason. `xlsx_narrow_column_range_scan`
is an eager-door control this change cannot reach; `xlsx_source_narrow_column_range_scan`
reads `B1:B256` of the dense-wide corpus through `SourceWorksheet::cells`, which
this change re-plumbed.

| window | selector | A/A floor p50 | A1→B1 p50 | B2←A2 p50 |
| --- | --- | ---: | ---: | ---: |
| 1 | `xlsx_source_narrow_column_range_scan` | +2.72% / +1.56% | +1.42% | −0.35% |
| 2 | `xlsx_source_narrow_column_range_scan` | +10.85% / −1.14% | −0.44% | +8.07% |
| 1 | `xlsx_narrow_column_range_scan` (11–13 µs) | +10.55% / +2.00% | +0.47% | +11.64% |
| 2 | `xlsx_narrow_column_range_scan` (11–13 µs) | +13.47% / −6.39% | +3.75% | +3.26% |

Every delta above is the size of its own window's A/A floor or smaller, in both
directions, and it changes sign between windows; the eager control's 11–13 µs
scale leaves it noise-dominated in all four figures, which is why its window-1
`+11.64%` is read as noise beside a `+10.55%` floor rather than as a result on a
path this change cannot reach. An exploratory window taken earlier in this batch,
before the final binaries were staged, put `xlsx_source_narrow_column_range_scan`
at −3.07% / −2.82% at p50; it did not reproduce in either retained window and its
raw samples were not kept, so it is reported here rather than in the packet.
The selector's callgrind
isolation pair moved −0.293%, inside the 0.34% spread that three pairs of the
*identical before binary* showed (`callgrind/harness-nondeterminism.txt`): the
harness uses per-process hashed containers, so that pair bounds the work rather
than measuring it.

## Correctness evidence

**The differential oracle.** `probe/probe.rs`'s `diff` mode runs four legs for
every worksheet of every named file, each from its own freshly opened workbook:
cold `visit_cells`, cold `cells`, `visit_cells` after `stored_extent` has
published the materialized store, and the eager `litchi_xlsx::Workbook` store's
own `cells` iterator. Each leg produces a visited count and an FNV-1a digest
over the `Debug` rendering of every `(address, cell)` pair in the order it was
produced, or the refusal's `Display` text. Over the **181 `.xlsx` files in the
tree plus the probe corpus — 397 worksheets** — all four legs agree on every
worksheet, on both binaries, and the two 397-row tables are byte-identical
(`diff-before.tsv`, `diff-after.tsv`, `diff -q` clean). Thirteen worksheets
refuse; the refusal text is identical across all four legs and both binaries,
including the shared-formula, empty-formula and invalid-dimension refusals.

**Tests added**, in `crates/litchi-xlsx/src/workbook/source.rs`'s new
`streaming_0642_tests` module:

* `visit_cells_0642_matches_cells_on_both_routes` — a shared-string-bearing
  sparse worksheet; the cold and the warm-store visit each reproduce `cells`'s
  sequence and count exactly.
* `visit_cells_0642_refuses_a_later_row_before_any_visit` — a worksheet whose
  malformed cell is in row 9, with `A1:A2` requested; the refusal text equals the
  whole-range read's and **zero** callbacks ran. This is the order-of-refusal
  contract as a test.
* `visit_cells_0642_stops_on_the_stored_route_at_the_failing_cell` — the stored
  route stops at the third callback, reports that count and returns the callback's
  own error.

All three pass on the **unmodified base** as well (`tests-on-base.txt`), which is
what they are for: they pin behavior this change preserves, not behavior it
introduces.

**Gates** (`gates.txt`): `cargo fmt --all --check`, `cargo clippy -p litchi-xlsx
--all-targets`, `cargo doc -p litchi-xlsx --no-deps`, `cargo test -p litchi-xlsx`
(1,311 tests, 0 failures, up from 1,308) and `--all-features` (1,330, 0), `cargo
test -p litchi --features docx,xlsx,pptx,xls` (266, 0), and `cargo test` in
`tools/perf-baseline` (531 tests, 0 failures, the lib suite alone 514 in
710.40 s) — the two suites change 0639 showed are reachable from no per-crate
gate. `tools/native-resave` is the third `litchi-xlsx` consumer and
was built and tested too. Two sets of warnings are pre-existing and reproduce
with identical text on the untouched before checkout: six in the `litchi` facade
under `--features docx,xlsx,pptx,xls`, and four in `tools/native-resave`.

## Validation preserved

Nothing was removed from any validation path. The bounded selected scan still
runs to worksheet EOF under the same `StreamLimits`, the same `Capabilities` and
the same `selected_stream_limits` ceiling; the shared-string and style dependency
resolutions are unchanged, as are their four fallback conditions; the
materialized store's parse, `validate_styles` and cancellation checks are
unchanged. A record-level check was **added**, not relaxed. Both source-version
fences and every `execution_check` occur at the same points, and the final
`finish_result` still gives a source mutation or a cancellation precedence over a
callback's own error.

## Limitations

* **No claim is registered.** The paired medians and counts are evidence.
* **Peak live bytes fall only for a walk over a materialized store.** On a cold
  whole-sheet read the peak is set before the removed vector is allocated and does
  not move at all. The record says so above with the measurement.
* **The cold path costs +0.19% to +0.20% of instructions** and its timing sits
  inside every window's floor. That cost is real, is mostly the price of
  validating the retained records before the first callback, and is reported
  rather than hidden.
* **The scan still retains one record per selected cell.** Making
  `scan_range` yield to the visitor would remove that vector too, but it would
  move refusals after the first visit, which is a contract change; it needs its
  own frozen design record, and this change deliberately did not take it. That
  remaining `Vec<SelectedRecord>` is the reason the cold peak does not move.
* **`cells` still over-allocates on the stored route.** `collect_stored_cells`
  keeps the original `try_reserve(1)` growth loop, which reached a 131,072-slot
  capacity for 75,770 cells (10,485,760 bytes, 42% over) on the POI worksheet.
  This change was scoped to leaving `cells` returning the same vector, so
  the loop is untouched; sizing it from the store's own count is a separate,
  unmeasured item.
* **Two corpora, one machine, one build.** Both are single worksheets read
  whole. The two harness selectors do cover a narrow `B1:B256` range, but through
  `cells`, not `visit_cells`, and they sit inside their own floors, so nothing is
  claimed for a narrow-range visit either. Nothing is claimed for the one-cell
  `cell()` door, for a range-source transport, for cold caches, for RSS, for
  throughput, for other formats, or for any platform but this one.
* **The probe corpus is not the harness corpus.** It follows the same shape and
  is 294 bytes smaller; it is built by the retained probe so the packet
  reproduces it, but it is not byte-identical to `litchi-xlsx-synthetic-v1`.

## Retained evidence

[`results/change-0642/README.md`](results/change-0642/README.md) — the probe
sources and their pinned toolchain, the differential tables for both legs, the
allocation report, every callgrind isolation-pair total, the paired-timing runs
and their floors, `decision.json`, `gates.txt` and `log-sections.md`.
