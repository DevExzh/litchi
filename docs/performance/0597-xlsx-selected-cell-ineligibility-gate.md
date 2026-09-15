# 0597: the ineligible selected-cell scan stopped seeing `sheetData` and refused merges it could no longer place; the gate that skips the scan stops at its design

Status: retained, mixed outcome. `performance_claim: none`. One correctness fix
is implemented in `litchi-xlsx`; the ineligibility gate that this change set out
to build is **not implemented** and is frozen here as a design with its measured
prize, its rejected patch and the exact error classes it moves.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This record answers item **XML-2** of
[`0587-remaining-opportunity-survey.md`](0587-remaining-opportunity-survey.md)
(rank 6): *"selected-cell stream runs to EOF after the sheet is known
ineligible, then the store re-parses"*.

## What was changed

`crates/litchi-xlsx/src/raw/worksheet/selected.rs`. The scanner's post-mark
branch no longer asks whether a `<mergeCells>` element "appears before
sheetData".

`Scanner::event` returns early for every event after `Scanner::mark` records an
ineligibility reason. That early return is also what stops `seen_sheet_data`
from ever advancing: a `<sheetData>` that follows the mark is never recorded.
The same branch then called `validate_merge_cells_placement`, whose first
clause refuses a `<mergeCells>` when `!seen_sheet_data`. On every worksheet
marked at worksheet level *before* its `<sheetData>` — which `<cols>` does, and
so does any unmodelled worksheet child such as `<sheetViews>` — a correctly
ordered `<mergeCells>` was therefore refused as misplaced.

The post-mark call is now `validate_marked_merge_cells_placement`, which keeps
the two clauses that read state completed before the mark (`seen_merge_cells`
for a duplicate `<mergeCells>`, `merge_window_closed` for one after a schema
successor) and drops the `seen_sheet_data` clause. `validate_dimension_placement`
is unchanged: its `seen_sheet_data` clause is one-sided in the other direction
(it can miss a late `<dimension>`, never invent one), and the pre-mark
`validate_merge_cells_placement` is unchanged.

The module and outcome documentation stops asserting that a
`NotEligible` result "asserts only that streaming XML/MCE/raw validation reached
EOF successfully"; it now says what change 0362 already required, that the
verdict is never worksheet semantic validity and the materialized parser owns
the mandatory validation and the first typed error.

Two regression tests in `raw/worksheet/tests.rs`
(`change_0597_marked_worksheet_keeps_a_correctly_placed_mergecells`,
`change_0597_marked_worksheet_still_refuses_misplaced_merge_markup`) and one
public-path test in `tests/source_backed_merge.rs`
(`source_backed_cell_reads_a_column_formatted_worksheet_with_merges`) pin both
directions. The first and third fail at `08d968f8e` and pass here; the second
passes on both legs, which is its point.

## What was not changed, and why

The brief asked for an ineligibility gate: stop the stream at the first mark, or
pre-gate with a bounded `memmem` before opening it, then fall back exactly as
today. Both were designed, one was implemented and measured, and it is **not**
landed. The design and the reasons are in "The frozen design" below.

## Why it is sound

**The refused condition was false.** `<mergeCells>` is required after
`<sheetData>` and the fixtures that tripped this put it there. In
`test-data/libreoffice-core/sc/qa/unit/data/xlsx/LambdaAndRelatedFunctions.xlsx`
`sheet1.xml` has `<cols>` at byte 837, `<sheetData>` at 1,027 and `<mergeCells>`
at 2,489. The refusal named a placement that the bytes contradict.

**Error identity is preserved where the condition is true.** A genuinely early
`<mergeCells>` on an unmarked worksheet still gets
`invalid("worksheet mergeCells appears before sheetData")` from the pre-mark
clause, and the mandatory materialized parser carries the identical message at
`raw/worksheet/codec.rs:509`, so an ineligible worksheet that really does place
`<mergeCells>` before `<sheetData>` is still refused with that exact string
through the fallback. The synthetic matrix case `after-sheetdata__merge-*`
rows show each of the four merge refusals arriving from the fallback with its
own established message.

**The fences are untouched.** No change to
`with_verified_decoded_reader`, to the CRC/size/source/execution fences that
changes 0363 and 0365 placed around the ineligible fallback, to
`selected_stream_limits`, to the MCE or x14ac observers, or to any typed limit.
The stream still reaches XML/MCE/x14ac EOF on every worksheet. `Scanner::finish`
is unchanged. No new `unsafe`, no new dependency, no public API change: both
`validate_merge_cells_placement` and its new sibling are private.

**ADR reading.** ADR 0005 puts container, relationship/catalog, security and
mandatory structural validation at open, and loads semantic payloads lazily;
the worksheet payload's mandatory validation is the materialized parser's, which
this change routes to rather than around. Change 0362 states the contract this
restores verbatim: "`NotEligible` is not worksheet semantic validity. The caller
MUST fall back to the eager parser after that outcome." Changes 0363 and 0365
attach the EOF condition to the *verified reader* (CRC, size, source, context),
not to the scanner's own bookkeeping, and that reader is untouched.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind/callgrind 3.26.0, `taskset -c 20`, release `--locked`. Eight agents
were building concurrently on the host throughout.

### Deterministic counts (instructions rank work, not latency)

Callgrind isolation pairs on a scratch probe (`results/change-0597/probe/`):
N fresh `SourceBackedWorkbook` opens over an in-memory counting positional
source, one `cell("H680")` each, N = 1 and N = 11, totals differenced and
divided by 10. Fixtures are the two the 0587 survey used: `real.xlsx` is
`test-data/poi/test-data/spreadsheet/Excel_file_with_trash_item.xlsx`
(`sheet1.xml` 209,931 B, MCE and x14ac markers, `<cols>` at byte 610) and
`control.xlsx` is 0587's marker-stripped derivative of it (`sheet1.xml`
194,257 B, `<cols>` at byte 421). Both are marked `NotEligible(Styles)`.

| fixture | leg | Ir per source-backed one-cell read | vs base |
| --- | --- | ---: | ---: |
| control | A, base `08d968f8e` | 220,680,657 | — |
| control | B, this change | 220,702,530 | +0.010% |
| control | C, base + the frozen gate | 81,229,237 | **−63.19%** |
| real | A, base `08d968f8e` | 1,110,380,265 | — |
| real | B, this change | 1,108,340,928 | −0.18% |
| real | C, base + the frozen gate | 954,006,724 | **−14.08%** |

**Tier: measured, one fixture pair, single deterministic run per leg.** The
landed change is instruction-neutral (+0.010% / −0.18%, code layout).

Inclusive per read at the base, control fixture:
`raw::worksheet::selected::scan` 140,852,422 Ir (63.80% of the whole child),
`SourceWorksheet::store` 76,285,164 (34.55%). 0587 reported 140.7 M / 63.6% and
76.5 M / 34.4% for the same pair through a different probe, so the two agree.
On the real fixture `selected::scan` is 152,947,518 (13.81%) and the whole read
is dominated by the MCE codec's namespace re-expansion (item XML-1,
8,414,209,833 Ir inclusive, 69.09%), which the fallback still pays. With the
gate the `selected::scan` symbol is absent from both tables.

0587's falsification criterion for XML-2 was "the post-mark share on ineligible
real worksheets is under 10%". It is 63.80% on the control and 13.81% on the
real fixture. **Not falsified.**

### Reads and bytes

Over every `.xlsx` under `test-data/` (180 packages, up to three sheets each,
twelve addresses and two ranges per sheet, a fresh `SourceBackedWorkbook` over a
counting positional source for **every single read**, so each row exercises the
cold streaming route): 4,606 read rows per leg, plus 180 sheet-catalog header
lines.

- On the 4,172 rows whose result is unchanged, logical reads and logical bytes
  are **identical** in all three legs.
- On the 434 rows the defect fix moves, reads rise by exactly nine per read
  (7→16, 10→19, 11→20, 12→21, 22→31): the materialized worksheet store that the
  spurious refusal used to prevent.
- Leg B (this change) and leg C (this change plus the gate) are **byte-identical
  transcripts**: the gate changes no value, no error, no read count and no byte
  count anywhere in this corpus.

### Paired timing

`tools/perf-baseline`, `--warmup 5 --samples 30`, order A1 B1 B2 A2 on CPU 20,
followed by four base runs (F1..F4) for the A/A floor in the same window.
Cases: `xlsx_file_selected_cell`, `xlsx_range_source_first_cell`,
`xlsx_narrow_column_range_scan` across their default shapes (8 records per run).
Full table: `results/change-0597/timing/abba-summary.txt`.

The A/A floor in this window, as `(max−min)/min` over F1..F4:

| case / shape | p50 | mean | p95 | p99 |
| --- | ---: | ---: | ---: | ---: |
| `xlsx_file_selected_cell` / medium | 0.874% | 1.148% | 3.522% | 6.191% |
| `xlsx_range_source_first_cell` / medium | 1.368% | 1.623% | 3.225% | 3.313% |
| `xlsx_range_source_first_cell` / tiny | 4.988% | 4.954% | 6.513% | 7.378% |
| `xlsx_range_source_first_cell` / dense-wide | 25.501% | 15.814% | 21.366% | 24.516% |
| `xlsx_narrow_column_range_scan` / medium | 1.176% | 5.336% | 18.367% | 18.367% |
| `xlsx_narrow_column_range_scan` / tiny | 9.091% | 7.831% | 53.846% | 69.231% |
| `xlsx_narrow_column_range_scan` / dense-wide | 24.011% | 41.918% | 113.279% | 221.268% |

Paired p50 deltas (positive = this change faster):

| case / shape | leg 1 (A1→B1) | leg 2 (A2→B2) |
| --- | ---: | ---: |
| `xlsx_file_selected_cell` / medium | +0.868% | +0.380% |
| `xlsx_range_source_first_cell` / medium | −0.686% | +0.532% |
| `xlsx_range_source_first_cell` / tiny | +0.101% | −0.186% |
| `xlsx_range_source_first_cell` / dense-wide | +0.157% | +0.899% |
| `xlsx_narrow_column_range_scan` / medium | −4.762% | 0.000% |
| `xlsx_narrow_column_range_scan` / tiny | 0.000% | 0.000% |
| `xlsx_narrow_column_range_scan` / dense-wide | +2.654% | +2.099% |

Every delta is inside that scenario's own A/A floor, in both directions. **No
timing claim is made and none is admissible from this batch.**

**Regression above 5%, reported as the briefing requires and not hidden in a
mean:** `xlsx_narrow_column_range_scan` / tiny shows −72.727% p95 and −83.333%
p99 in leg 1. That scenario times 110–220 ns, one to two timer ticks, and its
own A/A floor is 53.846% p95 and 69.231% p99; leg 2 shows 0.000% p95 and
+8.333% p99 for the same statistics. `xlsx_narrow_column_range_scan` /
dense-wide shows −11.130% p95 and −12.780% p99 in leg 1 against an A/A floor of
113.279% and 221.268%, and +10.596% / +12.652% in leg 2. Both scenarios are
below their floor and neither is attributable to this change, whose only runtime
effect is two fewer branch tests on a path the harness corpora never reach: the
generated worksheets contain no `<cols>`, no row styles and no `<mergeCells>`,
so no harness worksheet is marked ineligible at all.

## Correctness evidence

**Differential oracle, whole real corpus.** 180 `.xlsx` fixtures, 4,606 read rows
per leg, three legs, transcript sha256s in
`results/change-0597/differential/oracle-summary.md`. Base → this change:
434 rows differ, over **24 of the 180 fixtures**. Every one of them reported the
same base value, `Err(Invalid("worksheet mergeCells appears before sheetData"))`,
and nothing else in the corpus moved.

**Agreement with the mandatory reference parser.** The same (fixture, sheet,
address) grid read through the fully materialized `litchi_xlsx::Workbook`:
3,948 comparable rows. At the base the source-backed path disagrees with the
materialized parser on **372** of them, all of them carrying that one spurious
refusal. After this change it disagrees on **0**. The 434 moved rows resolve to
the materialized parser's own answer: 364 to a value, a covered merge or a
missing coordinate, and 70 to the materialized parser's own typed refusal
(`worksheet formula expression is empty`, `invalid worksheet dimension '1:5'…`,
`shared formula master at (14, 4) is not first in 'D11:D14'`). See
`differential/eager-agreement.txt` and `differential/oracle-changed-rows.tsv`.

**0541-style first-error matrix.** 60 synthetic packages
(`differential/make_matrix.py`) rewrite one small valid package's `sheet1.xml`
with a worksheet malformed at a chosen position — inside the gate prefix, at
`<cols>`, after `<cols>` before `<sheetData>`, inside `<sheetData>`, after
`<sheetData>`, and at the root — in a `<cols>`-bearing shape and a `<cols>`-free
control shape. Every case is read through the public `SourceWorksheet::cell` and
`cells`. For **this change**, `differential/matrix-{before,after}.tsv` differ
only in the rows whose base value is the spurious merge refusal; every XML, MCE,
x14ac, limit and scalar refusal keeps its exact typed variant and message. The
classes the **frozen gate** moves are listed below and are why it is frozen.

**Gates.** `cargo fmt --all --check`, `cargo clippy -p litchi-xlsx
--all-targets`, `cargo test -p litchi-xlsx`, `cargo doc -p litchi-xlsx
--no-deps`, all clean; tails in `results/change-0597/gates.txt`.
`cargo test -p litchi-xlsx` is 1,297 tests across 59 suites (986 in the library),
including the 19 in `source_backed_merge` and the six 0541 planning-error-order
guards in `source_backed_cell_values`.

## Validation preserved

Every mandatory validation that ran before this change still runs. The stream
still reaches XML/MCE/x14ac EOF on every worksheet and still reports its typed
errors with the established precedence; the verified OPC reader still drains and
CRC/size-verifies the whole part; the materialized fallback still performs the
complete worksheet validation. The one removed question was answered from state
the scanner had already stopped maintaining, and the parser that owns the
question still asks it with the same words.

## The frozen design: a bounded ineligibility pre-gate

**Mechanism.** Before `scan_stream` opens the semantic stream, read a bounded
prefix (8 KiB; the largest `<cols>` offset across the 200 worksheet parts of
this repository's 180 `.xlsx` fixtures that declare one is 1,563 B) and
`memmem`-search the window that ends at the first `<sheetData` for `<cols`
followed by a name terminator. On a hit, publish `NotEligible(Styles)` without
opening the stream; otherwise chain the prefix back onto the reader
(`Chain<&[u8], &mut dyn BufRead>` is itself `BufRead`, so no second buffer and
no copy beyond the probe window) and scan exactly as today. The probe reports no
error of its own: at EOF, at a short read or at any transport error it stops and
hands the untouched stream to the scanner, so a declined probe is a no-op.

The patch is retained at `results/change-0597/design-candidate/gate-candidate.patch`.

**Why a pre-gate rather than stopping at the first mark.** The MCE stream driver
cannot be stopped from `litchi-xlsx`. `invoke_active`
(`litchi-ooxml-common/src/mce/stream.rs:2581`) records an active-observer error,
sets `active_enabled = false` and lets the loop run to EOF: the active observer
is a diagnostic callback that an input or MCE error must be able to outrank.
Stopping at the mark therefore needs either a new "detach the observer" signal in
`litchi-ooxml-common` (out of this change's scope, and an API-shaped change to a
crate shared with other work) or a reader shim that manufactures an I/O error and
then swallows it. The pre-gate needs neither and is strictly smaller.

**The prize.** −63.19% of a source-backed one-cell read on the control fixture
and −14.08% on the real one (table above), on the 83 of 180 fixtures whose first
worksheet declares `<cols>`. Observationally the gate is a pure no-op on the real
corpus: legs B and C of the 4,606-row differential are byte-identical.

**Why it is frozen.** On the synthetic matrix the gate moves error classes that
this change is not authorized to move. For a `<cols>`-gated worksheet the
semantic stream never runs, so every XML, MCE, x14ac and limit question it used
to answer is answered instead by the materialized parser — or not at all. The
exact classes, from `differential/matrix-{before,after}.tsv`, comparing the
`cols` shape against its `nocols` control so the defect fix is factored out:

1. **Refusal becomes acceptance, twice.** `before-cols__prefix-bad-entity`
   (`topLeftCell="A&bogus;1"` on an ignored `<sheetView>`): base
   `Invalid("… unrecognized entity \`bogus\`")`, gated `Ok(Stored(Number("1")))`.
   `before-cols__prefix-undeclared-prefix` (`<zz:sheetPr xmlns:qq="urn:x"/>`):
   base `MarkupCompatibility(NonConformant("unbound prefix"))`, gated
   `Ok(Stored(Number("1")))`. In both the materialized parser — and therefore
   `litchi_xlsx::Workbook::open` — already accepts the same bytes today, so the
   gate converges the two paths rather than inventing an acceptance. It is still
   a refusal that moves, and `docs/GOAL.md` does not let this change decide that.
2. **The typed variant moves.** `in-sheetdata__sheetdata-bad-entity`:
   `MarkupCompatibility(NonConformant("custom entity"))` becomes
   `Xml(Invalid("unsupported XML entity reference '&nope;'"))`.
   `after-sheetdata__second-root`: `NonConformant("multiple roots")` becomes
   `Invalid("worksheet XML must have one SpreadsheetML worksheet root")`.
   `after-sheetdata__trailing-unclosed-root`: `NonConformant("unterminated XML")`
   becomes `Invalid("worksheet extension XML has an unterminated element")`.
3. **The message moves within one variant.** `in-sheetdata__invalid-utf8`
   (`"invalid worksheet extension XML: cannot decode input using UTF-8: … from
   index 0"` becomes `"worksheet XML is not UTF-8: … from index 339"`),
   `root__mce-ignorable-undeclared` (`"unbound Ignorable prefix"` becomes
   `"unbound Ignorable nosuchprefix"`), `before-cols__prefix-stray-lt`,
   `after-cols__post-cols-duplicate-dimension` and
   `after-sheetdata__dimension-after-sheetdata` (both
   `"worksheet has duplicate dimension elements"` becomes `"worksheet dimension
   appears after column or cell data"`).
4. **Unchanged, for the record.** Eight of the eighteen gated cases keep their
   exact typed error and message: the ill-formed-document cases
   (`prefix-mismatched-end`, `prefix-unclosed-element`, `post-cols-mismatched-end`,
   `tail-mismatched-end`, `cell-unclosed`), `prefix-bad-dimension`,
   `x14ac-bad-descent`, and the two `mc:AlternateContent` conformance cases.

The rule the design proposes, and which a frozen design record must carry before
any of this lands, is the one 0587 named: *for a worksheet the gate publishes,
the materialized parser's first error wins, exactly as it does today for the
values of every ineligible worksheet.* That rule is defensible — it is the rule
`litchi_xlsx::Workbook` already follows for the same bytes — but it relocates
where a refusal is decided, and one open question is a genuine limit question:
the semantic stream applies `selected_stream_limits` (`max_event_bytes`,
`max_input_bytes`) to the worksheet part, and a gated worksheet reaches only the
materialized path's own limits. Whether those two ceilings coincide for every
part size is not established here. Per the briefing's contract-change rule, the
gate stops at this design.

**What the design needs next, in order.** (a) Decide the limit question above.
(b) Decide whether the two acceptances in class 1 are a weakening of a defence
or a convergence onto the mandatory path; the second reading needs an ADR 0005
clarification about which reader owns malformed-input refusal for a lazily
loaded payload. (c) If both resolve, the patch and this matrix are the
implementation and its guard; if they do not, the alternative is the
`litchi-ooxml-common` observer-detach signal, which keeps every stream-side
refusal and still removes `SemanticEvent::from_element` and its name clones
(`Processor::start` is 46.06% of the base control read and the owned semantic
event is built inside `if parent_active && *active_enabled`), for a smaller and
as-yet unmeasured share.

## Limitations

- **No speedup, regression, latency, throughput, cold-cache, physical-I/O,
  allocation, peak-RSS or concurrency claim is made.** Every timing delta is
  inside its scenario's A/A floor, and the floor itself exceeds 24% at p50 on
  two of the seven scenario/shape pairs in this window.
- The instruction counts are one deterministic run per leg on two fixtures
  derived from a single 210 KB real worksheet. They rank work, not latency;
  callgrind counts `rep movsb` per byte and runs SHA-256 in software.
- The corpus-wide differential covers `SourceWorksheet::cell` and
  `SourceWorksheet::cells` only. `visit_cells` shares `cells` and is not
  separately transcribed. The source-backed *edit* routes, the snapshot readers
  and the eager `Workbook` route were not differenced, on the ground that the
  changed branch is reachable only from `raw::selected_worksheet`.
- The 434 moved rows are a behaviour change on a public read path: reads that
  refused now return values. The claim made here is only that they match the
  mandatory materialized parser on this corpus, not that no consumer depended on
  the refusal.
- The `<cols>` breadth figures count raw byte occurrences in worksheet parts;
  they do not prove each one is a schema `<cols>` element.
- The gate is measured at `08d968f8e` plus this change's fix. It was not
  measured against a leg that lacks the fix, because without the fix its own
  differential is dominated by the defect.

## Retained evidence

[`results/change-0597/README.md`](results/change-0597/README.md).
