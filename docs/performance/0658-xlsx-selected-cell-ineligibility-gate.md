# 0658: the ineligible selected-cell scan now stops at the event that settles the verdict, removing 64% of an ineligible one-cell read

Status: retained, implemented. `performance_claim: none` — the paired medians
and instruction counts below are reported as evidence, not registered as
claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred
until that goal completes; iWork is excluded.

This record lands the ineligibility gate that
[`0597`](0597-xlsx-selected-cell-ineligibility-gate.md) designed, priced and
froze, under decision 8 of
[`0652`](0652-owner-decisions-for-the-third-wave.md). It does **not** land
0597's frozen candidate; it lands a different mechanism that keeps more of the
refusal identity and removes more work. Both statements are measured below.
The decision asks for "the −63% priced on an ineligible read measured on the
landed change": it is **−64.43%** on the fixture 0597 priced at −63.19%, and
**−23.14%** on the one it priced at −14.08%. The two records measure against
different base commits (0597 against `08d968f8e`, this one against
`70d7768cc`), so each percentage is exact against its own base leg and the
comparison between them is indicative.

## What was changed

**`litchi-ooxml-common`, the MCE stream.** Two additive public items:

- `mce::ActiveFlow` — `Continue` or `Stop`, what an active observer wants the
  stream to do after the event it has just received.
- `mce::process_markup_compatibility_stream_with_stoppable_observers` — the
  existing `…_with_observers` driver with an active observer that returns
  `Result<ActiveFlow, ActiveE>`. `…_with_observers` now delegates to it with an
  adapter that always answers `Continue`, so its signature, its behaviour and
  its `StreamReport` are unchanged.

Internally `invoke_active` gained a stop flag beside the existing
`enabled`/`error` pair, and the driver ends the loop with `finish_callbacks`
once a stop has been requested — **after** the event that requested it has been
fully processed and after every failure path, so a typed XML, MCE, limit or
observer failure raised by that same event still outranks the stop.

**`litchi-xlsx`, the selected-cell read path.**

- `raw::worksheet::x14ac::capture_stream_with_stoppable_active` (crate-internal)
  composes the raw x14ac observer, the x14ac active observer and a caller
  observer that may stop. `capture_stream_with_active` is retained unchanged for
  the three callers that never stop.
- `raw::worksheet::selected::Scanner::event` returns `ActiveFlow`. It answers
  `Stop` for the event that records an ineligibility reason and for any event
  after it. `scan_stream` uses the stoppable capture; `Scanner::finish` is
  unchanged and still publishes the verdict.
- The post-mark branch of `Scanner::event` — whose merge clause change 0597
  corrected — and `validate_marked_merge_cells_placement`, which 0597 added for
  it, are removed: the scanner cannot observe a post-mark event any more, so
  both are unreachable. Those two checks asked worksheet-structure questions
  from state frozen at the mark, which is what made 0597's defect possible; the
  materialized parser answers them from complete state. The witnesses this
  moves are enumerated under "Correctness evidence".
- Module and outcome documentation now says that an ineligible scan stops at
  the marking event instead of reaching XML EOF.

## Authority

Decision 8 of 0652, in the owner's words: **"XLSX ineligible read: accept the
movement of the timing of error returns."** 0652 reads that as authorizing
"0597's gate: an ineligible selected-cell read may surface its refusal at a
different point in the read than today", and requires the implementing record
to prove that "the refusal keeps its type and text and still precedes any value
returned to the caller; the −63% priced on an ineligible read measured on the
landed change; the eligible path unchanged".

It authorizes exactly one movement: **where in the read a refusal is raised.**
It does not authorize weakening a defence, and standing trade-off 2
("correctness and safety first") forbids it. That constraint is what chose the
mechanism. 0597's frozen candidate — a bounded `memmem` pre-gate that publishes
`NotEligible(Styles)` without opening the stream — skips the *prefix* as well,
so it loses refusals the stream reaches **before** the verdict is decided; its
own matrix recorded two of them turning into acceptances. Stopping at the mark
instead keeps every byte before the verdict under exactly today's parsing and
validation, and therefore keeps every refusal located there. Where a refusal
lies after the verdict, the materialized parser the caller is already required
to fall back to raises it; that is the movement decision 8 authorizes, and every
witness is listed below.

## Breaking changes

**None.** `ActiveFlow` and
`process_markup_compatibility_stream_with_stoppable_observers` are additive.
`process_markup_compatibility_stream`,
`process_markup_compatibility_stream_with_active_observer`,
`process_markup_compatibility_stream_with_observers`, `StreamReport`,
`StreamError` and every `litchi-xlsx` public item keep their signatures and
their behaviour. `litchi-xlsx`'s `raw::selected_worksheet` surface is unchanged;
only the documented meaning of `NotEligible` moves, from "the scan reached XML
EOF" to "the scan stopped at the event that settled the verdict".

## Why it is sound

**The verdict is monotone, so nothing after the mark can change it.**
`Scanner::mark` keeps the *first* reason and ignores later ones, and
`Scanner::finish` returns `NotEligible(reason)` from `self.not_eligible` without
consulting any other state. Every event after the first mark was therefore
already incapable of changing the published outcome; the only thing it could
still do was raise a refusal.

**Every refusal before the verdict keeps its exact type and text.** The stop is
honoured only after the stopping event has been fully processed. Everything
that can refuse that event still has its say: `parse_element` (attribute
decoding, namespace expansion, duplicate attributes, every bounded-name and
attribute-byte limit), the raw observer, the MCE selection rules and the x14ac
active observer all run before the scanner's callback, and `close_start` — which
runs *after* it and can still refuse the same element — is reached too, because
the driver checks the event's `Result` before it checks the stop flag. The
stream returns that error rather than the stop. Two tests pin this — one in the MCE crate
(`a_stop_never_suppresses_a_failure_already_observed_on_that_event`) and one in
`litchi-xlsx` (`a_refusal_on_the_marking_event_outranks_the_verdict`).

**The bytes after the verdict are still fully validated — by the reader that
owns them.** An ineligible outcome obliges the caller to read the worksheet
through the materialized parser (change 0362: "`NotEligible` is not worksheet
semantic validity. The caller MUST fall back to the eager parser after that
outcome"), and `SourceWorksheet::stream_cell` does exactly that. That parser
re-reads the whole part and performs the complete mandatory validation, so no
byte of an ineligible worksheet goes unvalidated; what changes is which typed
refusal the read returns when the two parsers disagree. ADR 0005 puts container,
relationship/catalog, security and mandatory structural validation at open and
loads semantic payloads lazily; it does not state which reader owns a lazily
loaded payload's refusal, so it does not contradict the landed behaviour and is
**not amended**. After this change the answer is unambiguous and matches what
0362 always said: the materialized parser owns the worksheet payload's mandatory
validation and its first typed error, and the streaming scan owns only what it
observes before the verdict.

**No fence moved.** The XLSX read calls
`part.with_verified_decoded_reader(...)`, which reaches
`SoapberryOffice::with_verified_entry_reader` — the **non-abortable** variant,
documented as "a successful callback is still drained and fully verified". A
callback that stops early is a successful callback, so the member is still
drained to its end and its CRC and compressed/uncompressed sizes are still
verified. The source-version, execution-context, part-byte-limit and
memory-reservation fences of `with_verified_decoded_reader` are untouched, as
are the `source_version()`/`execution_check()` fences `stream_cell` runs before
falling back.

**ADR 0003: never a partial result.** The read publishes either the materialized
store's complete answer for the requested coordinate or a typed refusal; the
gate adds no intermediate state. 0642's refusal-before-visit guarantee is
preserved: `visit_cells` still resolves the whole `Selection` through
`finish_result` before the first callback, and for an ineligible worksheet that
selection is the all-or-nothing materialized store. 0642's *wording* — "the scan
still runs to EOF before `select_cells` returns" — now describes the eligible
branch only; the guarantee it supports is unchanged.

**Limits.** `selected_stream_limits` sets `max_input_bytes` and
`max_output_bytes` to the part's declared uncompressed size, which the ZIP
member verification independently enforces on every read, so neither can be the
binding defence for a stopped stream. The remaining `StreamLimits::default()`
bounds (`max_event_bytes`, attributes per event, context bytes, depth) no longer
apply to bytes after the verdict; the materialized parser's own bounds apply to
them instead. This narrows, but does not fully close, open question (a) of
0597 — see "Limitations".

**No new `unsafe`, no new dependency, no new allocation on the eligible path.**
The scanner returns a two-variant `Copy` enum derived from state it already
keeps: on an eligible worksheet the change is one extra `Option::is_some` per
semantic event and one enum value returned through an inlined closure. The
eligible path's verdicts, reads, bytes and instruction counts are unchanged
within noise (below).

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind/callgrind 3.26.0, `taskset -c 13`, release `--locked`. Eight agents
were building concurrently on the host throughout. Base `70d7768cc`; branch
`perf/0658-xlsx-selected-cell-ineligibility-gate`.

### Deterministic counts (instructions rank work, not latency)

Callgrind isolation pairs on the retained probe (`results/change-0658/probe/`):
N fresh `SourceBackedWorkbook` opens over an in-memory counting positional
source, one `cell("H680")` each, N = 1 and N = 11, totals differenced and
divided by 10. Fixtures are 0597's pair: `real.xlsx` is
`test-data/poi/test-data/spreadsheet/Excel_file_with_trash_item.xlsx`
(`sheet1.xml` 209,931 B) and `control.xlsx` is 0587's marker-stripped derivative
of it (`sheet1.xml` 194,257 B). Both are ineligible.

| fixture | Ir per source-backed one-cell read, before | after | delta |
| --- | ---: | ---: | ---: |
| control | 215,630,939 | 76,698,964 | **−64.43%** |
| real | 652,726,610 | 501,712,792 | **−23.14%** |

**Tier: measured, one fixture pair, single deterministic run per leg.** For
comparison, 0597 priced its frozen pre-gate at −63.19% and −14.08% on the same
two fixtures against its own base; this gate removes more on both.

Where the instructions went (inclusive, same isolation pairs):

| symbol | control before | control after | real before | real after |
| --- | ---: | ---: | ---: | ---: |
| `selected::scan_stream` | 140,956,112 | 292,990 | 153,241,477 | 353,599 |
| `Processor::start` | 101,740,971 | below threshold | 113,502,154 | below threshold |
| `clone_bounded_name_part` | 16,527,952 | below threshold | 17,128,561 | below threshold |
| `invoke_active` | 8,662,363 | below threshold | 9,034,926 | below threshold |
| `SourceWorksheet::store` | 71,545,278 | 71,596,784 | 496,335,518 | 496,171,167 |
| `mce::codec::process_markup_compatibility` | 8,756,434 | 8,749,152 | 310,154,795 | 310,161,586 |

"below threshold" means the symbol no longer appears in the annotation at
`callgrind_annotate --threshold=99.99`; the retained tables are in
`results/change-0658/callgrind/`. The streaming scan falls by 99.8% and the
materialized fallback — the work the caller was always going to do — is
unchanged to within 0.1%. On the control fixture `SourceWorksheet::store` is 93%
of the read after the change, and on the real one the MCE byte codec inside that
fallback (survey item XML-1) is 310,161,586 Ir of the remaining 501,712,792:
that is what the next change in this area would have to attack.

### Corpus reach of the gate

Over the 326 worksheet parts of the 180 `.xlsx` packages under `test-data/`
(first three parts per package), read through
`raw::selected_worksheet::scan` with a counting `BufRead`:

- **Every one of the 326 parts is ineligible** — 325 `UnsupportedStructure`,
  one `RichInlineText` — and the verdict is **identical** on both legs for all
  326. The selected-cell fast path is not reached by any real fixture in this
  repository, so the gate governs every source-backed selected read of this
  corpus.
- Bytes pulled from the worksheet reader, aggregated over the 326 parts:
  **9,324,342 → 795,708 (−91.47%)**. The counter is quantized to the stream's
  8 KiB fill, so the 288 parts smaller than one fill still show 100% whatever
  the gate does. Every one of the 38 parts larger than one fill (9,149 B to
  3,382,556 B) pulls **exactly one 8,192-byte fill** and no more — 0.24% of the
  largest part, 89.54% of the smallest of them.

### Reads and bytes at the source

The corpus differential (next section) records the logical `read_at` count and
byte count of every one of its 4,606 reads. They are **identical on both legs**:
the gate removes XML parsing, not source I/O, because the verified OPC reader
still drains and verifies the whole member.

### Paired timing

Two instruments, both on CPU 13, both with an A/A floor measured in the same
window as four extra base runs.

**(a) The ineligible read itself**, through the retained probe: 20 fresh
`SourceBackedWorkbook` opens and one `cell("H680")` each per sample, 5 warmup
samples and 30 measured samples per leg, order A1 B1 B2 A2 then F1..F4. Figures
are nanoseconds per read; positive = this change faster.

| fixture | stat | A1 base | B1 chg | leg 1 | A2 base | B2 chg | leg 2 | A/A floor |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| control | p50 | 15,818,189 | 4,337,611 | **+72.58%** | 14,148,077 | 4,358,577 | **+69.19%** | 0.66% |
| control | mean | 19,929,250 | 4,342,721 | +78.21% | 14,159,590 | 4,362,587 | +69.19% | 25.29% |
| control | p95 | 49,632,513 | 4,373,427 | +91.19% | 14,355,228 | 4,393,586 | +69.39% | 120.56% |
| control | p99 | 53,340,593 | 4,389,869 | +91.77% | 14,367,109 | 4,421,449 | +69.23% | 252.87% |
| real | p50 | 37,100,722 | 26,003,787 | **+29.91%** | 36,008,741 | 26,663,959 | **+25.95%** | 8.12% |
| real | mean | 40,406,517 | 26,132,022 | +35.33% | 35,983,966 | 26,562,729 | +26.18% | 18.40% |
| real | p95 | 71,229,718 | 26,722,921 | +62.48% | 36,283,316 | 27,300,845 | +24.76% | 53.99% |
| real | p99 | 71,457,016 | 26,970,702 | +62.26% | 36,493,383 | 27,460,037 | +24.75% | 47.74% |

The control fixture's p50 A/A floor is 0.66% and the effect is 69–73%, two
orders of magnitude above it, in both directions of the ABBA pair. The real
fixture's p50 floor is **8.12%, above the 5% threshold**, so as the briefing
requires this record says so and rests the magnitude on the instruction counts;
the effect there is 26–30%, about three times the floor, and agrees with the
−23.14% instruction figure. Every mean/p95/p99 floor in this window is large
(18% to 253%), which is what eight agents building concurrently on the host
looks like; those columns are reported, not relied on.

**(b) The eligible controls**, `tools/perf-baseline` with
`--warmup 5 --samples 30` over `xlsx_file_selected_cell`,
`xlsx_range_source_first_cell` and `xlsx_narrow_column_range_scan` across their
default shapes (7 scenario/shape pairs), same ABBA order and the same four
floor runs. **Nothing moves outside its own floor.** The A/A p50 floor in this
window is 48.260% for `xlsx_file_selected_cell` / medium, 89.368% for
`xlsx_range_source_first_cell` / dense-wide, 19.767% and 8.520% for two
`xlsx_narrow_column_range_scan` shapes, and 3.203%, 1.600% and 0.000% for the
other three — **four of the seven exceed 5% at p50**, so this batch admits no
timing conclusion about the eligible path in either direction. The largest p50
deltas observed are −1.747% and +5.111%, both far inside their own floors.

**Regressions above 5%, reported as the briefing requires and not hidden in a
mean:** `xlsx_file_selected_cell` / medium shows −60.420% p95 and −89.845% p99
in leg 2, against an A/A floor of 285.274% and 309.808% for the same statistics;
leg 1 shows −2.701% and −3.409% for them. `xlsx_narrow_column_range_scan` /
tiny shows −9.091% p95 and −8.333% p99 in leg 1 (floor 90.909% and 366.667%,
and leg 2 shows +20.000% and +71.429%), and `xlsx_narrow_column_range_scan` /
dense-wide −1.521% p50 in leg 1 (floor 8.520%, leg 2 +2.321%). All are inside
their own floors and none is attributable to this change, whose only effect on
an eligible worksheet is one `ActiveFlow::Continue` returned per semantic
event. The full tables are in `results/change-0658/timing/`.

## Correctness evidence

**Corpus differential, whole real corpus.** 180 `.xlsx` packages, up to three
sheets each, twelve addresses and two ranges per sheet, a fresh
`SourceBackedWorkbook` over a counting positional source for **every single
read** so each row exercises the cold streaming route: 4,606 read rows plus 180
sheet-catalog lines, 4,786 lines per leg. The two transcripts are
**byte-identical**, sha256
`23abae8318fb74c5c38a1ddc3d70f662b16b465ca8e7763abec2e89a906b227a` on both legs:
same values, same refusals (type and text), same fallback outcomes, same logical
reads, same bytes. `differential/oracle-summary.md`.

**Verdict census.** The 326-part table above: every verdict identical, and the
bytes the scan pulls reduced. `differential/verdicts-{before,after}.tsv`.

**0541-style first-error matrix.** 60 synthetic packages generated by
`differential/make_matrix.py` (0597's generator, retargeted): one small valid
package's `sheet1.xml` rewritten with a worksheet malformed at a chosen position
— inside the prefix, at `<cols>`, after `<cols>` before `<sheetData>`, inside
`<sheetData>`, after `<sheetData>`, and at the root — in a `<cols>`-bearing
shape and a `<cols>`-free control. Each package is read through the public
`SourceWorksheet::cell` and `cells` at 15 coordinates: 900 rows per leg. **210
rows over 15 of the 60 packages move**; the other 45 packages, 690 rows, are
identical. Every moved row is one of these witnesses, and each is a package that
is *both* ineligible *and* malformed after its mark:

| witness (package) | before | after | class |
| --- | --- | --- | --- |
| `before-cols__prefix-bad-entity` (both shapes) | `Invalid("invalid worksheet extension XML: at 2..7: unrecognized entity \`bogus\`")` | the value, or `Missing`/`Covered` | refusal becomes acceptance |
| `after-sheetdata__second-root` (both shapes) | `MarkupCompatibility(NonConformant("multiple roots"))` | `Invalid("worksheet XML must have one SpreadsheetML worksheet root")` | typed variant moves |
| `after-sheetdata__trailing-unclosed-root` (both shapes) | `MarkupCompatibility(NonConformant("unterminated XML"))` | `Invalid("worksheet extension XML has an unterminated element")` | typed variant moves |
| `in-sheetdata__sheetdata-bad-entity` (both shapes) | `MarkupCompatibility(NonConformant("custom entity"))` | `Xml(Invalid("unsupported XML entity reference '&nope;'"))` | typed variant moves |
| `after-cols__post-cols-duplicate-dimension` (`cols`) | `Invalid("worksheet has duplicate dimension elements")` | `Invalid("worksheet dimension appears after column or cell data")` | message moves |
| `after-sheetdata__dimension-after-sheetdata` (both shapes) | `Invalid("worksheet has duplicate dimension elements")` | `Invalid("worksheet dimension appears after column or cell data")` | message moves |
| `before-cols__prefix-stray-lt` (both shapes) | `Invalid("… position 13: attribute key must be directly followed by \`=\` or space")` | `Invalid("… ill-formed document: expected \`</>\`, but \`</worksheet>\` was found")` | message moves |
| `in-sheetdata__invalid-utf8` (both shapes) | `Invalid("invalid worksheet extension XML: cannot decode input using UTF-8: … from index 0")` | `Invalid("worksheet XML is not UTF-8: … from index 339")` | message moves |

Two comparisons matter.

1. **Against 0597's frozen candidate on the same matrix.** For the `<cols>`
   shape — the only shape 0597's pre-gate acted on — the pre-gate fired on 18
   of the 30 packages and moved 10 of them; this gate fires on all 30 and moves
   8, and the two it preserves are
   `before-cols__prefix-undeclared-prefix` (which the pre-gate turned from
   `MarkupCompatibility(NonConformant("unbound prefix"))` into an acceptance)
   and `root__mce-ignorable-undeclared`. **The gate halves 0597's
   refusal-becomes-acceptance class.** The one that remains,
   `prefix-bad-entity`, is malformed inside a *child* of the element that marks
   the worksheet (`<sheetView topLeftCell="A&bogus;1"/>` inside `<sheetViews>`),
   so the mark precedes it by one event.
2. **What the surviving acceptance means.** For that package the materialized
   parser accepts the same bytes, and accepted them before this change too, so
   `litchi_xlsx::Workbook::open` — the mandatory public read — already returned
   a value for it. No input that litchi refused is now admitted; one path's
   answer converges onto the answer the owning parser always gave. It is still
   a refusal that this path no longer raises, and it is stated here, not hidden.

**Mutation set: every way a worksheet can become ineligible.**
`change_0658_ineligibility_gate_tests::every_ineligibility_reason_stops_the_stream_at_its_mark`
builds one worksheet per `NotEligibleReason` variant — all nine:
`UnsupportedStructure`, `MergeSemantics`, `SharedStrings`, `Styles`,
`FormulaSemantics`, `RichInlineText`, `UnsupportedCellType`, `Ordering`,
`GeneralReference` — each followed by a long inert comment tail. For each, both
`scan` and `scan_range` publish the expected verdict, both consume less than
half the part, and the materialized parser's answer for the same bytes is
pinned exactly (four of the nine are refusals, with their messages). A companion
test asserts that an eligible worksheet still consumes the part to the last byte.

**Tests updated because the contract moved.** Four established tests pinned "a
typed error later in the stream outranks the ineligible verdict". Each is
rewritten to pin the new contract *and* the materialized parser's answer, so the
movement is visible in the suite rather than silent:

- `streaming_0364_dependency_metadata_keeps_late_malformed_xml_primary` →
  `change_0658_late_malformed_xml_after_the_mark_moves_to_the_materialized_parser`.
- `streaming_0366_keeps_invalid_references_as_typed_errors_without_early_publication`
  keeps every pre-mark case and now asserts that the post-mark `<future>&#xZZ;</future>`
  case takes the materialized parser's answer, which is acceptance; a new
  pre-mark case pins that an invalid reference before the verdict is still a
  typed stream error.
- `streaming_0367_rejects_invalid_merge_references_and_placement` keeps its
  seven unmarked cases unchanged and moves the `<hyperlinks/>` case to the
  materialized parser, **which raises the identical message**.
- `change_0597_marked_worksheet_still_refuses_misplaced_merge_markup` →
  `change_0658_marked_worksheet_hands_merge_placement_to_the_materialized_parser`:
  the "after a schema successor" case keeps its identical message from the
  materialized parser; the duplicate-`<mergeCells>` case now names the misplaced
  `<cols>` that marked the worksheet, which is the earlier structural error.

**New tests for the signal.** Six in `litchi-ooxml-common`
(`streaming_0658_active_stop_tests`): a continuing observer is indistinguishable
from the plain entry point (same report, same raw and active event sequences); a
stop ends the stream at that event with success and a report counting only the
processed events; a stopped stream does not run the end-of-document checks; a
stop never suppresses a failure already observed on that event; an observer
error before a stop still surfaces; and the movement itself is pinned — a
document malformed after the stopping event is refused without a stop and
accepted with one.

**Gates.** `cargo fmt --all --check`; `cargo clippy -p litchi-ooxml-common -p
litchi-xlsx --all-targets`; `cargo test -p litchi-ooxml-common`; `cargo test -p
litchi-xlsx`; `cargo doc -p litchi-ooxml-common -p litchi-xlsx --no-deps`; the
consumer crates of `litchi-ooxml-common` (`litchi-docx`, `litchi-pptx`,
`litchi-opc`); `cargo test -p litchi --features docx,xlsx,pptx,xls`; `cargo
test` in `tools/perf-baseline`; `python3 tools/non_iwork_gate.py verify`. Tails
in `results/change-0658/gates.txt`.

## Validation preserved

Every validation that ran before the verdict still runs, in the same order, with
the same typed errors. The verified OPC reader still drains and CRC/size-verifies
the whole part. The materialized parser still performs the complete mandatory
worksheet validation on every ineligible worksheet, which on this corpus is all
of them. No limit was raised or removed and no `unsafe` was added.

One precision, because the phrase "no malformed input is newly admitted" would
be too strong: the single acceptance witness above is a package that
`litchi_xlsx::Workbook::open` — the mandatory public read — accepted before this
change and accepts now. What changed is that `SourceWorksheet::cell` no longer
disagrees with it. No input that *every* public entry point refused is admitted
by any of them after this change.

## Limitations

- **No speedup, latency, throughput, cold-cache, physical-I/O, peak-RSS or
  concurrency claim is made**, beyond the paired medians reported above beside
  their A/A floor.
- The instruction counts are one deterministic run per leg on two fixtures
  derived from a single 210 KB real worksheet. They rank work, not latency;
  callgrind counts `rep movsb` per byte (0604) and runs SHA-256 in software
  (0649).
- **Refusal identity moves for a worksheet that is both ineligible and malformed
  after its mark.** The eight witnesses above are the complete list over this
  matrix; one of them turns a refusal into the acceptance the mandatory parser
  already gave. This is the movement decision 8 authorizes; it is not claimed
  that the matrix is exhaustive over all malformed inputs.
- **The stream's per-event bounds no longer apply after the verdict.**
  `max_input_bytes`/`max_output_bytes` are the part's declared size and are
  independently enforced by ZIP member verification, but `max_event_bytes`, the
  attribute and context-byte bounds and the stream depth bound now apply only up
  to the marking event; after it, the materialized parser's own bounds apply.
  Whether those two families of ceilings coincide for every part shape is still
  not established — this narrows 0597's open question (a) rather than closing it.
- The corpus differential covers `SourceWorksheet::cell` and
  `SourceWorksheet::cells`. `visit_cells` shares `cells`'s selection and is not
  separately transcribed. The source-backed *edit* routes and the eager
  `Workbook` route were not differenced, on the ground that the changed code is
  reachable only from `raw::selected_worksheet`.
- **Every worksheet part in this repository's corpus is ineligible.** The
  eligible path is therefore exercised here only by unit tests and by the
  harness's generated corpora, and this record did not verify which verdict the
  harness's own worksheets take. The "eligible path unchanged" evidence is
  therefore the verdict census (identical on all 326 parts), the corpus
  differential (byte-identical), the unit test that pins an eligible scan to the
  last byte of its part, and the code shape (the stop can only be requested
  after a mark) — not a real-corpus eligible differential and not the harness
  timing, whose A/A floor in this window exceeds 5% at p50 on four of its seven
  scenario/shape pairs and therefore concludes nothing in either direction.
- An alternative considered and rejected: stop at the *end* of the marking
  element instead of at the marking event, which would also preserve the
  `prefix-bad-entity` witness. It was rejected because a worksheet marked at its
  own `<worksheet>` root element — an unmodelled root attribute is enough —
  would then read the entire part and gain nothing, so the prize would depend on
  where the mark falls. The retained candidate has one rule and one cost model.

## Retained evidence

[`results/change-0658/README.md`](results/change-0658/README.md).
