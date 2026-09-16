# 0635: one pass over each cell reference and one ampersand probe per part cut the XLSX fact builder by 35-36%; the snapshot chain it would feed is unreachable

Status: retained, implemented in `litchi-xlsx`, with one designed part **rejected**
and kept as a patch in the evidence packet. `performance_claim: none` — the counts
and paired medians below are reported as evidence, not registered as a claim. Every
change is **value-identical**: every scenario's published package keeps its exact
`output_sha256`, and change 0622's differential oracle over every worksheet part of
every `.xlsx` fixture is unchanged and still passes.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This is queue item 17 of
[`0630-queue-refresh-after-the-first-wave.md`](0630-queue-refresh-after-the-first-wave.md)
("the builder is 15-18% of planning; snapshot chains drop facts after the first
commit; `<f>` worksheets declined", disposition *measurement, then a follow-on*)
together with item **XLSX-5** of
[`0587-remaining-opportunity-survey.md`](0587-remaining-opportunity-survey.md)
("`capture_auxiliary_source` parses the whole stylesheet at every planning for a
count … unknown size, invisible on the harness's minimal styles"). It follows
change [0622](0622-xlsx-compact-source-facts.md), which introduced the fact
builder, and change [0551](changes/0551-xlsx-layout-handoff-feasibility.md), which
froze its design.

## What was changed

Three parts were designed. **Two are retained**; the third is rejected on its own
measurement and retained as a patch.

### Retained: the fact builder's per-element and per-cell cost

`crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/facts.rs`. Change 0622's
builder observes every event of the shared planning traversal. Four changes remove
work from that observation without changing one element it admits or declines:

* **One ampersand probe per part instead of one per start tag.** `FactsBuilder::new`
  now takes the worksheet bytes and answers `memchr(b'&', content).is_none()` once.
  A worksheet with no `&` anywhere has no start tag with one either, so
  `FactsBuilder::element` skips its per-tag probe entirely; a worksheet that carries
  an `&` keeps the per-tag probe unchanged. The span is still range-checked against
  `content` on every element, because every retained offset is taken from it.
* **One pass over `<c r="…">` instead of four.** `cell_column` replaces the
  `is_ascii_alphanumeric` sweep, the `str::from_utf8` validation, `parse_a1` (which
  parses the column letters *and* the row digits and formats a diagnostic `String`
  on every failure) and the separate `reference_row != row` comparison. The builder
  already holds the row number from the `<row r="…">` it accepted, so the reference's
  row part only has to agree with it.
* **One pass over `<row r="…">` instead of three.** `row_number` likewise replaces
  the `is_ascii_digit` sweep, the UTF-8 validation and `parse_one_based_row`.
* **The namespace test alone where the name test was a tautology.**
  `is_spreadsheetml_name(namespace, name, local)` was always called with the
  element's *own* local name, so its `name.local_name() == local_name` half compared
  a value with itself while recomputing a colon scan. `is_spreadsheetml` keeps the
  namespace half. `raw_attribute` likewise compares the whole attribute key against
  a colon-free name in place of `prefix().is_none() && local_name() == name`, which
  is the same test with two fewer colon scans.

### Retained: the stylesheet parse at planning (XLSX-5)

`crates/litchi-xlsx/src/raw/styles.rs`. `raw::styles::parse` already retains nothing
but a count — `Catalog` is one `u32` — so there is no catalog to avoid building;
what XLSX-5 measures is the traversal. That traversal called
`Event::into_owned()` and `NsReader::resolver().clone()` on **every event**: the
first copies the whole start tag, name and attributes included; the second clones a
`NamespaceResolver`, which is a `Vec<u8>` of prefix and URI bytes plus a
`Vec<NamespaceBinding>` of the bindings in scope. Three heap allocations, three
copies and three frees per event, all discarded unread. `NsReader<&[u8]>` yields
events borrowed from the input and defers its namespace pop until the next read, so
the reader's own resolver answers every event of the iteration. The loop now reads
`reader.read_event()` and `reader.resolver().resolve_event(event)` directly — the
same driver, the same events, the same checks, the same messages, in the same order.

### Rejected: carrying the facts across a snapshot chain

Designed, implemented, tested and then **not landed**. The implementation is
[`results/change-0635/patch/0635-chain-facts.patch`](results/change-0635/patch/0635-chain-facts.patch);
"Why the chain part lost" below has the measurement and the reachability analysis.

No public API, no error type, no limit, no output byte and no `unsafe` changed.

## Why it is sound

**The builder admits and declines exactly what it did.** Each replaced test is
equivalent to the test it replaces, and each equivalence is pinned by a test:

* `row_number` against `parse_one_based_row`, and `cell_column` against `parse_a1`
  followed by the row comparison, over 484 byte strings built from an alphabet that
  covers every branch of both — the letter/digit split, both empty halves, a leading
  digit, a trailing letter, lowercase, separators, an ampersand, a colon, the column
  ceiling `XFD`/`XFE`, the row ceiling `1048576`/`1048577`, a leading zero and a
  `u32` overflow — each at three row numbers
  (`change_0635_row_number_agrees_with_the_shared_row_parser`,
  `change_0635_cell_column_agrees_with_the_shared_a1_parser`). The fused parsers
  accept only `[A-Za-z]+[0-9]+`, which is a subset of the alphanumeric alphabet the
  separate sweep enforced, so the property change 0622 relies on — that
  attribute-value normalization is the identity on these values — is preserved by
  construction rather than by a second pass.
* The whole-part ampersand probe is the same predicate: a start tag can only contain
  an `&` if the part does. When the part contains one anywhere, every tag is probed
  exactly as before.
* `is_spreadsheetml_name(namespace, name, local)` where `local` *is*
  `name.local_name()` reduces to its namespace arm; `attribute.key.as_ref() == name`
  for a colon-free `name` is `prefix().is_none() && local_name() == name`, because a
  key without a colon is its own local name and has no prefix, and a key with one can
  never equal a colon-free name. A `debug_assert!` pins the colon-free precondition.

**The styles traversal is the same traversal.** Nothing about the event sequence,
the namespace resolution, the `cellXfs` invariants, `MAX_CELL_FORMATS`, the
declared-count comparison or any diagnostic string changes: only the two per-event
heap copies whose results were never read. `quick_xml::reader::Config::default()`
keeps `check_end_names: true`, which is what supplies the mismatched-tag refusal, and
`parse` never touches `config_mut()`. The decoder is still taken *before* each
`read_event`, so a build that enables quick-xml's `encoding` feature sees the same
decoder for the same event as before.

**ADR reading.** ADR 0005 keeps mandatory structural validation where it is: no
validation moves, none is added and none is removed — both parts remove work whose
result was discarded or recomputed. ADR 0003's bounded resources are unaffected: the
retained state is byte-for-byte the same `SourceFacts`
(`change_0622_retained_records_stay_compact` still pins 16 bytes per cell and 32 per
row), and the styles parse now allocates *less*. ADR 0006's preservation contract is
untouched: the output bytes are identical, proved by the oracle and by the harness's
own `output_sha256`. Proposed ADRs 0030 and 0031 are not cited.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind/callgrind 3.26.0, `taskset -c 11`, release `--release --locked`. Eight
agents were building on the host throughout. Base `c7326f680`; branch
`perf/0635-xlsx-facts-builder-and-chains`.

### Value identity

`output_sha256` from `tools/perf-baseline` is identical on both legs for **34 of 34**
(case, shape) pairs: `one_edit_save`, `one_percent_edit_save`, `batch_edit_save`,
`multi_sheet_edit_save`, the two managed variants, `cell_clear_edit_save` and
`cell_remove_edit_save` over `medium`, `dense-sparse`, `noncompact` and
`vendor-extension`, plus the two 0601 producer-shaped edit selectors
(`results/change-0635/differential/output-hashes.txt`).

### Deterministic counts

Callgrind isolation pairs, N = 1 and N = 11 harness samples, totals differenced and
divided by 10, one deterministic run per leg. Full table, every symbol and all three
legs: `results/change-0635/counts/instruction-summary.txt`.

**The fact builder** (`FactsBuilder::observe`, inclusive Ir per operation):

| case | shape | before | after | delta |
| --- | --- | ---: | ---: | ---: |
| one edit | medium | 4,059,686 | 2,620,777 | **−35.44%** |
| one edit | dense-sparse | 28,947,191 | 18,538,306 | **−35.96%** |
| one edit | noncompact | 4,440,877 | 2,885,399 | **−35.03%** |
| one percent | medium | 16,237,087 | 10,469,982 | **−35.52%** |
| one percent | dense-sparse | 31,641,673 | 20,272,036 | **−35.93%** |
| managed one edit | medium | 4,064,158 | 2,617,901 | **−35.59%** |
| producer-medium edit | — | 1,799,277 | 1,161,012 | **−35.47%** |
| producer-dense edit | — | 28,930,653 | 18,460,231 | **−36.19%** |

The reduction is **−35.0% to −36.2% on every scenario measured**, and its components
are equally uniform: `FactsBuilder::element` falls 41.8-43.4%, `FactsBuilder::cell`
36.9-37.7% and `raw_attribute` 24.6-27.4%. Change 0622 reported the builder as the *rise* it caused in planning — 15.0-17.7% —
which is its cost over the traversal without it. On the same denominator the builder
was **16.6-17.1%** of the traversal before this change and is **10.7-11.0%** after; as
a share of planning as a whole it falls from 13.2-14.3% to 8.9-9.7%.

**Planning** (`Snapshot::from_source_selected`) and the **whole harness iteration**,
which also includes the per-sample output verification the timed interval excludes and
is therefore the conservative bound:

| case | shape | planning before | planning after | delta | iteration delta |
| --- | --- | ---: | ---: | ---: | ---: |
| one edit | medium | 28,438,972 | 26,986,772 | −5.11% | −0.13% |
| one edit | dense-sparse | 198,167,407 | 187,746,474 | −5.26% | −0.66% |
| one edit | noncompact | 30,985,516 | 29,426,923 | −5.03% | −0.19% |
| one percent | medium | 115,218,171 | 109,420,310 | −5.03% | −0.39% |
| one percent | dense-sparse | 218,095,112 | 206,754,162 | −5.20% | −0.63% |
| managed one edit | medium | 28,924,472 | 27,592,381 | −4.61% | −0.16% |
| producer-medium edit | — | 14,331,081 | 13,716,308 | −4.29% | **−1.63%** |
| producer-dense edit | — | 218,486,540 | 208,133,601 | −4.74% | **−2.19%** |

**No scenario got worse.** The two producer-shaped selectors show the largest
whole-iteration effect because they are not dominated by the 4 MB of media the
cell-CRUD corpora carry: on `producer-dense` the builder is 5.9% of the whole
iteration before this change and 3.8% after.

The commit is untouched — `rewrite_value_only_with_provenance` moves by at most 0.1%
in every case — and so is the rebind: `from_rewritten_value_source` moves by
−0.04% to +0.09%, which is the control that shows the chain capture is absent from
the retained change.

**Tier: measured**, one deterministic run per leg, harness corpora only.

### The stylesheet count (XLSX-5)

`capture_auxiliary_source`, which reads the styles and theme parts and counts the
cell formats, falls **−6.89% to −7.00%** in four of the five cases where the symbol is
attributed (242,822 → 225,830 Ir per operation on `medium`); on `producer-medium` the
same symbol is reported 11.5% higher, and its before values differ between cases that
run the identical code, so the instruction attribution for this symbol is not stable
enough to carry the result.

**The allocation figure is.** `litchi-perf-baseline-alloc` with the change-0538 phase
regions, five samples per leg, maximum per region: the `plan` region loses **exactly
80 allocation calls and exactly 6,457 allocated bytes** in every one of the six
(case, shape) combinations measured — 17,257 → 17,177 and 3,195,298 → 3,188,841 on
`medium` one-edit, 116,942 → 116,862 and 14,231,288 → 14,224,831 on `dense-sparse`,
129,491 → 129,411 on `dense-sparse` one-percent, and so on. That is the 603-byte,
26-event harness stylesheet losing about three heap copies per event. **Staging,
`commit_core`, commit and publication are byte-identical in every row**, and
`region_peak_live_bytes` moves by three bytes or less everywhere
(`results/change-0635/alloc/allocation-summary.txt`).

### The rejected chain capture, priced

Measured as its own leg (the retained change plus
`patch/0635-chain-facts.patch`) against the retained change:

| case | shape | `from_rewritten_value_source` | whole iteration |
| --- | --- | ---: | ---: |
| one edit | medium | +20.75% | +0.20% |
| one edit | dense-sparse | +21.35% | +1.07% |
| producer-medium edit | — | +10.39% | **+3.07%** |
| producer-dense edit | — | +16.68% | **+4.06%** |

### Paired timing

`tools/perf-baseline`, `--warmup 5 --samples 30`, order A1 B1 B2 A2 on CPU 11,
followed by four before-only runs F1..F4 for the A/A floor in the same window. Five
source-backed cell-value cases over three shapes plus the two 0601 producer-shaped
edit selectors — **17 scenarios**. Full table with mean, p95, p99 and the raw p50
nanoseconds: `results/change-0635/timing/abba-summary.txt`.

**The floor splits the matrix in two.** With eight agents on the host, the A/A floor
over F1..F4 is 1.84% to 103.5% at p50 on the fifteen cell-CRUD scenarios — those
corpora carry 4 MB of incompressible media per iteration and their publication
dominates — but **0.45% and 0.59% at p50 on the two producer-shaped scenarios**,
which are 40 to 490 times cheaper per iteration and hold still. Only the second pair
can carry a timing result; the rest are reported because the method requires every
scenario to be reported, not because they decide anything.

| case | statistic | leg 1 (A1→B1) | leg 2 (A2→B2) | A/A floor |
| --- | --- | ---: | ---: | ---: |
| producer-dense edit | p50 | **+2.89%** | **+3.35%** | 0.45% |
| producer-dense edit | mean | +2.57% | +3.36% | 0.46% |
| producer-dense edit | p95 | +2.59% | +3.73% | 0.42% |
| producer-dense edit | p99 | −0.07% | +3.70% | 0.92% |
| producer-medium edit | p50 | **+2.37%** | **+1.84%** | 0.59% |
| producer-medium edit | mean | +2.92% | +1.63% | 0.54% |
| producer-medium edit | p95 | +6.58% | −0.40% | 1.06% |
| producer-medium edit | p99 | +7.73% | +0.75% | 2.95% |

Raw p50: producer-dense 26.51 ms → 25.74 ms (leg 1) and 26.30 ms → 25.42 ms (leg 2);
producer-medium 2.086 ms → 2.036 ms and 2.065 ms → 2.027 ms. The timed interval
excludes the per-iteration verification the instruction count includes, so a timed
delta larger than the −1.63% and −2.19% whole-iteration instruction deltas is what
this change predicts.

**Every scenario, including the ones that got worse.** At p50, 16 of the 17 scenarios
are faster in *both* legs; the seventeenth, managed one-edit on `noncompact`, is
+7.29% in leg 1 and **−1.43%** in leg 2 against its own 3.28% floor. At the tails the
cell-CRUD scenarios are meaningless in both directions — managed one-percent on
`dense-sparse` reads −32.57% at mean in leg 1 and +5.26% in leg 2 against a 22.86%
floor, and −248.83% at p95 against a 90.22% floor, which is one outlier sample in A1,
not a regression. The two negative producer numbers are producer-medium p95 leg 2
(−0.40% against a 1.06% floor) and producer-dense p99 leg 1 (−0.07% against a 0.92%
floor); both are inside their own floor and both have a clearly positive partner leg.

**Tier: measured** on the two producer-shaped scenarios, whose floor is under 0.6%;
**bounded, not measured** on the fifteen cell-CRUD scenarios, whose floor exceeds the
effect. No claim is registered. **The counts carry this record's result**; the timing
confirms it where the host was quiet enough to see it.

## Correctness evidence

**The fused reference parsers are compared against the shared ones.**
`change_0635_row_number_agrees_with_the_shared_row_parser` and
`change_0635_cell_column_agrees_with_the_shared_a1_parser`
(`crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/facts.rs`) build 484 byte
strings from a 22-piece alphabet — `""`, `A`, `a`, `Z`, `AA`, `XFD`, `XFE`, `ZZZ`,
`ZZZZ`, `ZZZZZZZ`, `0`, `1`, `01`, `9`, `10`, `1048576`, `1048577`, `99999999`,
`4294967296`, `-`, `:`, `&` — and require `row_number` to agree with
`parse_one_based_row` and `cell_column` to agree with `parse_a1` followed by the row
comparison, at three row numbers each. They agree on every one
(`results/change-0635/differential/parser-equivalence.log`).

**Change 0622's differential oracle is unchanged and reports the same numbers.** Its
five tests still pass, and its corpus funnel line is identical to the one in change
0622's record: **391 worksheet parts, 207 admitted to the shared traversal, 1 accepted
by the value-only planning validator, 1 publishing facts, 2,790 comparisons** in the
bounded run and **2,840** with `LITCHI_0622_FULL_ORACLE=1` in release — the same 2,840
change 0622 recorded. Since the oracle drives every cell of every worksheet through
both commit routes and compares the complete output bytes, the readback provenance
and the typed error, an identical funnel and an identical comparison count is direct
evidence that this change moved neither which worksheets publish facts nor what those
facts describe (`results/change-0635/differential/oracle.log`,
`differential/oracle-full.log`).

**Value identity end to end.** 34 of 34 (case, shape) pairs publish a package with an
identical sha256.

**Suites.** `cargo test -p litchi-xlsx` is **1,310 tests in 60 suites**, all passing
(1,308 at the base plus the two new equivalence tests).
`cargo test -p litchi --features docx,xlsx,pptx,xls` is **266 tests in 26 suites**, all
passing.

**Gates.** `cargo fmt --all --check`, `cargo clippy -p litchi-xlsx --all-targets`,
`cargo test -p litchi-xlsx`, `cargo doc -p litchi-xlsx --no-deps`,
`cargo test -p litchi --features docx,xlsx,pptx,xls`, `cargo check -p litchi
--all-targets` and `cargo test` in `tools/perf-baseline` — tails in
[`results/change-0635/gates.txt`](results/change-0635/gates.txt). `cargo check -p
litchi --all-targets` emits three dead-code warnings in
`crates/litchi/tests/unexpected_format.rs`; they are **pre-existing at the base
commit** and were reproduced on the untouched before checkout.

## Why the chain part lost

The design, implemented in full and kept as
[`results/change-0635/patch/0635-chain-facts.patch`](results/change-0635/patch/0635-chain-facts.patch):
`Snapshot::from_rewritten_source`, `from_rewritten_value_source` and
`from_visibility_rewrite` each already validate the rewritten worksheet with
`validation::worksheet_xml`. Attaching change 0622's builder to *that* walk — the
way `worksheet_xml_and_parse_source` attaches it to the planning walk — makes a
rebound snapshot carry the same facts a fresh load of those bytes would carry, at
the cost of the builder alone and with no extra traversal. The retained spans index
into the `Vec`'s heap buffer, which does not move when the `Vec` is moved into an
`Arc`; `rebound_worksheet` checks that with `SourceFacts::describes` rather than
assuming it, and drops the facts on any disagreement.

It works. The patch adds four tests to change 0622's oracle: the two walks are
compared capture against capture, field by field, over all 13 shapes on which both
publish facts (they agree on every one, and neither walk publishes where the other
does not); the second commit of a chain is driven through the whole 0622 comparison
— every cell, a 256-cell batch, six insert and remove coordinates — on the output of
a first commit for every admitted shape and for the dense grid, with the route
counters proving the fact route ran; and the same chain comparison is swept over
every worksheet part of every `.xlsx` in `test-data/`. `cargo test -p litchi-xlsx`
is 1,312 tests, all passing, with the patch applied.

**It lost on reachability.** No public API seeds a value-only edit from a
post-commit snapshot, so the facts a rebound snapshot carries are never read:

* `SourceBackedEditor::edit` and `edit_sheets` build their `before` from
  `Snapshot::load_source_backed` / `MultiSnapshot::load_source_backed` against the
  immutable source-backed package. Every plan is a *fresh* load and already captures
  facts by the change-0622 path.
* `MultiSourceEdit::new` is private. `MultiCommit::into_parts` hands out the
  post-commit `MultiSnapshot`, but nothing public consumes one as the `before` of
  another edit.
* `publish_multi_commit_to_stream(self, …)` consumes the editor and takes exactly
  one commit.
* `row_visibility::SourceEditor::edit` has the same shape, and a row-visibility
  commit never consults the facts at all.
* The only way to chain two cumulative value-only commits is publish-then-reopen,
  which is what the crate's own
  `tests/source_backed_cell_values/compact_source_proof.rs` does — and a reopen
  re-plans, so change 0622 already covers it.

So the capture is paid on every commit and read on none. Measured on the same
isolation pairs, the retained change against the retained change plus this patch:

CHAIN_TABLE

The capture costs almost exactly what the builder costs — it *is* the builder, run a
second time on the rewritten worksheet, which is why `observe`, `element`, `cell` and
`raw_attribute` all double in that leg — and the whole-iteration cost reaches **+4.06%
on `producer-dense`**, within sight of the 5% review trigger. The benefit it would buy — a
second commit skipping `scan_with_limit`, which change 0622 measured at 17.7 M to
137.5 M Ir per operation — is not reachable from any public entry point, so it
cannot be measured at all.

The rule this falls under is "keep only what is statistically and practically
useful; revert speculative complexity". A saving that no caller can reach is not a
saving, and the cost of reaching for it is measured above. The patch is retained so
that a future change which *does* expose a chained edit — seeding a
`MultiSourceEdit` from a `MultiCommit`, say — can land it with its oracle already
written. Until then this record's answer to "snapshot chains drop facts after the
first commit" is: **they do, and closing that gap costs 0.20% to 4.06% of a save
for a benefit no caller can reach.**

## Why `stream_count` was not substituted for `styles::parse`

`raw::styles::stream_count` already exists (it serves the source-backed workbook
facade) and counts `cellXfs/xf` records without materializing projected XML, so it
looks like the count-only scan XLSX-5 asks for. It is **not refusal-equivalent** to
`parse`, and substituting it would move refusals rather than remove work.

The asymmetry has one cause. `process_ooxml` short-circuits: when the 59-byte markup
compatibility URI does not occur in the input it returns `Cow::Borrowed` after a
single linear scan, with no reader, no event and no allocation. So on an MCE-free
stylesheet — which is every stylesheet in this repository's corpora —
`parse` never runs the MCE validator at all, while `stream_count` always runs the
full validating stream. `parse` therefore **accepts** and `stream_count` **refuses**
each of: trailing text or CDATA after the root element; a `<!DOCTYPE>`; a processing
instruction; a custom general entity; a late XML declaration; an unbound namespace
prefix on an element `parse` classifies as foreign; a nesting depth above the
streaming limit; an event count above 1,000,000; and any event, attribute-count,
attribute-byte, name-byte or context-byte bound the stream enforces and `parse` does
not. The two also differ in their byte bounds (`parse` is pinned to the 256 MiB /
512 MiB `Limits::default()`, `stream_count` takes the caller's), in their allocation
failure shape (`parse` aborts, `stream_count` returns a typed
`Error::Allocation`), and — latently — in MCE branch selection, since `parse`
hard-wires `Capabilities::default()`.

The error identity would move as well. `stream_count` returns
`StreamError<crate::Error, crate::Error>`, and the established mapping in
`workbook/source.rs` discards the observer's own message whenever an MCE error
accompanies it, and relabels every `MceError::Xml` as
`"invalid worksheet extension XML: …"` — a worksheet-specific string for a failure in
`xl/styles.xml`.

That is a contract change: it moves when a refusal happens and what it says. Per the
program's rule it is frozen here as a design and not implemented. The reuse of an
already-built catalog that XLSX-5 offers as the alternative is also unavailable on
this path: `workbook/model.rs` and `workbook/source.rs` each memoize a
`raw::styles::Catalog` in a `OnceLock`, but `capture_auxiliary_source` is reached
from `SourceBackedEditor`, which is constructed from a `SourceBackedPackage` and
never builds a `SourceBackedWorkbook`, so no catalog exists to reuse.

## Validation preserved

Every validation that ran before this change still runs, on the same bytes, in the
same order, with the same messages. The value-only XML validator is untouched. The
fact builder's decline set is unchanged, proved by the two equivalence tests and by
change 0622's oracle, whose corpus funnel line is unchanged. The styles parser's
refusals are unchanged, proved by its own suite and by the reasoning above. The
commit's `scan` is still the authority and still runs whenever the builder declined
or the plan is outside the shape the facts describe.

## Limitations

- **No speedup, throughput, cold-cache, physical-I/O or concurrency claim is made.**
  The paired medians are reported with the window's own A/A floor.
- The instruction counts are one deterministic run per leg. They rank work, not
  latency; callgrind counts `rep movsb` per byte and runs SHA-256 in software.
- **The route is unreachable on real producer files.** Change 0602 established that
  the source-backed value editor admits none of the 95 real `.xlsx` fixtures, and
  change 0622's oracle funnel — reproduced unchanged here at 391 worksheet parts,
  207 admitted to the shared traversal, 1 accepted by the value-only validator and 1
  publishing facts — confirms it. This change neither widens nor narrows that
  surface. Every number above describes the harness corpora, of which the 0601
  producer-*shaped* `edit` variant is the closest approximation to a real file that
  this path can reach.
- **XLSX-5's reduction is real but small on every corpus this repository has.** The
  stylesheet the harness generates is 603 bytes and 26 events, so
  `capture_auxiliary_source` costs 0.17-0.28 M Ir per planning — 0.02% to 0.44% of a
  measured iteration — and the part removes a deterministic 80 allocation calls and
  6,457 allocated bytes from it. That allocation figure is exact and reproducible in
  every case and shape; the instruction figure is inside the run-to-run variation of
  the symbol. A real Excel stylesheet is larger and carries the markup-compatibility
  namespace, which takes the full tokenising MCE path rather than the borrowed
  short-circuit; **that case is not measured here**, because no corpus in this
  repository reaches `capture_auxiliary_source` with one, and the saving is
  proportional to events, so it is modelled, not measured.
- The builder's decline set is unchanged, so worksheets carrying `<f>` are still
  declined. That remains change 0622's deliberate narrowing, not a measured
  rejection.
- The chain part's cost is measured on four scenarios, not on the full matrix, and
  its benefit is not measured at all because no caller reaches it.
- No process peak RSS figure is reported: no retained state changed
  (`change_0622_retained_records_stay_compact` still pins 16 bytes per cell and 32
  per row) and the allocator regions show `region_peak_live_bytes` moving by 3 bytes
  or less in every region of every case.
- `debug_assert!` pins `raw_attribute`'s colon-free precondition in debug builds
  only; in release the argument is the four literal call sites, all of which pass
  `r`, `min`, `max` or `ref`.
- **The measured binaries are not reproducible byte for byte.** The workspace release
  profile sets `lto = true`, and rebuilding the committed source on this toolchain
  produced a different `litchi-perf-baseline` sha256 on each of three attempts. The
  hashes in the packet identify the exact artifacts that were timed and profiled, and
  nothing more; this is a property of the build, not of this change, and it is
  recorded here because every packet in this program quotes those hashes.

## Retained evidence

[`results/change-0635/README.md`](results/change-0635/README.md).
