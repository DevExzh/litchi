# 0622: sixteen bytes per cell carried from planning delete the XLSX commit's second whole-sheet scan

Status: retained, implemented in `litchi-xlsx`. `performance_claim: none` — the
counts and paired medians below are reported as evidence, not registered as a
claim. The change is **value-identical**: every scenario's published package
keeps its exact `output_sha256`, and a differential oracle compares the
fact-derived rewrite against the scan-derived rewrite, byte for byte and error
for error, for every cell of every worksheet it can reach.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements item **XLSX-1** of
[`0587-remaining-opportunity-survey.md`](0587-remaining-opportunity-survey.md)
(rank 25) as change [0551](changes/0551-xlsx-layout-handoff-feasibility.md) froze
it, after change [0552](0552-xlsx-compact-source-proof.md) rejected a
planning-time proof on resource gates and change
[0553](0553-xlsx-commit-local-compact-proof.md) rejected a commit-local variant
on latency and RSS gates.

## What was changed

`crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/facts.rs` is new. It
carries `SourceFacts`: the `<sheetData>` envelope, one 32-byte record per row,
one **16-byte record per cell** (`litchi_sheet::Cell` plus the `u32` start and
end of that cell's `<c>` element), and the declared `<dimension>` span. No owned
tag, no attribute, no payload span and no per-cell heap allocation is retained.
The scanner slot it replaces is 96 bytes plus a heap `Tag` and a heap `Box<[Span]>`
per cell; `change_0622_retained_records_stay_compact` pins all three sizes.

* **The builder** observes the events of the planning traversal that change
  [0546](changes/0546-xlsx-shared-traversal-retained.md) already runs, so it
  reads no extra byte and opens no second reader. `parse_source_with_observer`
  now hands each observer the event's source span, taken from
  `NsReader::buffer_position()` exactly as the commit scanner takes it.
* **The builder declines** — silently, dropping every fact — on anything whose
  scan outcome it cannot prove: a `<f>` element, a cell or row without an
  explicit `r`, a non-increasing row or cell, an ampersand anywhere inside a
  start tag, a duplicate or misordered `<dimension>`/`<sheetFormatPr>`/`<cols>`/
  `<sheetData>`, an empty `<cols>`, an invalid `<col>` range, any direct
  worksheet child outside `dimension | sheetViews | sheetFormatPr | cols |
  sheetData`, any cell child outside `<v>` and `<is>`, a second root, or a depth
  or event count at the scanner's own ceilings. A decline is never an error and
  never changes an event the traversal delivers.
* **The commit** (`raw/worksheet/edit/package.rs`) takes `Option<&SourceFacts>`.
  It uses them only when the facts describe exactly this source allocation and
  every staged action *updates a cell the facts already recorded* — no
  insertion, no removal, no shared-formula payload, no empty target row.
  Otherwise it runs today's `scan(content)` and today's writer unchanged.
* **Changed cells materialize on demand.** `materialize_cell` re-reads the at
  most 256 `<c>` elements one commit may touch, from their retained spans alone,
  and rebuilds the scanner's own `CellSlot` — `tag_end`, `close_start`, the
  `Option<Tag>` from the unchanged `cell_tag`, the payload spans, the empty
  form. `write_cell` is reused unmodified. `write_sheet_data_from_facts` is the
  fact-side counterpart of `write_sheet_data_with_provenance`; unchanged cells
  and rows are copied from their spans exactly as before, and the 0525 readback
  provenance is recorded from the same addresses.
* **`Snapshot` carries `Option<Arc<SourceFacts>>`**, set only by
  `from_source_selected` and cleared by every snapshot that rebinds its
  worksheet bytes (`from_rewritten_source`, `from_rewritten_value_source`,
  `from_visibility_rewrite`), so a derived snapshot commits through the complete
  scan exactly as before.

No public API, no error type, no limit, no output byte and no `unsafe` changed.

## Why it is sound

**The scan is not skipped, it is proved.** The commit's `scan(content)` both
builds a layout and refuses malformed worksheets. Removing it would erase those
refusals — so the builder publishes facts only after it has mirrored every
structural refusal `scan_with_limit` and `finish_layout` can raise for the
vocabulary that reaches it, and declines otherwise. Both walks read the same
bytes with the same `NsReader` and `check_end_names = true`, so they see the
same event sequence and the same spans.

**The validator's allow-list does most of the proving.** The value-only planning
validator (`cell_values/validation.rs`) admits only
`worksheet, dimension, sheetViews, sheetView, pane, selection, sheetFormatPr,
cols, col, sheetData, row, c, f, v, is, t`, in fixed parentage, with unprefixed
attributes drawn from a closed per-element list. A worksheet that reaches a
value-only commit therefore *cannot* carry `<sheetProtection>`, `<mergeCells>`,
`<dataValidation>`, an `x14` extension or any foreign or markup-compatibility
element. `Layout::protected`, `validations`, `extended_validation`, `merged`,
`merge_cells`, `merge_compatibility`, `defaults_compatibility` and every
`CellSlot::mce_payload` are consequently false or empty on this path, and the
builder's own `<f>` decline empties `formula_ranges`, `shared_formulas` and
`has_shared_formulas`. The guards `validate_actions` consults are proved absent
rather than reconstructed.

**Error order is preserved.** Facts are published only after the validator
reaches EOF, `complete_source_parse` returns `Ok`, and the builder's own final
ordering checks pass, so no provisional state survives a refusal and no builder
diagnostic can become a planning error. At commit the fact route runs
`validate_for_write` on the same payloads in the same address order that
`validate_actions` does, and every decline happens **before the first output
byte is written**. Change 0541's first-error order and change 0525's independent
changed-cell readback are untouched; the 256-action cap (`cell_values/source.rs`
`MAX_BATCH_EDITS`) still bounds how many cells materialize.

**The facts cannot be applied to the wrong bytes.** They are stored on the
snapshot that owns the payload they were built from, and every snapshot that
rebinds its worksheet bytes drops them. As a second, independent guard the facts
carry the address and length of the slice they describe, and `describes` refuses
any other slice — the same pointer-identity idiom `from_rewritten_value_source`
already uses for its readback provenance.

**Attribute decoding cannot silently move.** The scanner decodes and normalizes
every attribute of every tag it materializes, and that decode can only fail on
an entity reference. The builder therefore declines any start tag containing an
ampersand, so a tag the scanner would have refused is never skipped; and it
restricts the `r`, `min`, `max` and `ref` values it reads to alphabets on which
attribute-value normalization is the identity.

**ADR reading.** ADR 0005 keeps mandatory structural validation where it is and
loads semantic payloads lazily: this change moves no validation, it removes a
*second* traversal whose only unique product was a layout. ADR 0003's bounded
resources are respected — the retained state is a flat `Box<[CellFact]>` with no
per-cell heap object, sized by the same worksheet the planning `Store` already
holds, and every push uses `try_reserve`. ADR 0006's preservation contract is
unaffected: the output bytes are identical, proved by the oracle and by the
harness's own `output_sha256`.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind/callgrind 3.26.0, `taskset -c 15`, release `--release --locked`. Eight
agents were building on the host throughout. Base `1e4198321`; branch
`perf/0622-xlsx-compact-source-facts`.

### Value identity

`output_sha256` from `tools/perf-baseline` is identical on both legs for **32 of
32** (case, shape) pairs: `one_edit_save`, `one_percent_edit_save`,
`batch_edit_save`, `multi_sheet_edit_save`, the two managed variants,
`cell_clear_edit_save` and `cell_remove_edit_save`, over `medium`,
`dense-sparse`, `noncompact` and `vendor-extension`
(`results/change-0622/differential/output-hashes.txt`).

### Deterministic counts

Callgrind isolation pairs, N = 1 and N = 11 harness samples, totals differenced
and divided by 10, one deterministic run per leg
(`results/change-0622/counts/instruction-summary.txt`).

**Per-symbol inclusive Ir per operation.** The differenced per-operation cost of
`scan::scan_with_limit` collapses from 17.7-137.5 M Ir to at most 4,628 Ir —
0.003% of its former size and within the noise of differencing two profiles, so
the walk no longer runs per operation. The table writes those residuals as `≈0`:

| case | shape | symbol | before | after | delta |
| --- | --- | --- | ---: | ---: | ---: |
| one edit | medium | planning (`Snapshot::from_source_selected`) | 23,658,195 | 27,726,299 | **+17.20%** |
| one edit | medium | commit (`rewrite_value_only_with_provenance`) | 18,406,801 | 205,359 | **−98.88%** |
| one edit | medium | the removed walk (`scan_with_limit`) | 17,727,405 | ≈0 | −100% |
| one edit | dense-sparse | planning | 164,090,362 | 193,159,174 | +17.72% |
| one edit | dense-sparse | commit | 129,702,821 | 1,318,090 | **−98.98%** |
| one edit | dense-sparse | the removed walk | 125,515,694 | ≈0 | −100% |
| one edit | noncompact | planning | 26,308,248 | 30,251,272 | +14.99% |
| one edit | noncompact | commit | 25,625,689 | 210,662 | **−99.18%** |
| one edit | noncompact | the removed walk | 22,757,377 | ≈0 | −100% |
| one percent | medium | planning | 96,084,074 | 112,413,786 | +17.00% |
| one percent | medium | commit | 75,200,897 | 2,808,548 | **−96.27%** |
| one percent | medium | the removed walk | 71,031,988 | ≈0 | −100% |
| one percent | dense-sparse | planning | 180,852,405 | 212,635,380 | +17.57% |
| one percent | dense-sparse | commit | 146,851,844 | 7,343,534 | **−95.00%** |
| one percent | dense-sparse | the removed walk | 137,530,166 | ≈0 | −100% |

**Planning plus commit**, which is the sum change 0550 required a candidate to
improve rather than shift:

| case | shape | before | after | delta |
| --- | --- | ---: | ---: | ---: |
| one edit | medium | 42,064,996 | 27,931,658 | **−33.60%** |
| one edit | dense-sparse | 293,793,183 | 194,477,264 | **−33.80%** |
| one edit | noncompact | 51,933,938 | 30,461,935 | **−41.34%** |
| one percent | medium | 171,284,970 | 115,222,334 | **−32.73%** |
| one percent | dense-sparse | 327,704,249 | 219,978,915 | **−32.87%** |

**Whole harness iteration**, which also includes the per-sample output
verification the timed interval excludes, so it is the conservative bound:
−1.19% (one edit, medium), −5.29% (one edit, dense-sparse), −1.83% (one edit,
noncompact), −3.84% (one percent, medium), −5.53% (one percent, dense-sparse).

**Tier: measured**, one deterministic run per leg, harness corpora only.

### Allocations

`litchi-perf-baseline-alloc`, the change-0538 phase regions, maximum over five
samples (`results/change-0622/alloc/allocation-summary.txt`):

| shape | case | region | metric | before | after | delta |
| --- | --- | --- | --- | ---: | ---: | ---: |
| medium | one edit | plan | allocation calls | 17,231 | 17,251 | +0.12% |
| medium | one edit | plan | allocated bytes | 3,013,517 | 3,187,557 | +5.78% |
| medium | one edit | commit | allocation calls | 22,480 | 10,211 | **−54.58%** |
| medium | one edit | commit | allocated bytes | 2,598,300 | 758,957 | **−70.79%** |
| dense-sparse | one edit | plan | allocation calls | 116,915 | 116,936 | +0.02% |
| dense-sparse | one edit | plan | allocated bytes | 13,650,801 | 14,183,753 | +3.90% |
| dense-sparse | one edit | commit | allocation calls | 151,522 | 67,653 | **−55.35%** |
| dense-sparse | one edit | commit | allocated bytes | 11,991,930 | 4,825,535 | **−59.76%** |
| noncompact | one edit | plan | allocation calls | 18,384 | 18,404 | +0.11% |
| noncompact | one edit | plan | allocated bytes | 3,026,569 | 3,200,609 | +5.75% |
| noncompact | one edit | commit | allocation calls | 36,309 | 10,220 | **−71.85%** |
| noncompact | one edit | commit | allocated bytes | 3,033,244 | 771,435 | **−74.57%** |
| medium | one percent | commit | allocation calls | 91,391 | 42,772 | **−53.20%** |
| medium | one percent | commit | allocated bytes | 12,436,307 | 5,103,279 | **−58.96%** |
| dense-sparse | one percent | commit | allocation calls | 172,946 | 80,835 | **−53.26%** |
| dense-sparse | one percent | commit | allocated bytes | 18,258,492 | 8,883,660 | **−51.35%** |
| noncompact | one percent | commit | allocation calls | 146,707 | 43,341 | **−70.46%** |
| noncompact | one percent | commit | allocated bytes | 14,176,758 | 5,170,165 | **−63.53%** |

The full matrix, including the staging and publication regions and the
`region_peak_live_bytes` column (−1.56% to +0.98%), is in
`results/change-0622/alloc/allocation-summary.txt`.

Publication and staging are byte-identical in every row. Planning pays the
retained facts: **+0.02% to +0.12% allocation calls** and **+3.90% to +6.09%
allocated bytes**; commit returns **−53.2% to −71.9% of its allocation calls**
and **−51.4% to −75.2% of its allocated bytes**. This is the trade change 0553
failed to make.

### Process peak RSS

The gate change 0552 failed (`+6.97%`) and change 0553 failed (`+5.16%`), on the
same `dense-sparse` shape, measured here with `/usr/bin/time -f %M`, fifteen runs
per leg, `--warmup 2 --samples 10` (`results/change-0622/alloc/rss-repeat.txt`):

| leg | min KiB | p50 KiB | max KiB | own spread |
| --- | ---: | ---: | ---: | ---: |
| before | 86,900 | 87,840 | 92,544 | 6.49% |
| after | 87,656 | 91,556 | 92,596 | 5.64% |

Median-to-median **+4.23%**, minimum-to-minimum **+0.87%**. The before leg's own
run-to-run spread is 6.49%, larger than the delta, so this figure bounds rather
than measures the retained state's cost; the minimum-to-minimum figure is the
one consistent with the ~0.5 MB the allocator region actually attributes to the
facts. Reported because it is the number the two previous attempts died on.

### Paired timing

`tools/perf-baseline`, `--warmup 5 --samples 30`, order A1 B1 B2 A2 on CPU 15,
followed by four before-only runs F1..F4 for the A/A floor in the same window.
Five source-backed cell-value cases over three shapes, 15 records per run. Full
table, with mean, p95 and p99 and the raw p50 nanoseconds:
`results/change-0622/timing/abba-summary.txt`.

Paired p50 deltas, positive = this change faster, beside that scenario's own A/A
floor:

| case | shape | leg 1 (A1→B1) | leg 2 (A2→B2) | A/A floor p50 |
| --- | --- | ---: | ---: | ---: |
| one edit | medium | +17.92% | +17.18% | 6.78% |
| one edit | dense-sparse | +18.22% | +13.08% | 3.33% |
| one edit | noncompact | +31.21% | +24.29% | 8.87% |
| one percent | medium | +18.13% | +15.76% | 6.35% |
| one percent | dense-sparse | +19.00% | +16.92% | 3.33% |
| one percent | noncompact | +27.68% | +26.96% | 5.50% |
| batch (256 cells) | medium | +18.00% | +19.96% | 6.36% |
| batch (256 cells) | dense-sparse | +17.72% | +16.98% | 3.76% |
| batch (256 cells) | noncompact | +26.12% | +27.21% | 5.82% |
| managed one edit | medium | +25.99% | +16.36% | 15.85% |
| managed one edit | dense-sparse | +16.82% | +17.88% | 4.60% |
| managed one edit | noncompact | +28.79% | +26.09% | 8.12% |
| managed one percent | medium | +18.12% | +18.21% | 6.16% |
| managed one percent | dense-sparse | +17.53% | +18.08% | 5.07% |
| managed one percent | noncompact | +27.43% | +27.02% | 6.69% |

**Every scenario improved, in both legs, at p50, mean, p95 and p99.** The
smallest delta anywhere in the matrix is +11.97% (one edit, dense-sparse, leg 2,
p99) and the largest +31.81% (one edit, noncompact, leg 1, p99). There is **no
regression to report** on any scenario or statistic. Every p50 delta exceeds its
own floor; the narrowest margin is managed one-edit on `medium`, +16.36% against
a 15.85% floor in leg 2, and its leg 1 is +25.99%.

Raw p50 for the two extremes: one edit on `noncompact` 6.73 ms → 4.63 ms and one
percent on `dense-sparse` 40.20 ms → 32.56 ms.

**Tier: measured**, four legs plus a four-run floor in one window on a host with
eight agents active. No claim is registered.

## Correctness evidence

**Differential oracle** (`crates/litchi-xlsx/src/cell_values/facts_oracle_tests.rs`,
five tests). For one worksheet it runs the planning traversal, then for each
cell — and for a 256-cell batch, and for three coordinates the sheet does not
contain, updated and removed — calls `rewrite_value_only_with_provenance` twice
over the same bytes and the same actions, once with the facts and once with
`None`, and requires the complete output bytes, the readback provenance spans
and the typed error to agree exactly. Test-only route counters make the
comparison non-vacuous by proving which route ran.

| scope | comparisons | result |
| --- | ---: | --- |
| eleven admitted synthetic shapes (plain, styled/typed, empty cells, inline string, comments and whitespace between cells, prefixed `<x:c>`, `<cols>` and `<sheetFormatPr>`, `<sheetViews>`, an expanding dimension, no dimension, an MCE declaration, an empty `<row/>`) | every cell, plus a batch, plus six insert/remove coordinates | identical bytes, provenance and errors; the fact route ran on each |
| sixteen declined shapes (a formula, a shared formula, inferred cell and row references, an entity in an attribute, an empty `<sheetData/>`, a duplicate and a late `<dimension>`, descending rows and cells, an empty `<cols/>`, merged ranges, sheet protection, a data validation, an unknown cell child, a cell outside its row) | the same grid | identical bytes and errors; the fact route ran on none of them |
| a dense 48x48 scalar grid in the harness's own shape | 96 strided cells, a 256-cell batch and six insert/remove coordinates | identical; the fact route ran on all but the six |
| **every worksheet part of every `.xlsx` under `test-data/`** | 391 parts, 2,790 comparisons | identical bytes, provenance and errors in every comparison |
| the same four tests with `LITCHI_0622_FULL_ORACLE=1`, which removes the stride cap and grows the grid to 128x100 | 2,840 corpus comparisons plus 12,800 grid cells | all five tests pass (`differential/oracle-full.log`) |

The corpus funnel, printed by the sweep: of **391** real worksheet parts, **207**
are admitted to the shared planning traversal, **1** is accepted by the
value-only planning validator, and that same **1** publishes facts
(`results/change-0622/differential/oracle-bounded.log`). A lexical census of the
same 389 parts read through `zipfile` shows why: 388 carry a direct worksheet
child outside the validator's allow-list — `<pageMargins>`, `<pageSetup>`,
`<drawing>`, `<extLst>` and their siblings — 184 carry `x14ac:dyDescent`, 202 a
`<cols>` and 122 an `<f>`
(`results/change-0622/differential/corpus-marker-census.txt`).

**Value identity end to end.** 32 of 32 (case, shape) pairs publish a package
with an identical sha256, across `one_edit_save`, `one_percent_edit_save`,
`batch_edit_save`, `multi_sheet_edit_save`, the two managed variants,
`cell_clear_edit_save` and `cell_remove_edit_save`, over `medium`,
`dense-sparse`, `noncompact` and `vendor-extension`
(`results/change-0622/differential/output-hashes.txt`).

**Suite.** `cargo test -p litchi-xlsx` is 1,305 tests in 59 suites, all passing,
including the five new change-0622 tests and the six change-0541
planning-error-order guards in `source_backed_cell_values`.

**Gates.** `cargo fmt --all --check`, `cargo clippy -p litchi-xlsx --all-targets`,
`cargo test -p litchi-xlsx`, `cargo doc -p litchi-xlsx --no-deps` and
`cargo check -p litchi --all-targets` are clean; tails in
[`results/change-0622/gates.txt`](results/change-0622/gates.txt).
`cargo check --workspace --all-targets` fails in three `litchi-iwa` examples with
`E0308`; that failure is **pre-existing at the base commit** and was reproduced on
the untouched before checkout, and iWork crates are outside this change's scope.

## Validation preserved

Every validation that ran before this change still runs. The value-only XML
validator sees the same bytes and reports the same first error in the same
order; the raw parser's `Store` is built by the same transition function from
the same events; `validation::worksheet_xml` still validates the complete
rewritten output; change 0525's independent reduced readback still parses it and
still merges the omitted cells; both publication audits of change 0528 are
untouched. The commit's `scan` is not weakened — it is still the authority, and
it still runs whenever the builder declined or the plan is outside the shape the
facts describe.

## Limitations

- **No speedup, throughput, cold-cache, physical-I/O or concurrency claim is
  made.** The paired medians are reported with the window's own A/A floor.
- The instruction counts are one deterministic run per leg. They rank work, not
  latency; callgrind counts `rep movsb` per byte and runs SHA-256 in software.
- **The route is unreachable on real producer files, and so is the editor it
  serves.** Of the 391 real worksheet parts under `test-data/`, the value-only
  planning validator accepts exactly one, and that one publishes facts, so the
  builder adds no restriction beyond the editor's own. Change 0602 already
  established that the source-backed value editor admits none of the 95 real
  `.xlsx` fixtures; this
  change does not widen that surface and inherits it. Every measured number
  above therefore describes the harness corpora, which are the only shapes that
  reach this path today.
- The builder declines every worksheet carrying a `<f>` element. That is a
  deliberate narrowing, not a measured rejection: formula ranges, shared-formula
  groups and the scanner's formula-text bookkeeping are not compactly
  reproducible, and the scan keeps them.
- Only the first commit of a snapshot chain uses the facts. A snapshot derived
  from a rewrite drops them, so a second commit on the same chain pays the scan.
- The corpus sweep visits every cell of every worksheet with at most 96 cells and
  a fixed stride through larger ones; `LITCHI_0622_FULL_ORACLE=1` removes the
  cap and was used for the retained evidence run.
- Process peak RSS is a process-lifetime high-water mark, not an operation-local
  peak, and its before-leg spread exceeds the measured delta.

## Retained evidence

[`results/change-0622/README.md`](results/change-0622/README.md).
