# 0560: one source observation per XLS freshness check

Status: retained. `performance_claim: none` — this record carries paired and
deterministic measurements, not a registry claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## What was removed

`litchi_xls::workbook::source::ensure_current_parts` took **three** source
observations with no `read_at` between them:

```rust
let observed = source.version()?;           // observation 1
// ... compare against expected_version
cfb.source_version()?;                      // observation 2, result discarded
let observed = source.version()?;           // observation 3
```

`SourceInner.source` is defined as `SharedOleFile::source_arc()` and
`SourceInner.expected_version` as that view's `source_version()`, so the two
expectations are the same value observed against the same object. One
observation discharges both. `SharedOleFile` gains a `const`
`captured_source_version()` accessor so the second expectation can be compared
without taking another observation; its rustdoc states plainly that it observes
nothing and proves nothing on its own.

Both retained-metadata helpers — `SourceBackedWorkbook::metadata` and
`SourceBackedWorksheet::metadata` — called `ensure_current` **twice** around a
closure that reads only in-memory state. They now fence once, after the value is
produced, because it is the trailing observation that bounds what the caller
receives. The worksheet helper's missing-worksheet branch fences before
reporting, so a changed source still takes precedence over `WorksheetNotFound` —
the order the removed leading fence produced.

Combined, a retained metadata query falls from **six** observations to **one**.

## Why it is exactly equivalent

Nothing between the collapsed observations consumes a source byte: observation 2
is a pure fence whose result is discarded, and the metadata closures read
`Box<[_]>` and `Arc<_>` fields the snapshot already owns. Two observations
separated by no consumed bytes prove the same thing one does.

The error identity is unchanged. `SharedOleFile::source_version` reports
`OleError::SourceChanged { expected, observed }`, which
`From<OleError> for SourceBackedError` maps to
`SourceBackedError::SourceChanged { expected, observed }` — the same variant and
payload the retained comparison produces.

A `debug_assert_eq!` records the identity the collapse relies on, and the
comparison still checks `cfb.captured_source_version()` explicitly rather than
assuming it.

## Measured effect

Environment and identities are in [`plan.json`](results/change-0560/plan.json),
frozen before capture. CPU 17 pinned; ASLR disabled for the selector matrix.
Host-wide quiescence is not established.

### Deterministic counters

Corpus `test-data/ole/xls/ConditionalFormattingSamples.xls`:

| Operation | `version()` calls | whole-child `statx` | `read_at` calls | read bytes |
| --- | --- | --- | --- | --- |
| open | 644 → 631 (−2.02%) | 3,898 → 3,820 | unchanged | unchanged |
| list | 644 → 631 (−2.02%) | 3,898 → 3,820 | unchanged | unchanged |
| one-cell | 925 → 902 (−2.49%) | 5,584 → 5,446 | unchanged | unchanged |

`pread64` counts are identical in every cell. The proportion is small here
because this scenario is dominated by the per-read CFB fences that change 0558
already halved; the collapse matters where metadata queries and cell reads are
frequent.

### Paired latency

A1/E1/E2/A2 over ten XLS selectors, 5 warmups and 60 samples per child: 14
case/corpus rows, 56 statistic comparisons, **34 improving in both directions**.

| Selector | p50, both directions |
| --- | ---: |
| `xls_source_backed_open` | −5.48% / −5.84% |
| `xls_source_backed_open_list_worksheets` | −14.78% / −2.85% |
| `xls_source_backed_open_one_cell` | −6.01% / −7.11% |
| `xls_owned_source_open` | −10.88% / −8.18% |
| `xls_owned_source_open_list_worksheets` | −5.28% / −1.91% |
| `xls_owned_source_open_one_cell` | −0.09% / −4.45% |
| `xls_semantic_one_cell` large | −22.22% / −12.50% |
| `xls_semantic_list_worksheets` large | −14.29% / −14.29% |
| `xls_semantic_one_edit_save` large | −2.26% / −3.26% |

### Review triggers

Four comparisons are adverse in both directions by more than 5%. All four are on
the `xls-tiny` corpus at nanosecond scale, where the harness clock's resolution
dominates:

| Case | Statistic | Change | Absolute |
| --- | --- | ---: | --- |
| `xls_semantic_full_cell_scan` xls-tiny | p50 | +5.26% / +5.26% | 190 ns → 200 ns |
| `xls_semantic_full_cell_scan` xls-tiny | mean | +5.54% / +6.55% | 186 ns → 197 ns |
| `xls_semantic_full_cell_scan` xls-tiny | p95 | +5.26% / +10.24% | 190 ns → 200 ns |
| `xls_semantic_list_worksheets` xls-tiny | p99 | +109.33% / +115.33% | 30 ns → 63 ns |

A 190 ns median moves by one 10 ns clock tick for 5.26%, so these cells cannot
distinguish a real regression from quantization. They are reported rather than
excluded; the `xls-large` twins of both selectors are neutral or improve.

### Confirmation against the final source

Independent review of the first candidate found that moving the worksheet
helper's fence after its lookup would report `WorksheetNotFound` where a changed
source previously reported `SourceChanged`. The fence was restored on that
branch before the final build. A binary built from the final source reproduces
the result against the pre-0560 baseline: `version()` calls 644 → 631 and
925 → 902, `read_at` calls and read bytes byte-identical, and file-source p50
−1.74%, −1.66% and −1.88% for open, list and one-cell. Those children are
retained in [`latency/`](results/change-0560/latency) with a `final-` prefix.

## Correctness evidence

Two tests are added to `crates/litchi-xls/tests/source_backed.rs`:

- `retained_metadata_queries_observe_the_source_once` counts observations on an
  instrumented `ReadAt` and asserts **exactly one** observation and **zero**
  source bytes for `worksheet_count`, `worksheet_names`, acquiring a worksheet
  handle, `name` and `visibility`. It converts the collapse into an enforced
  invariant, so re-adding a fence fails the suite.
- `retained_metadata_queries_still_refuse_a_changed_source` bumps the source
  revision and asserts both workbook and worksheet metadata still return
  `SourceChanged`.

The `litchi-xls` suite passes 1,028 library tests plus the integration suites
with zero failures.

## Limitations

No cold-cache, physical-device, remote/range-source, peak-RSS, allocation,
concurrency-scaling, real-producer or cross-platform result is claimed. The
`xls-tiny` corpora sit below this harness clock's useful resolution and should
not be read as evidence in either direction.

Replay:

```sh
python3 -B -c "import json;print(json.load(open('docs/performance/results/change-0560/analysis.json'))['summary'])"
```
