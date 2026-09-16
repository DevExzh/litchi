# Change 0641 evidence

Items **XLS-5** and **XLS-8** of change
[0587](../../0587-remaining-opportunity-survey.md): a measure-only instantiation
for the cell records a selected-cell query does not keep, and a second
`StreamChainHint` dedicated to the worksheet region. The record is
[`0641-xls-scan-measure-only-cells-and-sheet-hint.md`](../../0641-xls-scan-measure-only-cells-and-sheet-hint.md).

## Contents

| path | what it is |
| --- | --- |
| [`probe/cell_record_census.py`](probe/cell_record_census.py) | the corpus census XLS-5 asked for first: the CFB, BIFF8 and worksheet-substream walk, in pure arithmetic, sharing no code with `litchi-xls`. The container walk is taken unchanged from change 0608's `sst_prefix_census.py`, which took it from change 0584's `sst_walk.py`. |
| [`probe/census.jsonl`](probe/census.jsonl) | one JSON object per fixture: per-sheet and folded record counts by kind, packed-cell counts, `Label` character bytes, `Formula` token bytes, `RgbExtra`-bearing and string-pending `Formula` counts. |
| [`probe/census-summary.json`](probe/census-summary.json) | the corpus fold: 126 fixtures, 372 worksheet substreams, 106,689 records, 127,072 cell values, and the `Label`/`Formula` shares. |
| [`counts/counts.sh`](counts/counts.sh) | the control driver: one warmup and five samples per scenario per leg. |
| [`counts/counters.json`](counts/counters.json) | the control: read calls, read bytes, `version()` calls and the harness observation for 16 scenarios on both legs, with an `identical` flag per row. All 16 are identical. |
| [`counts/scenarios.txt`](counts/scenarios.txt) | the 16 scenarios, as `label\|fixture\|worksheet-index\|operation`. |
| [`chain/instrumentation.md`](chain/instrumentation.md) | the three runtime switches, the three legs, why one binary is sound for this quantity and not for the others, and the removal check. |
| [`chain/chain_probe.rs`](chain/chain_probe.rs) | the probe, formerly `crates/litchi-xls/examples/chain_probe.rs`, removed before commit. |
| [`chain/corpus.tsv`](chain/corpus.tsv) | full text over every `.xls`/`.xlt` fixture, three legs, with open links, scenario links and an FNV-1a digest of the extracted text per row. |
| [`chain/full-text.tsv`](chain/full-text.tsv) | the six fixtures the record tabulates, captured separately. |
| [`chain/query.tsv`](chain/query.tsv) | one-cell and all-cells on four fixtures, `base` against `hint`: unchanged to the link. |
| [`chain/pre-instrumentation.sha256`](chain/pre-instrumentation.sha256) | the digests of the four instrumented files, verified after removal. |
| [`callgrind/callgrind.sh`](callgrind/callgrind.sh) | the isolation-pair driver: 2 and 12 harness samples per leg per scenario, pinned to CPU 24. |
| [`callgrind/cg-scenarios.txt`](callgrind/cg-scenarios.txt) | the eight profiled scenarios. |
| [`callgrind/fold_callgrind.py`](callgrind/fold_callgrind.py) | differences the two profiles per symbol and divides by 10. |
| [`callgrind/totals.txt`](callgrind/totals.txt) | whole-operation instructions per scenario, both legs. |
| [`callgrind/symbols.txt`](callgrind/symbols.txt) | the per-symbol self-cost movers for four scenarios, including the two that do not improve. |
| [`cycles/cycles.sh`](cycles/cycles.sh) | `perf stat` isolation pairs, 20 and 120 samples, `-r 5`, five repetitions per leg, three legs (`before`, `after`, `before2`) so the A/A floor is measured in the same window. |
| [`cycles/perfsum.py`](cycles/perfsum.py) | the fold: median of the five repetitions per leg. |
| [`cycles/summary.txt`](cycles/summary.txt) | cycles and instructions per operation for 11 scenarios, with the A/A floor beside each row. |
| [`callgrind/inline-trial.md`](callgrind/inline-trial.md) | the two codegen measurements that placed the three `#[inline]` attributes: the first pass without them, and the `#[inline(always)]` alternative that was measured and rejected. |
| [`cycles/inline-trial/`](cycles/inline-trial/) | the raw `perf stat` output of the `#[inline(always)]` trial, three legs, three repetitions. |
| [`cycles/flagship-repeat/`](cycles/flagship-repeat/) | five repetitions of the flagship one-cell pair on three legs, retained as the reason the cycle table is a median of five rather than one pair. |
| [`timing/time.sh`](timing/time.sh) | the single-window paired driver, order A1 B1 B2 A2 A3 on CPU 24. |
| [`timing/time_rounds.sh`](timing/time_rounds.sh), [`timing/time_rounds_window2.sh`](timing/time_rounds_window2.sh) | the same order repeated as interleaved rounds and pooled per leg, so a drift inside the window is charged to every leg equally. Window 2 uses it for all eleven scenarios. |
| [`timing/stats.py`](timing/stats.py) | p50/mean/p95/p99 per leg, the two paired directions, and the A/A and B/B floors. |
| [`timing/time-scenarios.txt`](timing/time-scenarios.txt) | the 11 timed scenarios. |
| [`timing/summary-window2.txt`](timing/summary-window2.txt), [`.json`](timing/summary-window2.json) | **the table the record quotes**: the quiet window, load average 6.2, 60 samples per leg. |
| [`timing/summary-window1.txt`](timing/summary-window1.txt), [`.json`](timing/summary-window1.json) | the earlier window at load 11-30, retained in full as the confirming window. |
| [`timing/window2-raw/`](timing/window2-raw/), [`timing/window1-raw/`](timing/window1-raw/) | every retained sample, in nanoseconds, one file per scenario per leg per window. |
| [`mutations.txt`](mutations.txt) | the eight record-level mutations, M1-M8, and what caught each. |
| [`mutate_records.py`](mutate_records.py) | applies each record-level mutation to the merged tree, runs `cell_measure_tests`, reverts. |
| [`mutations-scan.txt`](mutations-scan.txt) | the four scan-level mutations, N1-N4, and what caught each. |
| [`mutate_scan.py`](mutate_scan.py) | the same driver for N1-N4, against the whole `litchi-xls` suite. |
| [`decision.json`](decision.json) | the decision record: reason codes, accepted evidence, accepted costs, known gaps, withheld results, provenance. |
| [`gates.txt`](gates.txt) | the tail of every gate run, including the two pre-existing harness failures and their reproduction on the untouched base. |
| [`log-sections.md`](log-sections.md) | the four log paragraphs for the coordinator to merge into `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`. |

## Provenance

| | |
| --- | --- |
| base commit | `c7326f68065edf6f2198ca3cb39c38c48cf00ed9` |
| branch | `perf/0641-xls-scan-measure-only-cells-and-sheet-hint` |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0, cargo 1.95.0, valgrind 3.26.0 |
| CPU pin | `taskset -c 24` |
| host load | load average 11 to 30 on 32 cores; **quiescence is not established**, which is why every timing row carries an A/A floor |

Both legs are `cargo build --release --locked --features xls-source-attribution
--bin xls_source_attribution` from `tools/perf-baseline`. The before leg is the
shared read-only checkout at
`/home/zhuhe/code/litchi-worktrees/before-c7326f680` with its own
`CARGO_TARGET_DIR`. Both binaries were copied out of their Cargo target
directories before being timed, because change 0627 had a concurrent build
relink one mid-run.

| binary | sha256 |
| --- | --- |
| `before/xls_source_attribution` | `2aa5a2a47f18f82d8e3bb007074cd76cb06fd1b11dc07dc4686823b657f0917a` |
| `after/xls_source_attribution` | `4d48a2dbf86a0fd26900ccc18d0088663d64fb3c93b4a3318efb1b476550df14` |

The chain-link legs are **not** from these binaries: they come from one
instrumented build carrying three runtime switches, for the reasons set out in
[`chain/instrumentation.md`](chain/instrumentation.md).

## Fixtures

| label | path | why |
| --- | --- | --- |
| `54016` | `test-data/poi/test-data/spreadsheet/54016.xls` | the survey's reference query: 37,929 records on one sheet, no `Formula` and no `Label` — so it prices the enum itself |
| flagship | `test-data/ole/xls/ConditionalFormattingSamples.xls` | 16 worksheets; the program's standing multi-sheet fixture. Its text extraction is refused by this reader (`Invalid record 0x0006: shared Formula metadata requires a leading PtgExp token`), which is an outcome held identical across legs, not a reason to drop the scenario |
| `WithCustomViews` | `test-data/ole/xls/WithCustomViews.xls` | the survey's third reference fixture |
| `15228` | `test-data/poi/test-data/spreadsheet/15228.xls` | **added by this change**, on the census: 18 worksheets and 11,240 `Formula` records carrying 175,054 `Rgce` bytes, 73.8% of its worksheet records. The survey named no formula-heavy fixture because the mix was unknown |
| `HyperlinksOnManySheets` | `test-data/ole/xls/HyperlinksOnManySheets.xls` | named by XLS-8; three worksheets and 113 records, small enough that its timing sits inside the floor |
| `59858` | `test-data/poi/test-data/spreadsheet/59858.xls` | seven worksheets, the second-largest chain-link total in the corpus |

## Reproducing

```sh
# the census (no build needed)
python3 -B docs/performance/results/change-0641/probe/cell_record_census.py test-data

# the differential and the mutation sweeps
cargo test -p litchi-xls --all-features --lib cell_measure_tests -- --nocapture
cargo test -p litchi-xls --all-features --test source_backed

# the measurement drivers, after staging both binaries as bin/xsa-{before,after}
docs/performance/results/change-0641/callgrind/callgrind.sh cg-scenarios.txt
docs/performance/results/change-0641/cycles/cycles.sh time-scenarios.txt
docs/performance/results/change-0641/timing/time.sh time-scenarios.txt
```

The chain-link legs need the instrumentation in
[`chain/instrumentation.md`](chain/instrumentation.md) reapplied first.
