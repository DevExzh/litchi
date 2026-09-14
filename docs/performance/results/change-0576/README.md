# Evidence: change 0576, index the SST without decoding it

Change record: [`0576-xls-sst-scan-without-materialization.md`](../../0576-xls-sst-scan-without-materialization.md).

Disposition: retained. `performance_claim: none`.

## Contents

| Path | What it is |
| --- | --- |
| `environment.json` | Host, toolchain, quiescence log, fixture hashes, and the SHA-256 of both leg binaries and of the one file that differs between them. |
| `analysis.txt` | The exact output of `analyze.py` over this directory. Every number the record cites appears here. |
| `analyze.py` | Recomputes all three measurement tables from this directory alone. Also folds a raw capture (`--fold`). Pure standard library. |
| `latency-summary.json` | The folded A/B/B/A wall-clock capture: 64 cells, each with p50/p90/p99/mean, sample count, the three logical counters, and the binary SHA-256 the cell was produced by. |
| `counters/perf-<leg>-<fixture>-s{100,1100}.csv` | `perf stat -x,` isolation pairs. Differencing the 1,100-sample and 100-sample child and dividing by 1,000 isolates one open, exactly as change 0574 did. |
| `callgrind/ann-<leg>-<fixture>-s{small,large}.txt` | `callgrind_annotate` self cost, same isolation method. |
| `callgrind/incl-<leg>-<fixture>-s{small,large}.txt` | The same profiles, `--inclusive=yes`, which is where the scan's inclusive cost comes from. |
| `differential.txt` | The corpus differential: both instantiations of the scan over every `.xls`/`.xlt` fixture under `test-data`, with the per-fixture skip and refusal reasons. |
| `allocations.txt` | Allocations and allocated bytes per source-backed open, before and after, from the same test file run against both legs. |
| `mutations.txt` | Six mutations of the new code, and which tests catch each. |
| `mutate.py` | Applies one mutation by name. |
| `sst_composition.py`, `sst-composition.txt` | Shared-string composition of the three fixtures, read straight from their SSTs with `olefile` plus the standard library. No repository code runs. It is what the record's empty-string, rich-text, `ExtRst` and `Continue`-crossing counts come from. |
| `capture_latency.sh` | The A/B/B/A wall-clock driver. Delegates each leg to change 0574's `capture_counters.sh`, unchanged. |
| `capture_perf.sh` | The hardware-counter driver. |
| `capture_callgrind.sh` | The callgrind driver. |

## Result

One source-backed open of `WithCustomViews.xls` falls from **208,394 ns to
31,950 ns** (−84.6%) and from **5,312,769 to 627,989 instructions** (−88.2%); one
open of `54016.xls` from **688,168 ns to 256,704 ns** (−62.6%) and from
**17,139,968 to 5,979,090 instructions** (−65.1%); one open of the flagship
`ConditionalFormattingSamples.xls` from **90,520 ns to 80,138 ns** (−11.4%) and
from **1,643,022 to 1,391,409 instructions** (−15.3%). The measured same-binary
noise floor in that window is ±1.4% at p50.

Allocations for one open of `54016.xls` fall from **16,031 to 247** — exactly two
per shared string, removed. Every logical read counter is unchanged.

The scan's index is byte-identical: **17,434 shared-string entries across 117
fixtures**, plus 4 fixtures refused identically by both paths.

## Replay

```sh
python3 -B docs/performance/results/change-0576/analyze.py \
  docs/performance/results/change-0576
```

## Rebuilding the captures

Both legs are the standalone `tools/perf-baseline` package, built inside a
detached git worktree placed **outside** the repository so that neither leg can
pick up unrelated working-tree changes and neither build touches a repository
target directory:

```sh
SCRATCH=/some/scratch/dir
git -C . worktree add --detach $SCRATCH/leg HEAD
cd $SCRATCH/leg/tools/perf-baseline
CARGO_TARGET_DIR=$SCRATCH/legtarget cargo build --release --locked \
  --features xls-source-attribution --bin xls_source_attribution
cp $SCRATCH/legtarget/release/xls_source_attribution $SCRATCH/xls_source_attribution.before

cp <repo>/crates/litchi-xls/src/records.rs $SCRATCH/leg/crates/litchi-xls/src/records.rs
CARGO_TARGET_DIR=$SCRATCH/legtarget cargo build --release --locked \
  --features xls-source-attribution --bin xls_source_attribution
cp $SCRATCH/legtarget/release/xls_source_attribution $SCRATCH/xls_source_attribution.after
```

Then, from the repository root:

```sh
R=docs/performance/results/change-0576
W=50 S=1000 $R/capture_latency.sh $SCRATCH/xls_source_attribution.{before,after} $SCRATCH/latency
$R/capture_perf.sh      $SCRATCH/xls_source_attribution.before before $SCRATCH/counters
$R/capture_perf.sh      $SCRATCH/xls_source_attribution.after  after  $SCRATCH/counters
$R/capture_callgrind.sh $SCRATCH/xls_source_attribution.before before $SCRATCH/callgrind
$R/capture_callgrind.sh $SCRATCH/xls_source_attribution.after  after  $SCRATCH/callgrind
python3 -B $R/analyze.py --fold $SCRATCH/latency $R/latency-summary.json $R/environment.json
```

The differential, allocation and mutation logs come from the test suite:

```sh
cargo test -p litchi-xls --lib every_sst_fixture_indexes_identically_both_ways -- --nocapture
cargo test -p litchi-xls --test sst_scan_allocations -- --nocapture
python3 -B $R/mutate.py M1   # then cargo test -p litchi-xls --lib sst_measure_tests, then restore
```

The allocation log's *before* leg is the same test file copied into the
unmodified worktree; it fails there, which is the point.

## What is not here

The raw A/B/B/A captures are 41 MB of per-sample arrays and were discarded after
folding; `latency-summary.json` keeps the statistics, the logical counters and
the binary identity, which is everything the record cites. The raw callgrind
`.out` files are discarded after annotation.

No cold-cache, physical-device, remote or range-source, peak-RSS,
concurrency-scaling, real-producer or cross-platform result. No corpus-wide
latency sweep — the corpus-wide evidence here is the differential, which is a
correctness result, not a timing. Host quiescence is **not** established; the
load average is recorded in `environment.json` and the same-binary A/A and B/B
columns are reported for every cell as the defence against it.

Callgrind instruction counts carry change 0574's ERMS caveat: `memset` and
`memcpy` are instrumented per iteration and their shares are upper bounds. This
change touches neither, and no symbol cited above is a string loop.
