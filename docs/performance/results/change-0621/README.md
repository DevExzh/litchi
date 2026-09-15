# Change 0621 evidence packet

Record: [`docs/performance/0621-xls-open-fence-count.md`](../../0621-xls-open-fence-count.md).

Source-version observation placement on the source-backed XLS open, one-cell and
full-text paths, and on the facade's workbook detection route. Items **CORE-2**
and the fence half of **XLS-4** from change 0587.

## Provenance

| | |
| --- | --- |
| base commit | `1e41983213dc378c13774ed7038c51faf231977f` (branch `feat/office-format-completeness`) |
| branch | `perf/0621-xls-open-fence-count` |
| before checkout | `/home/zhuhe/code/litchi-worktrees/before-1e4198321` (shared, read-only, untouched) |
| after worktree | `/home/zhuhe/code/litchi-worktrees/0621` |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0 (59807616e 2026-04-14), valgrind 3.26.0 |
| CPU pin | 16 (`taskset -c 16` on every measured process) |
| build | `cargo build --release --locked --features xls-source-attribution --bin xls_source_attribution`, separate `CARGO_TARGET_DIR` per leg |
| quiescence | not established; seven other agents were active on other cores |

Binary digests:

| binary | sha256 |
| --- | --- |
| `xls_source_attribution` (before) | `19796a40a085037ed1d3040812976960a735fc2c423fc50304edeb6b8f84b47c` |
| `xls_source_attribution` (after) | `ed87751de0343e830d657b091969a676476de284514cf26492438047d3a182ec` |
| `xls-observation-sites` probe (before) | `c2fc6cdc0a22db3bea6017b80fbe13ba086bb34bd37e82eeaf58a74361ab3f3d` |
| `xls-observation-sites` probe (after) | `c06e6b9b7102049c178bfdb20078e21ba9b40ae3c56d93849b160810f0b80478` |

## Contents

| path | what it is |
| --- | --- |
| [`gates.txt`](gates.txt) | every gate with its tail, the one pre-existing `litchi` test failure and the pre-existing warning set both reproduced on the untouched before checkout, and the negative control that shows the three counted-invariant tests failing with the fences put back |
| [`decision.json`](decision.json) | the decision record in the `litchi-perf-change-decision` schema |
| [`log-sections.md`](log-sections.md) | the four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, for the coordinator to merge |
| `sites/54016-{open,one-cell,all-cells,full-text}-{before,after}.txt` | **the per-site analysis evidence**: every `ReadAt::version()` call on that operation, bucketed by the `litchi` frames of its backtrace with file and line, and counted |
| `probe/main.rs`, `probe/Cargo.toml` | the probe that produced those files: a `ReadAt` that wraps `FileSource`, captures a backtrace per observation and buckets it. A scratch Cargo project with path dependencies; no harness selector attributes observations per site |
| `counts/observation-counts-{before,after}.txt` | the summary table of counted observations, reads and read bytes for ten fixture/mode/operation cells |
| `counts/*.json` | the `xls_source_attribution` report behind each of those cells, including the projection digest each cell is compared on |
| `strace/summary-{before,after}.txt` | per-operation `statx` and `pread64`, from the one- and eleven-sample runs differenced |
| `strace/*-s{1,11}.txt` | the raw `strace -f -c` tables those summaries difference |
| `callgrind/summary-{before,after}.txt` | per-sample instruction counts from the one- and six-sample isolation pairs |
| `callgrind/*-s{1,6}.txt` | the valgrind output those summaries read (the `callgrind.out` files are not retained: they are large and carry nothing the logs do not) |
| `timing/summary.txt` | p50, mean, p95 and p99 per leg with both paired directions and the A/A floor |
| `timing/*.json` | the 24 timing children, six per case (A1 B1 B2 A2 AA1 AA2), each with its full `elapsed_samples_ns` |
| `corpus/corpus-{before,after}.txt` | the full-text projection digest, the whole-sheet-walk digest, reads, read bytes and observations for all 109 `.xls` fixtures on each leg |
| `corpus/corpus-summary.txt` | the 218-cell comparison |
| `scripts/` | every script above: `counts.sh`, `strace.sh`, `callgrind.sh`, `timing.sh`, `corpus.sh`, `analyze_timing.py`, `analyze_corpus.py`, and `restore-fences.py`, which puts the removed observations back for the negative control |

## Replay

```sh
S=docs/performance/results/change-0621
python3 -B $S/scripts/analyze_corpus.py $S/corpus
python3 -B $S/scripts/analyze_timing.py $S/timing
```

The count, `strace`, callgrind and timing scripts take a built
`xls_source_attribution` and the checkout it was built from as arguments; the
site-attribution probe needs its `Cargo.toml` path dependencies repointed at the
checkout under measurement.

## What the packet does not contain

No cold-cache, physical-device, remote or range-source, peak-RSS, allocation,
concurrency-scaling or cross-platform measurement. No per-site decomposition of
the `facade-file` `statx` counts: the release binary carries no symbols for
`strace -k` to resolve, so those are whole-child figures.
