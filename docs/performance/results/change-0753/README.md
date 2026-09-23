# Change 0753 evidence packet

Record: [0753-legacy-fresh-writer-text-paths](../../0753-legacy-fresh-writer-text-paths.md).
Base `6d989cad63`; measured production code `4c079ef4bf` (branch
`perf/0753-legacy-fresh-writer-text-paths`). `performance_claim: none`.

## Contents

| path | what it is |
| --- | --- |
| `environment.txt` | host, kernel, toolchains, pinning, measurement windows |
| `binaries.sha256` | SHA-256 of the six measured binaries (harness, allocator-metrics harness and probe, per leg) |
| `gates.txt` | every gate command, the tail of its output and its exit code: a full run at `630a24b17b`, the base-side golden and clippy runs, and the final full run at `4c079ef4bf` |
| `latency/window-1/`, `latency/window-2/` | ABBA runs of the harness selectors (`scripts/abba_harness.py`): `summary-*.json` (per-process p50/p95/mean, paired ratios, bootstrap CI, output SHA-256 per leg) and every raw harness report in `raw/` (`<selector>-<shape>-<index>-<leg>.json`) |
| `counters/` | differenced `perf stat` counters (`scripts/perf_counters.py`): `summary-*.json` and the raw `perf stat -x,` outputs in `raw/` |
| `callgrind/` | exact instructions per write from the probe under callgrind (`scripts/callgrind_probe.py`): `summary.json` and each callgrind log |
| `alloc/` | counting-allocator results of the probe (`probe-A.json`, `probe-B.json`, `table.txt`) |
| `profiles/payload-heavy-attribution.json` | timed-region attribution of frame-pointer profiles, before and after, payload-heavy (`scripts/profile_summary.py` over stacks folded by profile r2's `fold.py`; the folded stacks and perf data were deleted) |
| `probe/` | source of the allocation/callgrind probe (the harness's three `write_fresh_*` bodies with a counting global allocator) and its manifest template |
| `scripts/` | the build, measurement, profiling, gate and table scripts used |
| `cleanup.json` | what was removed after the evidence was copied here |

## Reproducing

1. Build each leg with `scripts/build_leg.sh SRC TARGET LOG` (identical command:
   `cargo build --release --locked --offline --manifest-path
   tools/perf-baseline/Cargo.toml`, then the same with `--features
   allocator-metrics --bin litchi-perf-baseline-alloc`), before from a detached
   worktree at `6d989cad63`, after from the branch. Copy the binaries to
   equal-length paths (`bin/A/lpb`, `bin/B/lpb`).
2. `python3 scripts/abba_harness.py OUT 8 [selector/shape,...] LABEL` for timings,
   `python3 scripts/perf_counters.py OUT 8 [selector/shape,...]` for counters.
3. Build `probe/` against each leg's sources (`Cargo.toml` paths point at the
   leg's `crates/`) with `cargo build --release --offline`, then
   `python3 scripts/callgrind_probe.py OUT` and `probe 5` per leg for
   allocations.
4. `scripts/build_prof.sh` builds a frame-pointer, line-table variant used only
   for attribution; `scripts/profile_case.sh` records and folds one case.
5. `python3 scripts/tables.py .` renders the record's tables from this packet.

The golden digests in `crates/litchi-{doc,ppt,xls}/tests/*_writer_text_goldens.rs`
were captured by running those test files against the base first; the final
files pass unchanged on the base worktree (`gates.txt`).
