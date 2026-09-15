# Evidence: change 0595, a leaner XLS frame loop and a cheaper eager SST walk

Change record:
[`0595-xls-frame-loop-and-sst-walk.md`](../../0595-xls-frame-loop-and-sst-walk.md).

Disposition: retained. `performance_claim: none`.

## Provenance

| | |
| --- | --- |
| Base commit | `08d968f8ec7db27cf1187d01911fd08b9d014d91` (change 0587) |
| Branch | `perf/0595-xls-frame-loop-and-sst-walk` |
| Files that differ between the legs | `crates/litchi-xls/src/records.rs`, `crates/litchi-xls/src/workbook/source.rs` |
| Before binary | `xls_source_attribution`, sha256 `f83468d766af62fbacbecde8d235532e18adc22a823dffdbf0f72c5a02ef8ef0` |
| After binary | `xls_source_attribution`, sha256 `e8ddd41cba8fb82c5f01106558271e04266cfff0db56eeb0cd0e55f21c450aee` |
| Host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| Toolchain | rustc 1.95.0, valgrind 3.26.0, perf 7.0.14 |
| Pinning | CPU 15, `setarch x86_64 -R` (ASLR off), single-threaded |
| Quiescence | **not established**; eight measurement agents shared the host, load 44.41 → 40.45 across the capture window (see `quiescence.log`) |

Full identities, including the three fixture SHA-256s, are in
[`environment.json`](environment.json).

## Contents

| Path | What it is |
| --- | --- |
| `analysis.txt` | The exact output of `analyze.py` over this directory. Every number the record cites appears here. |
| `analyze.py` | Recomputes all six tables from this directory alone. Also folds a raw capture (`--fold`). Pure standard library. |
| `environment.json` | Host, toolchain, leg identities, binary and fixture hashes, harness scope. |
| `counters-summary.json` | The deterministic logical counters: 18 cells (3 fixtures × 2 in-memory modes × open/list/one-cell), before and after, with the harness's implementation projection of each cell. Every sample in a cell carries identical counters, which `analyze.py` asserts while folding. |
| `callgrind/ann-<leg>-<fixture>-<op>-s{small,large}.txt` | `callgrind_annotate --threshold=99.9` self cost. Differencing the large- and small-sample child and dividing by the extra operations isolates one operation, which is change 0574's method as change 0576 and 0584 used it. |
| `perf/perf-<leg>-<fixture>-<op>-s{100,1100}.csv` | `perf stat -x,` cycles, instructions, branches, branch misses and task-clock, isolated the same way. Native counters: `rep movsb` is one instruction and SHA-256 uses the hardware extension, neither of which is true under callgrind. |
| `latency-summary.json` | The folded A1 B1 B2 A2 wall-clock capture: 4 rounds × 18 cells, each with n, p50, mean, p95, p99 and the SHA-256 of the binary that produced it. 500 samples per round after 50 warmups. |
| `sst-index-before.txt` | The SST index over every XLS fixture with an SST, captured on the **untouched base commit** before any production edit: per fixture, segment and entry counts plus a digest over every `(source_offset, logical_offset, len)` and `(start, end)` tuple. |
| `sst-index-after.txt` | The same listing on this branch. |
| `sst-index-diff.txt` | The `diff` of the two: empty. 121 fixtures, identical line for line, refusal messages included. |
| `gates.txt` | The tail of `cargo fmt --all --check`, `cargo clippy -p litchi-xls --all-targets`, `cargo test -p litchi-xls`, `cargo doc -p litchi-xls --no-deps`, and change 0576's corpus differential. |
| `quiescence.log` | Host load at the two ends of the timing window. |
| `decision.json` | What was accepted as evidence, what was accepted as cost, what is withheld. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, for the coordinator to merge. |
| `capture_counters.sh` | The logical-counter driver. |
| `capture_callgrind.sh` | The callgrind isolation-pair driver. Unlike change 0576's, it profiles `one-cell` as well as `open`, because half of this change lands in the worksheet frame loop, which `open` never enters. |
| `capture_perf.sh` | The hardware-counter driver. |
| `capture_latency.sh` | The A1 B1 B2 A2 wall-clock driver. |

## Result

One source-backed open of `54016.xls` falls from **253,317 ns to 193,739 ns**
(−23.5%), from **7,240,329 to 5,375,911 instructions** under callgrind (−25.8%)
and from **1,165,313 to 869,124 cycles** natively (−25.4%). One `one-cell` query
on the same fixture falls from **1,011,306 ns to 832,605 ns** (−17.7%) and from
**24,676,105 to 18,602,879 instructions** (−24.6%). `WithCustomViews.xls` falls
6.2 to 13.0% at p50 on every operation and mode.
`ConditionalFormattingSamples.xls` falls 2 to 5% in counts but its wall clock is
inside the measured floor and no speedup is claimed for it.

The measured same-binary noise floor in this window is **−2.64% to +3.06%** at
p50 across all 36 comparisons.

`drop_in_place<SourceBackedError>` and `SstCursor::read_exact` are **exactly
zero** on all three fixtures after the change. Every logical read counter, and
every one of the harness's semantic projections, is unchanged.

The SST index is byte-identical across the change: **121 fixtures**, digest
`0x9cb14f5daa02eebc`, pinned in `the_sst_index_over_the_corpus_is_pinned` from a
value captured before any production edit.

## Replay

```sh
python3 -B docs/performance/results/change-0595/analyze.py \
  docs/performance/results/change-0595
```

## Rebuilding the captures

Both legs are the standalone `tools/perf-baseline` package. The before leg is
built from a detached checkout of the base placed **outside** the repository, so
that neither leg can pick up unrelated working-tree changes and neither build
touches a shared target directory:

```sh
BEFORE=/some/path/before-08d968f8e
git -C . worktree add --detach "$BEFORE" 08d968f8e
CARGO_TARGET_DIR=/some/path/targets/0595-before \
  cargo build --manifest-path "$BEFORE/tools/perf-baseline/Cargo.toml" \
  --release --locked --features xls-source-attribution --bin xls_source_attribution

# the after leg, from a worktree of this branch
cargo build --manifest-path tools/perf-baseline/Cargo.toml \
  --release --locked --features xls-source-attribution --bin xls_source_attribution

OUT=/some/scratch
CPU=15 docs/performance/results/change-0595/capture_counters.sh  "$BEFOREBIN" before "$OUT/counters"
CPU=15 docs/performance/results/change-0595/capture_counters.sh  "$AFTERBIN"  after  "$OUT/counters"
CPU=15 docs/performance/results/change-0595/capture_callgrind.sh "$BEFOREBIN" before "$OUT/callgrind"
CPU=15 docs/performance/results/change-0595/capture_callgrind.sh "$AFTERBIN"  after  "$OUT/callgrind"
CPU=15 docs/performance/results/change-0595/capture_perf.sh      "$BEFOREBIN" before "$OUT/perf"
CPU=15 docs/performance/results/change-0595/capture_perf.sh      "$AFTERBIN"  after  "$OUT/perf"
CPU=15 W=50 S=500 docs/performance/results/change-0595/capture_latency.sh \
  "$BEFOREBIN" "$AFTERBIN" "$OUT/latency"

python3 -B docs/performance/results/change-0595/analyze.py "$OUT" --fold
python3 -B docs/performance/results/change-0595/analyze.py "$OUT"
```

The SST index listings come from the pinned-digest test:

```sh
cargo test -p litchi-xls --lib the_sst_index_over_the_corpus_is_pinned -- --nocapture
```

The test is part of this change, so it does not exist at `08d968f8e`.
`sst-index-before.txt` was produced by adding **only the test** to a worktree
whose two production files were still the base's, running it there, and then
applying the production change; the constant `CORPUS_DIGEST` is the value that
run printed. To reproduce it, `git checkout 08d968f8e -- crates/litchi-xls/src/records.rs
crates/litchi-xls/src/workbook/source.rs`, re-add the test body, run, and restore.

## What this evidence does not carry

No cold-cache, physical-device, range-source, peak-RSS, allocation-profile,
concurrency-scaling, real-producer or cross-platform result. No XLS full-text or
all-cells figure: `xls_source_attribution` has no such selector, which is change
0587's recorded measurement blocker, so the two edited sites in
`SourceTextSheet::insert` are modelled rather than measured. No eager-owner
figure. Every XLS source in this harness is in memory, so a `from_path`
workbook's per-observation `fstat` cost is invisible here.
