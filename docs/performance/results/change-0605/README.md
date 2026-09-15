# Evidence: change 0605, one validated scan for a whole XLS worksheet

Change record:
[`0605-xls-retained-sheet-index.md`](../../0605-xls-retained-sheet-index.md).

Disposition: **part (1) retained** (the whole-sheet walk and the three harness
scenarios); **part (2) design only** (the snapshot-scoped retained sheet index is
specified and not implemented). `performance_claim: none`.

## Provenance

| | |
| --- | --- |
| Base commit | `01a5f5731` — `perf(xlsb): parse the workbook once per cell-value commit (0599)` |
| Branch | `perf/0605-xls-retained-sheet-index` |
| Files that differ between the legs | `crates/litchi-xls/src/workbook/source.rs`, `crates/litchi-xls/tests/source_backed.rs`, `tools/perf-baseline/src/bin/xls_source_attribution.rs`, `tools/perf-baseline/README.md` |
| Before binary | `xls_source_attribution`, sha256 `d594674deffb3e686ceb920396f0f1352a78583dec060ad52e394abee29c48e9`, built from the shared read-only checkout `litchi-worktrees/before-01a5f5731` |
| After binary | `xls_source_attribution`, sha256 `b22152a6cd59637663d6bc2acb4dab25f716b5c98b90f25ff899fee67da20b45` |
| Build | `cargo build --release --locked --bin xls_source_attribution --features xls-source-attribution`, identical flags on both legs |
| Host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| Toolchain | rustc 1.95.0, valgrind 3.26.0, perf 7.0.14 |
| Pinning | CPU 25, `setarch x86_64 -R` (ASLR off), `RAYON_NUM_THREADS=1`, single-threaded |
| Quiescence | **not established**; eight measurement agents shared the host, 1-minute load average 8.27 at the end of the timing window |

Full identities, the three fixture SHA-256s and the per-fixture notes that
explain which worksheet each scenario uses are in
[`environment.json`](environment.json).

## Contents

| Path | What it is |
| --- | --- |
| `analysis.txt` | The exact output of `analyze.py` over this directory. Every number the record cites appears here. |
| `analyze.py` | Recomputes all five tables from this directory alone, and folds a raw capture (`--fold`). Pure standard library. |
| `environment.json` | Host, toolchain, leg identities, binary and fixture hashes, harness scope, and which worksheet each scenario uses. |
| `counters-summary.json` | Deterministic logical counters: 3 fixtures × 2 in-memory source modes × 7 operations, before and after where both legs have the operation, with each leg's semantic projection. Every sample within a cell carries identical counters, which `analyze.py` asserts while folding. |
| `walk-summary.json` | The all-cells scenario's paired legs: the walk against a per-cell reading at 8, 64 and 256 positions, on the worksheet of each fixture that actually stores cells, both source modes, with counters and folded elapsed statistics. |
| `latency-summary.json` | The folded A1 B1 B2 A2 wall-clock capture: 4 rounds, 42 after-leg cells of which 18 exist on both legs, each with n, p50, mean, p95, p99 and the SHA-256 of the binary that produced it. |
| `callgrind/ann-<leg>-<fixture>-<label>-s{small,large}.txt` | `callgrind_annotate --threshold=99.9` self cost. Differencing the large- and small-sample child and dividing by the extra operations isolates one operation: change 0574's method, as 0576, 0584 and 0595 used it. |
| `callgrind-trial/ann{,2,3}-{5,25}.txt` | The three trial profiles behind *A regression that had to be designed away*: `ann` is the shared loop with `#[inline]` hints (18,817,353 Ir), `ann2` with the original nested branch shape restored (18,817,347), `ann3` with the `ScanContext` that landed (18,603,993). The before leg is 18,601,973, in `callgrind/ann-before-54016-one-cell-*`. |
| `perf/perf-<leg>-<fixture>-<label>-s<n>.csv` | `perf stat -x,` cycles, instructions, branches, branch misses and task-clock, isolated the same way. Native counters, because callgrind counts `rep movsb` per byte and runs SHA-256 in software. |
| `corpus-before.txt`, `corpus-after.txt`, `corpus-diff.txt` | The corpus differential: all 126 `.xls` fixtures under `test-data` x open/list/one-cell x worksheet indices 0 and 1, 756 cells per leg, each carrying that cell's logical counters, the harness's source and eager semantic projections, and the typed refusal where the reader declines a fixture. The diff is **empty**, 99 refusals with identical messages included. |
| `walk-differential.txt` | `--operation all-cells` over the same 126 fixtures at worksheet indices 0, 1 and 2. Each run's oracle walks the worksheet, re-reads up to 128 of the reported positions one selected-cell query at a time, and fails the run on any disagreement. 291 worksheets walked, 90,543 cells reported, **zero disagreements**. |
| `capture_corpus.sh` | The corpus-differential driver. Runs unchanged on both legs; nulls are dropped from the projections so the after leg's two extra schema fields do not read as a difference. |
| `capture_walk_differential.sh` | The corpus-wide walk-against-queries driver. |
| `gates.txt` | The tail of `cargo fmt --all --check`, `cargo clippy` on both touched crates, `cargo test` on both, and `cargo doc -p litchi-xls --no-deps`. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, for the coordinator to merge. |
| `decision.json` | What was accepted as evidence, what was accepted as cost, what is withheld. |
| `capture_counters.sh` | The logical-counter driver. Skips an operation the leg's binary does not have, so the same script drives both legs. |
| `capture_walk.sh` | The all-cells paired-leg driver, at the worksheet of each fixture that stores cells. |
| `capture_callgrind.sh` | The callgrind isolation-pair driver. |
| `capture_perf.sh` | The hardware-counter driver. |
| `capture_latency.sh` | The A1 B1 B2 A2 wall-clock driver. |

The raw per-sample JSON (9 MB across 106 files) is **not** retained: the folded
summaries carry every number the record cites, and `analyze.py --fold` rebuilds
them from a fresh capture.

## Result

`SourceBackedWorksheet::visit_cells` reports every stored cell of a worksheet
from one validated scan. On `54016.xls` that is **38,950 cells for 16,145 reads
and 1,256,139 bytes**, against **157,972,311 bytes for 256 cells** read one
selected-cell query at a time — 32.3 bytes per cell against 617,079.3, a factor
of **19,134**, and 9,146 instructions per cell against 13,330,564, a factor of
**1,457**. The crossover is about three dozen cells on every fixture measured.

Nothing else moved. Reads, read bytes, source-version observations, `len` calls
and seeks are identical to the byte on `open`, `list` and `one-cell`, on all
three fixtures and both source modes, and the instruction counts of those
operations moved by at most **+0.146%**. The corpus differential extends that to
**all 126 `.xls` fixtures** — 756 cells per leg, `diff` empty, 99 typed refusals
with identical messages — and the walk was differentiated against selected-cell
queries over **291 worksheets and 90,543 cells** with zero disagreements.

The snapshot-scoped retained sheet index is designed in the record and not
implemented: its measured weight is 623 KB-935 KB for **one** worksheet of a
984 KB fixture, and ADR 0005's weighted, bounded, evictable cache has no
implementation anywhere in this repository to build on.

## Reading the noise

The A/A floor in this window is unusually good — p50 0.71%, mean 0.73%, max
2.14% over the 18 cells that exist on both legs — but the B/B floor over all 42
after-leg cells reaches **46.51%**, on the `54016` walk, where 30 samples of a
20 ms operation on a busy host is simply not enough. Two paired comparisons
exceed the +5% review trigger, both on `WithCustomViews.xls` `file-source` and
both on `open` and `list`, which this change does not touch; the record states
them with the same-cell B/B floor beside them rather than folding them into a
mean.
