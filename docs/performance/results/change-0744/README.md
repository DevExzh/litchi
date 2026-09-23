# Change 0744 evidence packet

Record: [`../../0744-xlsx-eager-workbook-cell-path.md`](../../0744-xlsx-eager-workbook-cell-path.md).
Base `009d515bef`; code at `6ac6b92d8a`. `performance_claim: none`.

All processes ran on CPU 12 (`taskset -c 12`) of an AMD EPYC 9R45 host shared
with other agents. No binaries, `perf.data`, callgrind outputs or corpora are
retained; the binary digests are in `binaries.sha256`, `profile/binaries.sha256`
and `post-review/binaries.sha256`.

**Compiler.** Every measured binary was built by the repository's pinned Rust
1.95.0 (`rust-toolchain.toml`); the rebuilt post-review binaries' `.comment`
sections read `rustc version 1.95.0 (59807616e 2026-04-14)`. The
`environment.rustc` field of the raw reports says `rustc 1.98.1` because the
harness runs `rustc --version` at run time in its working directory, which was
outside the repository, where the ambient toolchain is 1.98.1. That field does
not describe the binaries.

## Legs

* **Before**: `009d515bef` built from a detached worktree whose path length
  equals the branch worktree's, with the identical command, features and
  profile as the after leg:
  `CARGO_TARGET_DIR=<target> CARGO_BUILD_JOBS=6 cargo build --release --locked
  --offline --manifest-path tools/perf-baseline/Cargo.toml --bin
  litchi-perf-baseline` (and `--features allocator-metrics --bin
  litchi-perf-baseline-alloc` for the allocator lane).
* **After**: the same commands in the branch worktree.
* The prebuilt base harness binary was used only by the superseded preliminary
  run in `abba/prelim-prebuilt-base/` (stopped after the coordinator's layout
  note; not used by the record's tables).

## Contents

| path | what |
| --- | --- |
| `abba/run_abba.sh` | primary ABBA driver: groups `dense`, `light`, `medium`, `controls`; order A1 B1 B2 A2 A3 B3 B4 A4. The primary run passed `A` and `B` explicitly; the script's default `A` is now the `0744-before` binary it used (`ac48ed6b…`), where it previously named the superseded prebuilt base binary |
| `abba/raw/*.json` | the 32 primary harness reports (`<group>-<slot>.json`) |
| `abba/analyze.py`, `abba/analysis.json` | per-process p50/p95/mean, per-arm median of process p50s, paired ratios, bootstrap 95% (20,000 resamples, seed 744), same-arm drift, output and corpus identity, >5% flags |
| `abba/run_confirm.sh`, `abba/confirm/*.json`, `abba/confirm-analysis.json` | confirmation ABBA for the flagged no-op pairs (2,000 samples/process) and the eager control (100 samples/process) |
| `abba/prelim-prebuilt-base/` | superseded preliminary run against the prebuilt base binary (15 processes) and its analysis |
| `alloc/run_alloc.sh`, `alloc/raw/*.json`, `alloc/alloc_summary.py`, `alloc/alloc-summary.json` | allocator-metrics lane, A1 B1 B2 A2, five samples after two warm-ups |
| `profile/run_profiles.sh`, `profile/attrib.py` | matched frame-pointer `profiling` builds of both legs; perf phase attribution and callgrind over the timed regions |
| `profile/{before,after}-{onecell,first}-phases.txt` | perf phase and leaf summaries (dense-wide one-cell commit+save, 30 samples; first cell, 100 samples) |
| `profile/{before,after}-{onecell,first}-callgrind-inclusive.txt` | top 120 lines of `callgrind_annotate --inclusive=yes` for the same regions |
| `profile/{before,after}-*.json` | the harness reports of the profiled processes |
| `output-identity/probe/` | standalone probe source; `Cargo.toml.in` is instantiated with each leg's `litchi-xlsx` path |
| `output-identity/sha256-{before,after}.txt` | digests of the 18 probe artifacts per leg (identical) |
| `post-review/run_remeasure.sh`, `post-review/raw/*.json`, `post-review/analysis.json` | dense-wide ABBA at `a755ff8c26` after the review fixes, both legs rebuilt by the identical command |
| `post-review/alloc/` | allocator lane at `a755ff8c26` (raw reports and summary) |
| `post-review/mutation/` | `mutate.py` reverts each review fix in turn and runs its targeted tests; `mutation-results.json` records that every revert is caught |
| `post-review/binaries.sha256` | the rebuilt native and allocator binaries of both legs |
| `gates.txt`, `run_gates.sh` | the gate runs with commands, exit codes and test counts: at `6ac6b92d8a`, and at `a755ff8c26` with a fresh target directory |
| `binaries.sha256` | native and allocator binaries of both legs, and the prebuilt base |
| `cleanup.json` | what was removed and what was kept |
| `log-sections.md` | ready-to-paste sections for `HOTSPOTS.md`, `REPORT.md` and `GOAL_AUDIT.md` |

## Replay

From `abba/`: `python3 analyze.py raw analysis.json` and
`python3 analyze.py confirm confirm-analysis.json`; for the post-review run,
`python3 ../abba/analyze.py raw analysis.json` from `post-review/`. From `alloc/`:
`python3 alloc_summary.py raw alloc-summary.json`. The probe builds with
`cargo build --release` after replacing `@XLSX@` in `Cargo.toml.in` and runs as
`xlsx-output-probe <out-dir>`.
