# Evidence: change 0608, the XLS shared-string walk re-priced and the lazy-index design frozen

Change record:
[`0608-xls-lazy-sst-index-design.md`](../../0608-xls-lazy-sst-index-design.md).

Disposition: **design only.** `performance_claim: none`. **No file under
`crates/` was modified.** The walk is 3.19% to 35.73% of an open in callgrind
instructions and 2.99% to 30.99% in native cycles on the three profiled
fixtures; the corpus census then shows that on **94 of 94** fixtures carrying
shared strings the highest SST index any cell references is the *last* entry, so
a prefix index saves nothing on a full text or an all-cells read, and the one
form in which no refusal moves is worth 0.44% to 3.13% of an open in cycles on
open-and-list while making every other scenario pay the walk twice. The design
is frozen and not implemented.

## Contents

| Path | What it is |
| --- | --- |
| `analysis.txt` | The exact output of `scripts/analyze.py` over `callgrind/` and `perf/`, followed by `scripts/fold_latency.py` over `latency-summary.json`. Every instruction, cycle and latency figure the record cites appears here. |
| `census-summary.txt` | The exact output of `scripts/fold_census.py` over `sst-prefix-census.jsonl`. Every census figure the record cites appears here. |
| `sst-prefix-census.jsonl` | One JSON object per fixture from the census probe: `cstUnique`, `cstTotal`, SST logical bytes, segment count, `LabelSst` cells, distinct indices, `idx_max`, the prefix a full text / a mean cell / the first cell must walk, the `isst` at each worksheet's row 1 column 0, and the header refusal for the four fixtures that have one. Ends with the probe's own summary line. |
| `probe/sst_prefix_census.py` | The census probe. Pure standard library; walks the CFB container, the BIFF8 framing, the SST header and every `LabelSst` record with arithmetic, sharing no code with `litchi-xls`, so it is an independent oracle. Its CFB/BIFF walk is adapted from change 0584's `results/change-0584/analysis/sst_walk.py`. |
| `fixtures.txt` | The 126 `.xls` and `.xlt` paths the census covers, as `find` produced them. Two contain spaces; the probe is driven with `xargs -0`. |
| `counters.txt` | The deterministic logical counters: 54 cells (3 legs × 3 fixtures × 2 in-memory modes × `open`/`list`/`one-cell`), each asserted identical across its own samples, with the harness's implementation projection of the worksheet count and the selected cell. The control for this record. |
| `callgrind/ann-<leg>-<fixture>-<op>-s{small,large}.txt` | `callgrind_annotate --threshold=99.9` **self** cost. |
| `callgrind/inc-<leg>-<fixture>-<op>-s{small,large}.txt` | `callgrind_annotate --inclusive=yes` from the same raw profile, which is what gives `scan_shared_string_records`'s share of the open. Differencing the large- and small-sample child and dividing by the extra operations isolates one operation: change 0574's method as 0576, 0584 and 0595 used it. |
| `perf/perf-<leg>-<fixture>-<op>-s{100,1100}-r<1..5>.csv` | `perf stat -x,` cycles, instructions, branches, branch misses and task-clock, isolated the same way, five repetitions per leg so the median can be taken before the pair is differenced. Legs `aa` and `bb` are byte-for-byte copies of the base binary and are the floor. |
| `latency-summary.json` | The folded A1 B1 B2 A2 wall-clock capture: 2 scaffolds × 3 fixtures × 2 modes × 4 rounds, each with n, p50, mean, p95, p99 and the SHA-256 of the binary that produced it. 400 samples per round after 50 warmups. The 24 raw round captures behind it are 11 MB of per-sample JSON and are **not** retained; `scripts/fold_latency.py` reads this summary directly and reprints every latency table in `analysis.txt` from it. |
| `latency/quiescence-*.log` | Host load at the two ends of each timing window. |
| `scripts/scaffold.py` | Applies either measurement scaffold to `crates/litchi-xls/src/records.rs`. |
| `scripts/scaffold-nostore.patch` | **Measurement scaffold, not a candidate.** Walks every shared string exactly as production does, reserves nothing and records nothing. Prices the storage alone. Applied, built, measured, reverted. |
| `scripts/scaffold-nowalk.patch` | **Measurement scaffold, not a candidate.** Builds the segments and runs every SST header check, then skips the per-string walk. Prices the ceiling of full deferral at open. Applied, built, measured, reverted. |
| `scripts/capture_counters.sh` | The logical-counter driver. |
| `scripts/capture_callgrind.sh` | The callgrind isolation-pair driver; unlike change 0595's it takes an inclusive annotation as well as a self annotation from each profile. |
| `scripts/capture_perf.sh` | The native hardware-counter driver, with repetitions. |
| `scripts/capture_latency.sh` | The A1 B1 B2 A2 wall-clock driver. |
| `scripts/analyze.py`, `scripts/fold_census.py`, `scripts/fold_latency.py` | Recompute every table from this directory alone. Pure standard library. |
| `environment.json` | Host, toolchain, leg identities and what each leg is, binary and fixture hashes, harness scope, and both measured floors. |
| `gates.txt` | The tails of `cargo fmt --all --check`, `cargo clippy -p litchi-xls --all-targets`, `cargo test -p litchi-xls` (aggregated over all 72 binaries) and `cargo doc -p litchi-xls --no-deps`, all on the clean worktree at the base commit. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, for the coordinator to merge. |

## Provenance

- Base commit: `818e58bee9dd50e6a064ba7cd0a9a7abbf44e82a`
  (`perf(xlsx): stop refusing merges the ineligible scan can no longer place …`).
- Branch: `perf/0608-xls-lazy-sst-index-design`.
- Working copy: a detached worktree at `/home/zhuhe/code/litchi-worktrees/0608`
  with its own `CARGO_TARGET_DIR`; the shared repository was neither built in nor
  modified. Nothing under `crates/` changed:
  `git diff 818e58bee -- crates/` is empty.
- Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws. Eight agents
  were building and measuring on the same machine throughout; the load average
  ran between 6.1 and 10.8, which is what both floors are for.
- Toolchain: rustc 1.95.0 (59807616e 2026-04-14), cargo 1.95.0, valgrind 3.26.0,
  perf 7.0.14.
- Measured CPU: 28, `taskset` pinned with `setarch x86_64 -R` for every counted
  and timed run.
- Legs, all `tools/perf-baseline`'s `xls_source_attribution` built
  `--release --locked --features xls-source-attribution`, differing only in
  `crates/litchi-xls/src/records.rs`:
  base `8fa24d70e4833aa3a855e7734b979124c037b1b0713a9c00adadd09c9bea8971`,
  `nostore` `ee47c51d506d3da5b79ebe805bcc5faa7b381a368a386988e0c226d62f19615a`,
  `nowalk` `f385c38b0509725ecb3249d5ff71cd39b070adc8706b3e00c90c44c4d4f4cea5`.
  The `aa` and `bb` floor legs are byte-for-byte copies of the base binary.
- Fixtures, from the repository's own corpus:
  `test-data/ole/xls/ConditionalFormattingSamples.xls`
  (`d1942d85…`, worksheet 1),
  `test-data/ole/xls/WithCustomViews.xls` (`3c0c168f…`, worksheet 1),
  `test-data/poi/test-data/spreadsheet/54016.xls` (`2e050f1f…`, worksheet 0).
  The census covers all 126 `.xls` and `.xlt` fixtures.

## Result

On this base, `scan_shared_string_records` costs **74,236 / 147,687 /
1,906,053** instructions per open on the flagship, `WithCustomViews.xls` and
`54016.xls` — **3.19% / 18.27% / 35.73%** of the operation, taken as its
inclusive callgrind cost. Natively, removing it removes **10,302 / 23,577 /
271,026** cycles per open — **2.99% / 18.32% / 30.99%** — measured as the
difference against the `nowalk` leg, because `perf stat` cannot attribute to a
symbol. The survey's retained 49.5% for `54016.xls` was a pre-0595 figure. Per
shared string the walk is 241.5 Ir against the 454 change 0587 retained for
the pre-0595 scan.

Splitting it with the two scaffolds: the **walk** is 3.18% / 18.09% / 35.75% of
the open and the **storage** — the two `logical_position()` calls, the
`entries.push` and the one reservation — is 0.30% / 1.22% / 2.09%.

The census decides the item. Over all 126 `.xls`/`.xlt` fixtures: 123 carry an
SST, 4 are refused by header checks (the same four change 0576 refuses,
identified here independently with the reason each fails), 119 are indexed at
open, 94 declare shared strings, and on **94 of 94** the highest SST index any
cell references is `cstUnique − 1`. The probe reproduces change 0576's corpus
totals exactly — 17,434 entries, the same 3 SST-less fixtures, the same 4
refusals, and 123/119 against 0576's 121/117 plus the two fixtures `litchi-cfb`
refuses before any SST is reachable — from code that shares nothing with the
scan. A deferred index is a prefix index, so a full text or an all-cells read
drives it to completion and saves nothing; a uniformly chosen string cell needs
a median 63.88% of the table; only open and list save the whole walk.

Both floors, measured in the same windows: **−0.17% to +0.66%** on the
isolation-pair cycle metric (six same-binary comparisons) and **−1.75% to
+1.70%** at p50 in wall clock (24 same-binary comparisons).

## Reproducing

```sh
# 1. the corpus census (no build required)
cd <repo>
find test-data \( -name '*.xls' -o -name '*.xlt' \) -print0 | sort -z \
  | xargs -0 python3 -B docs/performance/results/change-0608/probe/sst_prefix_census.py \
  > sst-prefix-census.jsonl
python3 -B docs/performance/results/change-0608/scripts/fold_census.py sst-prefix-census.jsonl

# 2. the three legs (edit scripts/scaffold.py's worktree path first)
cd <worktree>/tools/perf-baseline
CARGO_TARGET_DIR=<outside the repo> \
  cargo build --release --locked --features xls-source-attribution \
  --bin xls_source_attribution          # -> the base leg
python3 -B <packet>/scripts/scaffold.py nostore   # then rebuild -> the nostore leg
python3 -B <packet>/scripts/scaffold.py nowalk    # then rebuild -> the nowalk leg
git checkout -- crates/litchi-xls/src/records.rs  # always revert

# 3. counters, then instructions, then cycles, then wall clock, in that order
SC=<where the three binaries are> <packet>/scripts/capture_counters.sh counters.txt
CPU=28 <packet>/scripts/capture_callgrind.sh <bin> <leg> callgrind/
CPU=28 OPS=open REPS=5 <packet>/scripts/capture_perf.sh <bin> <leg> perf/
CPU=28 <packet>/scripts/capture_latency.sh latency/ <base-bin> <scaffold-bin> <name>

# 4. fold
python3 -B <packet>/scripts/analyze.py callgrind/ perf/
python3 -B <packet>/scripts/fold_latency.py latency/          # from raw rounds
python3 -B <packet>/scripts/fold_latency.py latency-summary.json  # from the packet
```

## What this packet does not establish

- **No speedup and no regression.** Nothing was implemented, so nothing was
  compared before against after. `base − nowalk` is a *ceiling*: it removes the
  walk without paying for the lock, the memoised refusal, the re-read of the SST
  bytes or the growth reservations that a real deferral would pay for.
- **The two scaffolds are controls, not candidates.** Both return an empty
  `entries`, so shared-string resolution is broken by construction on those legs
  — the flagship's `one-cell` reports `Invalid SST index: 3 (max: 0)` there,
  which is how the probe's prediction of `isst` 3 was cross-checked against
  production. Nothing about their correctness is claimed.
- **In-memory sources only.** Every XLS source in this harness wraps staged
  bytes or a staged immutable file; the page cache is warm and no physical I/O is
  measured. `xls_source_attribution` has no full-text, no all-cells and no
  `from_path`-observation selector, so the scenarios the frozen design would make
  *worse* cannot be timed at all, and its second read pass over the SST is
  modelled from byte counts.
- **Three fixtures, one host, one build, one CPU** for every timing and counter
  figure. The census alone is corpus-wide.
- **The refusal move is unexercised.** No fixture in the repository is refused by
  the per-string SST walk, so nothing here validates the design's error-identity
  argument; the synthetic malformed-SST differential its admission gates require
  does not exist.
- **Callgrind counts `rep movsb` per byte**, which is why the `nostore` leg's
  47,026 Ir `memcpy` artifact on `54016.xls` is named in the record rather than
  folded into its saving.
- Nothing here says anything about XLS-3, XLS-6, XLS-9, XLS-10 or the other
  items change 0587 ranks near this one.
