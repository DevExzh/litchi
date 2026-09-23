# Change 0742 evidence packet — owned PPTX cross-copy media transfer

Record: [0742](../../0742-pptx-owned-cross-copy-media-transfer.md).

Base `009d515bef`; production commits `317920af5c`, `52db88c24c` and
`b2132486af` (first review), `172501ac89` and `d2b2aa3d75` (second review),
`34255fea84` and `ddefc2cfe8` (third review) on
`perf/0742-pptx-owned-cross-copy-media-transfer`. The reported matrix measured
`d2b2aa3d75`; the third review's commits add only constant-time checks to the
measured path and were not measured (see the record). Host: AMD EPYC 9R45, Linux
7.0.0-1012-aws, 32 logical CPUs shared with other agents; every measured
process pinned with `taskset -c 4`. Toolchain from `rust-toolchain.toml`
(Rust 1.95.0); release profile, `--locked --offline`, `CARGO_BUILD_JOBS=6`.
The harness records the runtime default `rustc` (1.98.1) in its environment
block; that is not the build toolchain.

## Binaries of the reported matrix (removed after the run; identities retained)

Both legs were built with the identical command,
`cargo build --release --locked --offline --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline`
(and `--features allocator-metrics --bin litchi-perf-baseline-alloc` for the
allocator lane), the before leg from the read-only base checkout into
`targets/0742-before`, the after leg from `d2b2aa3d75` into
`targets/0742-after`.

| arm | lane | SHA-256 |
|---|---|---|
| before | native | `b375de3695473a3f0bf870ae25d9c2e7a8f9b8124aa43370698f6557a75f3385` |
| before | alloc | `75c5ee1707d0f45ee2366af3e77f26b9748cb34474e14550cea44e566eeddab2` |
| after | native | `9d714e54b1f34a13d1778ab34a0f7ae41711ffab41760f7864e796db12521d2e` |
| after | alloc | `d8ff3f80738cf97aa27ba0fc27d6e7a51d7bc2a2f2b207df2ab90d4a34391434` |

The before binaries were rebuilt for this matrix with the same command from
the same base tree as the superseded `b2132486af` matrix's, and their hashes
differ from that matrix's (`0f20b4d0…`/`3e304e2d…`); the reason was not
investigated. Each matrix's pair was built in one session with one command.

The harness is unchanged by this change. The after reports say
`git_worktree_dirty: true` because the harness runs `git status --porcelain`
in its source tree at run time and this packet was being written there; the
measured binaries were built from clean committed trees.

## Superseded matrices

Three earlier matrices are kept as summaries (`analysis.json`, `tables.md`,
`faults.json`, `receipts-native.jsonl`, `receipts-alloc.jsonl`); their raw
reports were removed to keep the packet small ([`cleanup.json`](cleanup.json)):

- `superseded-317920af5c/` and `superseded-52db88c24c/` measured those commits
  against the coordinator's prebuilt base binary (`fb535ebb…`), which a later
  coordinator note showed can shift untouched paths by 2.7–3.4% relative to a
  before leg built with the after leg's exact command. After binaries:
  `3aef4c7f…`/`661034af…` (native/alloc, `317920af5c`) and
  `90840243…`/`2cd0153c…` (`52db88c24c`).
- `superseded-b2132486af/` is the first identical-command matrix, of the
  commit the second review examined: before `0f20b4d0…`/`3e304e2d…`, after
  `6237d295…`/`191dfeeb…` (native/alloc). The second review's fixes change
  the measured path (classification and capture now run before the candidate
  is built), so the matrix was repeated on `d2b2aa3d75`.

## Contents

- `run_abba.py` — the runner: per case, four rounds of before, after, after,
  before; every process a fresh pinned harness child. Commands, binary hashes,
  monotonic times and exit codes are in `raw/<lane>/receipts.jsonl`.
- `raw/native/<case>/r<round>-s<slot>-<arm>.json` — all 80 native reports
  (five cases × 16 processes): media-rich cases 20 samples after 3 warmups;
  plain and source-backed cases 40 samples after 3 warmups.
- `raw/alloc/<case>/…` — all 32 allocator-lane reports (two lifecycle cases ×
  16 processes, 3 samples after 1 warmup). Allocator binaries make no latency
  claim.
- `analyze.py` → `analysis.json`, `tables.md` — per-process p50/p95/mean,
  median of process p50s per arm with spread, paired after/before ratios
  ((s0 before, s1 after) and (s3 before, s2 after) in each round; eight pairs)
  with a percentile bootstrap of the median paired ratio (10,000 draws of the
  pairs, seed 742), phase medians and allocator counters.
- `faults.py` → `faults.json` — per-sample regrouping of the owned lifecycle
  and the source-backed control by the lifecycle's minor-fault count, which
  explains the spread of both arms (see the record).
- `size-guard/` — `zlib_framing.py` and its output `zlib-framing.txt`: the
  Deflate framing zlib 1.3.1 adds to 2 MiB of incompressible bytes at levels
  1/6/9 and memory levels 1–9, against what the review's prescribed guard
  unit and the chosen one allow.
- `legacy-fixture/` — the generator (built against the base tree with
  `BASE` replaced by a checkout of `009d515bef`) that wrote the committed
  `test-data/ooxml/pptx/cross-copy-legacy/lpcp0003-{forward,inverse}.patch`,
  and its `MANIFEST.tsv` of input, target and patch hashes.
- `route-compare/` — the informational probe that publishes one copy through
  the owned and the source-backed routes and diffs the archives member by
  member (`BRANCH` = this worktree), with its output `route-compare.txt`, run
  on `b2132486af`.
- `attribution/` — the optional phase attribution of a frame-pointer build of
  `b2132486af`: `attribute.py`, its summaries `cycles-depth3.json` and
  `cycles-depth5.json`, the page-fault site summary `faults-by-site.txt`, and
  the two harness reports of the profiled runs (no `perf.data`).
- `gates.txt` — every gate command with its exit code and counts, on the
  third review's commits and, below them, on `d2b2aa3d75` and on the first
  review's commits.
- `cleanup.py` → `cleanup.json` (third review), `cleanup-d2b2aa3d75.json`
  (second review) and `cleanup-b2132486af.json` (first review) — binary
  identities taken before removal and every removed tree.
- `log-sections.md` — paragraphs for `HOTSPOTS.md`, `REPORT.md` and
  `GOAL_AUDIT.md`, for the coordinator to merge.

## Replay

From the repository root (no binary needed):

```sh
python3 -B docs/performance/results/change-0742/analyze.py \
  docs/performance/results/change-0742/raw
python3 -B docs/performance/results/change-0742/faults.py \
  docs/performance/results/change-0742/raw
python3 -B docs/performance/results/change-0742/size-guard/zlib_framing.py
```

A new capture needs its own binaries and directory; the receipts show the
invocation used here.
