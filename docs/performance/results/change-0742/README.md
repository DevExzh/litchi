# Change 0742 evidence packet — owned PPTX cross-copy media transfer

Record: [0742](../../0742-pptx-owned-cross-copy-media-transfer.md).

Base `009d515bef`; production commits `317920af5c`, `52db88c24c` (review
follow-ups) and `b2132486af` (transfer-index allocation) on
`perf/0742-pptx-owned-cross-copy-media-transfer`. Host: AMD EPYC 9R45, Linux
7.0.0-1012-aws, 32 logical CPUs shared with other agents; every measured
process pinned with `taskset -c 4`. Toolchain from `rust-toolchain.toml`
(Rust 1.95.0); release profile, `--locked --offline`, `CARGO_BUILD_JOBS=6`.
The harness records the runtime default `rustc` (1.98.1) in its environment
block; that is not the build toolchain.

## Binaries of the reported matrix (removed after the run; identities retained)

Both legs were built with the identical command,
`cargo build --release --locked --offline --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline`
(and `--features allocator-metrics --bin litchi-perf-baseline-alloc` for the
allocator lane), the before leg from the read-only base checkout, the after
leg from `b2132486af`.

| arm | lane | SHA-256 |
|---|---|---|
| before | native | `0f20b4d07456ebb6493c8f70a11876cf2b88c93f6068dee5f33272f6c6004bf3` |
| before | alloc | `3e304e2d0b1d364fe25b0857d65dd35f20a9ddc70e986c476200d6c5910c6b09` |
| after | native | `6237d2950fb324821e0e51e5d496b71b882f8b6fbaa1f7ac619ee859e6cdacf1` |
| after | alloc | `191dfeeb042ced08aec52959ec8adf092bbaf5fbf76089693905b90d10f30bf1` |

The harness is unchanged by this change. The after reports say
`git_worktree_dirty: true` because the harness runs `git status --porcelain`
in its source tree at run time and this untracked packet was being written
there; the measured binaries were built from clean committed trees.

## Superseded matrices

Two earlier matrices measured `317920af5c` and `52db88c24c` against the
coordinator's prebuilt base binary (`fb535ebb…`), which a later coordinator
note showed can shift untouched paths by 2.7–3.4% relative to a before leg
built with the after leg's exact command. Their summaries are kept in
`superseded-317920af5c/` and `superseded-52db88c24c/` (`analysis.json`,
`tables.md`, `faults.json`, `receipts-native.jsonl`, `receipts-alloc.jsonl`);
their raw reports were removed
to keep the packet small ([`cleanup.json`](cleanup.json)). After binaries:
`3aef4c7f…`/`661034af…` (native/alloc, `317920af5c`) and `90840243…`/`2cd0153c…`
(`52db88c24c`).

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
- `legacy-fixture/` — the generator (built against the base tree with
  `BASE` replaced by a checkout of `009d515bef`) that wrote the committed
  `test-data/ooxml/pptx/cross-copy-legacy/lpcp0003-{forward,inverse}.patch`,
  and its `MANIFEST.tsv` of input, target and patch hashes.
- `route-compare/` — the informational probe that publishes one copy through
  the owned and the source-backed routes and diffs the archives member by
  member (`BRANCH` = this worktree), with its output `route-compare.txt`.
- `attribution/` — the optional phase attribution of a frame-pointer build:
  `attribute.py`, its summaries `cycles-depth3.json` and `cycles-depth5.json`,
  the page-fault site summary `faults-by-site.txt`, and the two harness
  reports of the profiled runs (no `perf.data`).
- `gates.txt` — every gate command with its exit code and counts.
- `cleanup.py` → `cleanup.json` — binary identities taken before removal and
  every removed tree.
- `log-sections.md` — paragraphs for `HOTSPOTS.md`, `REPORT.md` and
  `GOAL_AUDIT.md`, for the coordinator to merge.

## Replay

From the repository root (no binary needed):

```sh
python3 -B docs/performance/results/change-0742/analyze.py \
  docs/performance/results/change-0742/raw
python3 -B docs/performance/results/change-0742/faults.py \
  docs/performance/results/change-0742/raw
```

A new capture needs its own binaries and directory; the receipts show the
invocation used here.
