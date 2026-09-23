# 0749 evidence: CFB Reuse-plan validation compares streams in place

[Record](../../0749-cfb-reuse-plan-validation.md). `performance_claim: none`.
The numbers are evidence, not registered claims.

## Arms

| Arm | Source | What it is |
|---|---|---|
| base | `ab29ac6291` | Branch tip with record 0745. `ReusePlan::validate` reparses the planned view and reads every stream back through `OleFile::open_stream`. |
| cand | `897d24fd0f` | The same reparse; the readback goes through `OleFile::stream_equals`, comparing each range in place. `776ad05175` adds tests only and changes no binary. |

`binaries.json` records the SHA-256 and size of every measured binary (probe,
probe_alloc, litchi-perf-baseline) per arm and the exact build commands. Both
arms were built with identical cargo commands, flags and features under rustc
1.95.0; the probe manifests differ only in the source tree their dependency
paths point at, and `--remap-path-prefix` maps both trees to `/litchi`.

## Directory

- `probe/`: the 0749 probe source (`src/`), its two manifests
  (`manifests/{before,after}/Cargo.toml`) and the one lockfile both used
  (`manifests/Cargo.lock`). `src/lib.rs` documents every timed region.
- `scripts/`:
  - `build-harness.sh`: the identical harness build for both arms.
  - `run.py`: the ABBA timing matrix (`ROUND_START` continues it).
  - `analyze.py`: per-process statistics, paired changes, bootstrap
    intervals, flags and the output-digest oracle.
  - `counters.py`: the per-owner `perf stat` lane.
  - `counters_read_controls.py`: the same lane for the harness read
    controls.
  - `pagefault_layouts.py`: page faults across ten layouts, with or without
    `GLIBC_TUNABLES`.
  - `alloc.py`: the counting-allocator lane.
  - `profile_phases.py`: phase × leaf attribution of a frame-pointer
    `perf record` restricted to the timed owner.
  - `package.py`, `summarize.py`: packet assembly and `summary.md`.
  - `gates.sh`: the gate commands.
- `matrix/`: every process report of the timing matrix, 22 cases × 2 arms ×
  8 rounds, with 7 cases continued to 18 rounds (492 processes, none
  failed).
  - `commands.json` and `commands-from-r08.json`: exact commands, exit codes
    and start times.
  - `analysis.json`: the statistics.
  - `reduction-manifest.json`: each report's original SHA-256 and size. The
    reports are kept whole but re-serialized without whitespace.
- `counters/`: per lane, the summary (`<lane>.json`) and every raw
  `perf stat -x,` output verbatim (`<lane>-perf-stat.json`). The lanes:
  - `per-owner`: probe cases and harness edit/save selectors;
  - `read-controls`: harness `cfb_open` and `cfb_read_one`;
  - `page-fault-layouts` and `page-fault-layouts-glibc-tuned`:
    `doc_semantic_one_edit_save` across ten layouts;
  - `per-owner-v1-digest-diluted`: the superseded first pass, which also
    digested every owner's output.
  - `direct-runs/`: the unsymlinked supplementary runs the record cites
    (front-end counters for `cfb_open/few-large`, page-fault record logs and
    `strace` counts), described in their own README.
- `allocation/`: `allocation.json` and each raw `probe_alloc` report.
- `profiles/`: phase and leaf summaries of frame-pointer `perf record`
  profiles (`perf-phases-*`), and Callgrind inclusive listings of the Reuse
  `write_to` on 45543.ppt (`callgrind-inclusive-*`). Raw `perf.data` and
  Callgrind outputs were deleted after summarizing.
- `tuned/`: the timing check under fixed glibc malloc thresholds (see the
  record).
- `summary.md`: every timing, counter and allocation table, and every flag
  above +5%.
- `gates.txt`, `cleanup.json`, `log-sections.md`.

## Method

Every process is pinned with `taskset -c 20` on the 32-core AMD EPYC 9R45,
with warm caches, while other agents build on other cores.

**Rounds.** Round r runs every case once per arm, and the arm order
cycles base/cand, cand/base, cand/base, base/cand. The case order rotates.
Each binary starts through a symlink whose path grows 8 bytes per round, so
both arms of a round share one argv[0] heap layout and the rounds sample
many layouts. This is change 0745's method.

**Statistics.**

- Process p50 is the midpoint median, and p95 is the nearest rank.
- A comparison is the median of the per-round paired changes (cand/base −
  1).
- Its interval is a percentile bootstrap of that median: 10,000 resamples,
  seed 749.

**Oracles.**

- Every probe process digests every owner's output, and all digests of one
  case agree across all processes of both arms.
- The 45543.ppt removal publishes the sealed 0728 digest
  `545ed7e5ba7b8cc0687f8d21ac95aca68fa4d67fd89fd4277ed7cce047bdb859`.
- The harness checks its own outputs and fails the process otherwise.

## Replay

Rebuild the binaries as `binaries.json` describes into
`/home/zhuhe/code/litchi-worktrees/scratch/0749/bin/{base,cand}`, where
`run.py` looks for them. Then run from the scratch root, with this packet's
`scripts/` there:

```sh
python3 scripts/run.py matrix 8
ROUND_START=8 python3 scripts/run.py matrix 10 ppt-remove doc-replace harness-ole-edit-save harness-cfb-read cfb-write-rewrite-45543
python3 scripts/analyze.py matrix > analysis.json
python3 scripts/counters.py counters-raw > counters.json
python3 scripts/counters_read_controls.py counters-read-raw > counters-read.json
python3 scripts/pagefault_layouts.py pagefaults-raw > pagefaults.json
GLIBC_TUNABLES=glibc.malloc.trim_threshold=268435456:glibc.malloc.mmap_threshold=268435456 \
  PF_CASES=doc_semantic_one_edit_save/large python3 scripts/pagefault_layouts.py pagefaults-tuned-raw > pagefaults-tuned.json
GLIBC_TUNABLES=glibc.malloc.trim_threshold=268435456:glibc.malloc.mmap_threshold=268435456 \
  python3 scripts/run.py matrix-tuned 8 ppt-remove doc-replace-floating harness-ole-edit-save
python3 scripts/analyze.py matrix-tuned > analysis-tuned.json
python3 scripts/alloc.py allocation-raw > allocation.json
python3 scripts/package.py . packet && python3 scripts/summarize.py packet > packet/summary.md
```

Timings are specific to this host, its load and these fixtures. A replay
should compare the paired changes, not the absolute microseconds.
