# 0767 evidence: CFB open and read clear chain maps in proportion to the chains

[Record](../../0767-cfb-reparse-linear.md). `performance_claim: none`. The
numbers are evidence, not registered claims.

## Arms

| Arm | Source | What it is |
|---|---|---|
| base | `1d1044e3ac` | The base: A5 clears a table-sized chain map per stream; `open_stream` allocates one per call. |
| cand | `573554acfe` | The retained change: maps kept all-clear between walks, restored by word for short chains and in bulk otherwise; `OleFile` retains one chain scratch for its reads. |
| cand-v1 | `af72192785` | The superseded first version (per-bit restoration whenever the chain is shorter than the table's words); only `matrix-v1/` and `counters-v1/` measured it. |

`binaries.json` has every measured binary's SHA-256 and size, the build
commands and the lockfile digests. Both arms were built with identical cargo
commands, flags and features under rustc 1.95.0. The probe manifests differ
only in the source tree their dependency paths point at, and
`--remap-path-prefix` maps both trees to `/litchi`. `inputs.json` has the
generated inputs' digests and the command that regenerates each.

## Directory

- `probe/`: the probe source (`src/`, `src/verdicts.rs` for the fault lane),
  its two manifests (`manifests/{before,after}/Cargo.toml`) and the one
  lockfile both used (`manifests/Cargo.lock`; copy it into each manifest
  directory before building).
- `scripts/`:
  - `build-probe.sh`, `build-harness.sh`: the identical per-arm builds.
  - `run.py`: the ABBA timing matrix (cases, samples, fixtures).
  - `analyze.py`: per-process statistics, paired changes, bootstrap
    intervals, flags and the output-digest oracle (0749's).
  - `counters.py`: per-owner `perf stat` counters, two sample counts at
    three layouts.
  - `callgrind_scaling.py`: Callgrind of one timed owner per arm, mode and
    generated file.
  - `alloc.py`: the counting-allocator lane.
  - `package.py`, `summarize.py`: this packet and `summary.md`.
  - `gates.sh`: the gate commands.
- `matrix/`: the retained timing matrix, 39 cases × 2 arms × 12 rounds (936
  processes, none failed):
  - `analysis.json`: the statistics;
  - `commands.json`: exact commands, exit codes and start times;
  - `reports.tar.zst`: every process report and stderr;
  - `manifest.json`: each member's SHA-256 and size.
- `matrix-v1/`: the same for the first version, 8 rounds (624 processes).
- `counters/`, `counters-v1/`: `counters.json` (per-owner summary and every
  row) and `perf-stat.tar.zst` (every raw `perf stat -x,` output).
- `scaling/`: `callgrind-scaling.json` (48 owners: rows and scale ratios) and
  `inclusive-heads.tar.zst` (the head of each inclusive listing). The raw
  Callgrind outputs were deleted.
- `profiles/`: `base-inclusive-heads.tar.zst`, the base profiles that located
  the two terms (open and read-all at 1,000, 3,000 and 10,000 streams).
- `allocation/`: `allocation.json` and the raw `probe_alloc` reports.
- `verdicts/`: the cross-build fault lane. `verdicts.tar.zst` holds the 104
  JSON-line outputs, one copy, because the base's and the candidate's are
  byte-identical; `manifest.json` records both arms' SHA-256 for every
  file; `verdicts-summary.json` has the totals and each input's digest.
- `summary.md`: every timing, flag, counter, Callgrind and allocation row.
- `binaries.json`, `inputs.json`, `gates.txt`, `cleanup.json`,
  `log-sections.md`.

## Method

Every process is pinned with `taskset -c 28` on the 32-core AMD EPYC 9R45,
with warm caches, while other agents build on other cores.

**Rounds.** Round r runs every case once per arm, with the arm order cycling
base/cand, cand/base, cand/base, base/cand, and the case order rotated.
Each binary starts through a symlink whose path grows 8 bytes per round, so
both arms of a round share one argv[0] heap layout and the rounds sample
many layouts (0745's method).

**Statistics.** Process p50 is the midpoint median, and p95 is the nearest
rank. A comparison is the median of the per-round paired changes (cand/base
− 1). Its interval is a percentile bootstrap of that median: 10,000
resamples, seed 767.

**Oracles.**

- Every probe process digests every owner's output, and all digests of one
  case agree across all processes of both arms.
- The 45543.ppt Reuse write publishes the sealed 0728 digest
  `545ed7e5ba7b8cc0687f8d21ac95aca68fa4d67fd89fd4277ed7cce047bdb859`.
- The harness checks its own outputs and fails the process otherwise.
- The verdict lane's outputs are compared byte for byte across the builds.

## Replay

Build the binaries as `binaries.json` describes, with
`scripts/build-probe.sh` and `scripts/build-harness.sh`, and copy them to
`scratch/0767/bin/{base,cand}`. Regenerate the inputs into
`scratch/0767/gen/` with the commands in `inputs.json`, and check their
digests. Then, from the scratch root with this packet's `scripts/`:

```sh
python3 scripts/run.py matrix 12
python3 scripts/analyze.py matrix > matrix-analysis.json
python3 scripts/counters.py counters > counters.json
python3 scripts/callgrind_scaling.py scaling > callgrind-scaling.json
python3 scripts/alloc.py allocation > allocation.json
# the verdict lane, per arm and input (inputs listed in verdicts-summary.json)
bin/ARM/probe --mode verdicts --input FILE --cases 300 --seed 767 > verdicts/ARM/NNN.jsonl
python3 scripts/package.py . PACKET && python3 scripts/summarize.py PACKET > PACKET/summary.md
```

The generated files take 100 fault cases each (`--cases 100`).

Timings are specific to this host, its load and these inputs. A replay
should compare the paired changes, not the absolute microseconds.
