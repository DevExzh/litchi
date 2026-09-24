# 0769 evidence: CFB open and every reader agree on the partial last mini sector

[Record](../../0769-cfb-mini-sector-open-read-agreement.md). `performance_claim: none`.
The numbers are evidence, not registered claims.

## Arms

| Arm | Source | What it is |
|---|---|---|
| base | `97e558cef4` | Record 0767's head (under review): A5 admits any mini sector that begins inside the root size; the whole-stream readers require the whole 64-byte sector inside it. |
| cand | `ebf20bba87` | The retained change: A5 and every reader bound the bytes a stream takes; A5's ownership loop as in the base, with the partial sector checked once per stream. `47d19bd609` and `ebf20bba87` add a test and the refinement to `d13c52a45f`; the 0767 review follow-ups are `97bf5f8f82`. |
| cand-v1 | `d13c52a45f` | The first version (the bound inside the ownership loop); only `matrix-v1/` and `counters-v1/` measured it. |

`binaries.json` has every measured binary's SHA-256, size and `.text` size,
the build commands and the lockfile digests. Both arms were built with
identical cargo commands, flags and features under rustc 1.95.0; the probe
manifests differ only in the source tree their dependencies point at, and
`--remap-path-prefix` maps both trees to `/litchi`. `inputs.json` has every
input's digest: the twelve generated files are record 0767's, regenerated
byte for byte.

## Directory

- `probe/`: the probe source (record 0767's, plus `src/census.rs` for the
  `census` and `root-sweep` lanes), its two manifests and the one lockfile
  both used (copy `manifests/Cargo.lock` into each manifest directory before
  building).
- `scripts/`:
  - `build-probe.sh`, `build-harness.sh`: the identical per-arm builds.
  - `census.py`: an independent CFB reader and the root-size census over the
    1,830 compound files (no litchi code).
  - `normalize_zvi.py`: derives the lanes' `gen/zvi-normalized.cfb`.
  - `lanes.sh`, `analyze_lanes.py`: the correctness lanes and their
    cross-arm comparison, including the independent byte-bound oracle.
  - `run.py`, `analyze.py`: the ABBA timing matrix and its statistics
    (record 0767's).
  - `counters.py`: per-owner `perf stat` counters (record 0767's method).
  - `alloc.py`: the counting-allocator lane.
  - `callgrind_summary.py`: the `doc_semantic_open/large` Callgrind summary.
  - `summarize.py`: renders `summary.md`.
  - `gates.sh`: the gate commands.
- `census/`: the independent census (`independent-census.json.zst`, one row
  per file; `independent-census-summary.json`, the counts and the three files
  whose root size is not a multiple of 64) and the lanes' path lists.
- `lanes/`: `census.tar.zst` (both arms' reader-agreement census over 1,831
  files), `sweep.tar.zst` (both arms' root-size sweep, 21,741 copies),
  `verdicts-cand.tar.zst` (the candidate's outputs of record 0767's fault
  lane; the base arm's are byte-identical to record 0767's packet
  `verdicts/verdicts.tar.zst`, which `analyze_lanes.py` checks), and
  `analysis.json`.
- `matrix/`, `matrix-v1/`: `analysis.json`, `commands.json` (exact commands,
  exit codes, start times) and `reports.tar.zst` (every process report and
  stderr): 27 cases × 2 arms × 16 rounds each, 608 processes, none failed.
- `counters/`, `counters-v1/`: `counters.json` and `perf-stat.tar.zst`
  (every raw `perf stat -x,` output).
- `allocation/`: `allocation.json` and the raw `probe_alloc` reports.
- `callgrind/doc-semantic-open-large.json`: program and DOC-parse totals per
  arm, every `litchi_cfb` function's inclusive cost, and the largest
  per-function self-cost changes. The raw Callgrind outputs were deleted.
- `summary.md`: every table.
- `binaries.json`, `inputs.json`, `gates.txt`, `cleanup.json`,
  `log-sections.md`.

## Method

**Correctness lanes** (`scripts/lanes.sh`, every process on core 30):

- *census*: every compound file under `test-data/` and the gitignored
  `3rdparty/` POI, LibreOffice and Open-XML-SDK corpora (1,830), plus the
  normalized Zeiss file. Per file: both opens, then every listed stream
  through six readers (the cursor reader's whole read and whole range, a
  fresh shared reader per stream, one shared reader for all streams, the
  shared reader's range and its cursor).
- *fault lane*: record 0767's `verdicts` mode, unchanged, on its 104 inputs
  with the same case counts and seed.
- *root-size sweep*: for each of those 104 inputs and the normalized Zeiss
  file, the root entry's stream size rewritten to every value within 128
  bytes of its own (from 1), each copy read as in the census without the
  fresh shared reader per stream.

`analyze_lanes.py` compares the arms line by line, classifies every
difference, and holds the candidate's sweep verdicts to an oracle computed
by `census.py` from each file's own metadata: admitted exactly when the
root chain keeps its length and every byte every mini stream takes lies
below the root size.

**Timing** (`scripts/run.py`): 16 rounds of 27 cases, ABBA across rounds
with rotated case order; each binary starts through a symlink 8 bytes longer
each round (record 0745's layout sampling); every process pinned with
`taskset -c 30`. Statistics are record 0767's: process p50 as the midpoint
median, a comparison as the median of per-round paired changes, and a
10,000-resample bootstrap interval (seed 769).

**Counters**: per-owner `instructions:u` and `cycles:u` from two processes of
different sample counts at three argv[0] layouts; their difference removes
process setup. Harness rows still include each iteration's untimed setup
(record 0767), so their cycle deltas do not measure the timed region.

**Callgrind**: the harness's `doc_semantic_open/large` with 3 samples and no
warmup, per arm, whole program.

**Allocation**: one counting-allocator process per probe case and arm, one
warmup and five measured owners.

## Replay

Build the binaries as `binaries.json` describes, with
`scripts/build-probe.sh` and `scripts/build-harness.sh`, and copy them to
`scratch/0769/bin/{base,cand}`. Regenerate the inputs into `scratch/0769/gen/`
with the commands in `inputs.json` (and `normalize_zvi.py`), check their
digests, write the path lists from `census/`, then from the scratch root:

```sh
bash scripts/lanes.sh && python3 scripts/analyze_lanes.py > lanes-analysis.json
python3 scripts/run.py matrix 16 && python3 scripts/analyze.py matrix > matrix-analysis.json
python3 scripts/counters.py counters > counters.json
python3 scripts/alloc.py allocation > allocation.json
python3 scripts/summarize.py PACKET > PACKET/summary.md
```

`analyze_lanes.py` expects record 0767's fault-lane outputs extracted into
`v0767/`. Timings are specific to this host, its load and these inputs; a
replay should compare the paired changes, not the absolute nanoseconds.
