# 0745 evidence: lazy PPT artifact digests

[Record](../../0745-ppt-lazy-artifact-digests.md). `performance_claim: none`.
The numbers below are evidence, not registered claims.

## Arms

| Arm | Source | What it is |
|---|---|---|
| A | `009d515bef` | Base. Slide-order commit hashes two whole artifacts. |
| B | `f5f2922750` | Digests deferred to `Patch::to_durable` and memoized per snapshot. |
| C | `e94bd56f58` | B, plus editor opens read the Document and Current User streams once, and the slide-order commit reuses its publishing editor for the before-payload capture. |

`binaries.json` records the SHA-256 of every measured binary (probe, allocation
probe, sealed 0734 probe, harness) for each arm and the exact build commands.

**Compilers.** The probe and sealed-0734-probe binaries were built with rustc
1.98.1, the host's default stable toolchain. They were built from scratch
directories outside the repository's `rust-toolchain.toml` pin (1.95.0). The
harness binaries and all gates used 1.95.0. Every comparison is between arms
built with the same compiler; absolute times are not directly comparable with
records that used 1.95.0 probes.
Probe builds remap source paths with `--remap-path-prefix`, so the arm binaries
differ only through the source change. A's harness is the shared prebuilt base
binary. B's and C's harnesses were built the same way, with no RUSTFLAGS.

## Directory

- `probe/`: the 0745 probe source (`src/`) and its three manifests, which
  differ only in the arm source tree (`manifests/{before,after-B,after-C}`),
  plus the shared `Cargo.lock`. `src/lib.rs` documents every timed region.
- `scripts/`:
  - `run.py`: the timing matrix.
  - `analyze.py`: statistics.
  - `counters.py`: `perf stat` lane.
  - `alloc.py`: counting-allocator lane.
  - `profile_summary.py`: perf-script attribution restricted to the timed owner.
  - `reduce_p0734.py`: reduces the sealed-probe reports.
  - `summarize.py`: renders `summary.md`.
- `matrix/`: 648 process reports, 12 cases × 3 arms × 18 rounds.
  - `commands.json`: exact commands, exit codes and start times.
  - `analysis.json`: statistics.
  - `reduction-manifest.json`: original SHA-256 and size of every reduced report.
  - Empty stderr files were removed.
- `summary.md`: every timing, counter and allocation table, and every flag
  above +5%.
- `counters/`: per-owner `perf stat` results (`counters.json`) and the raw
  `perf stat -x,` outputs.
- `allocation/`: per-owner counting-allocator regions per case and arm.
- `goldens/`: deterministic durable-patch goldens from all three arms. They are
  byte-identical (SHA-256
  `aae468e1ba9dbf0dc523020d5f0bf002c283e5d78338039c455f6d9dc831a54c`).
- `profiles/`: frame-pointer `perf record` summaries of the timed owner (base
  and B), for `remove`, `remove-durable` and `chain-durable`. Raw `perf.data`
  files were deleted after summarizing.
- `layout/`: the environment-padding scan (A and B). Stack/environment padding
  from 0 to 4,032 bytes in 64-byte steps did not move any result.
- `p0734-probe-manifests/`: the sealed 0734 probe's manifest with its five
  dependency paths pointed at the base tree (`before`) or the worktree
  (`after`, used for B at `f5f2922750` and for C at `e94bd56f58`). The probe
  source and `Cargo.lock` are the sealed 0734 files, unchanged.
- `p0734-controls/`: the sealed 0734 probe's eight negative corruption
  controls. Every control is rejected in every arm, for 45543.ppt and
  FloatingPictures.doc.
- `gates.txt`, `cleanup.json`, `log-sections.md`.

## Timing method

Every process runs pinned with `taskset -c 16` on the 32-core AMD EPYC 9R45
host, using warm caches. Other agents were building on other cores at the same
time. Probe processes run 5 warmups and 50 measured owners. The sealed 0734
probe runs 3 warmups and 50 owners, as in 0734. The harness runs
`ppt_semantic_one_edit_save,ppt_semantic_noop_edit_save` with writer shapes
tiny and large, 5 warmups and 40 samples.

**Heap-layout randomization.** A pilot run of this matrix (two arms, fixed
command lines) showed `chain-durable` at +12% for B against A. The same pair
measured −15% when started from the shell. Scans then isolated the cause:

- Padding the environment by 0–4,032 bytes changed nothing (`layout/`).
- The length of argv[0] changed everything. The probe's own
  `std::env::args()` call (`probe/src/lib.rs:116`) copies it into a heap
  allocation. Rust's startup does not.
- Per-owner user instructions stayed fixed per arm across layouts (±0.01 M).
- Page faults per owner varied by up to 3.5× with layout (`counters/`).

The effect is allocator state, not code. glibc's mmap-threshold and trim
behaviour is the suspected mechanism, but that is inferred, not tested: no run
fixed the `MALLOC_*` settings. The final matrix therefore starts
each binary through a symlink whose path grows by 8 bytes per round:
`argv0/<arm>/<binary>-p` followed by `8*round` `q` characters. Every arm in a
round uses the same argv[0] length, so the comparison stays paired, and 18
rounds cover 18 startup heap layouts. Arm order cycles through all six
permutations, and case order rotates each round.

**Statistics.**

- Process p50 is the midpoint median, and p95 is the nearest rank.
- Each comparison (B vs A, C vs B, C vs A) is the median over the 18 per-round
  paired changes.
- Its interval is a percentile bootstrap of that median: 10,000 resamples,
  seed 745.
- `analysis.json` lists every per-round change above +5% under `flags`.

## Oracles

- Every probe process checks that all its owners published identical bytes.
- `analysis.json` checks that the output digest, and the durable digest where
  present, are identical across all 54 processes of a case. It holds for every
  case.
- Every sealed-0734-probe sample passed its full oracle. Its expected output
  SHA-256 `545ed7e5ba7b8cc0687f8d21ac95aca68fa4d67fd89fd4277ed7cce047bdb859`
  is the sealed 0728 digest.
- The goldens cover 32 scenario and fixture entries: 20 commits publish and
  12 are refused by the fixture. Each records:
  - the commit output;
  - the forward and inverse deterministic durable JSON;
  - durable replay on a fresh base (20 of 20 succeed);
  - durable restore (16 succeed; 4 slide-removal restores on 45543.ppt and
    41246-1.ppt are refused in every arm with "Presentation contains multiple
    OfficeArt BStore containers");
  - the refusal text wherever a stage refuses.

  They are byte-identical across A, B and C. The unit test
  `durable_wire_bytes_match_the_eager_digest_implementation` pins eight of
  those entries to A's values.

## Replay

Run from this directory after rebuilding the binaries as `binaries.json`
describes, into `/home/zhuhe/code/litchi-worktrees/scratch/0745/bin/{A,B,C}`
(`run.py` uses that root):

```sh
python3 scripts/run.py matrix 18
python3 scripts/reduce_p0734.py matrix
python3 scripts/analyze.py matrix > matrix/analysis.json
python3 scripts/counters.py counters-raw > counters/counters.json
python3 scripts/alloc.py allocation/raw > allocation/allocation.json
python3 scripts/summarize.py . > summary.md
```

Timings are specific to this host, its load and these fixtures. Replays should
compare the paired changes, not the absolute microseconds.
