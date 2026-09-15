# Evidence: change 0591, the ordinary DOCX edit scans the main part twice

Change record: [`0591-docx-edit-single-scan.md`](../../0591-docx-edit-single-scan.md).

Disposition: retained. `performance_claim: none`. Two of the four whole-part
scans in an ordinary DOCX edit and save are removed. Everything cited by the
record is here; nothing here is registered as a claim.

## Contents

| Path | What it is |
| --- | --- |
| `counts/counts.sh` | The capture script: `litchi-perf-baseline --warmup 0 --samples N --case <case>` under callgrind at N=1 and N=3, for the three DOCX selectors, on each leg, pinned to CPU 11. |
| `counts/extract.py` | The isolation-pair extractor: inclusive Ir per symbol from the annotations, call counts parsed from the raw `callgrind.out` `cfn=`/`calls=` pairs, both differenced between N=3 and N=1 and halved. |
| `counts/incl-<leg>-<case>-<N>.txt` | The top 200 rows of `callgrind_annotate --inclusive=yes --threshold=100` for each run (`head -200`; every symbol the record cites is inside the top 70). |
| `counts/symbols-<leg>-<case>-<N>.txt` | The same annotations filtered to `PROGRAM TOTALS` and the cited symbols, for reading without the full table. |
| `counts/before-summary.json`, `counts/after-summary.json` | Every deterministic count the record's first table states, per case: whole-iteration Ir, and per symbol the inclusive Ir and call count per lifecycle. |
| `probe/` | Source of the timed-region probe (`Cargo.toml` with a `LITCHI_ROOT` placeholder for the leg's checkout, `rust-toolchain.toml` pinning the workspace's 1.95.0, `src/main.rs`). It opens a constant four packages, then runs `edit_document`..`to_stream` — exactly the harness's timed region — on `measured` of them. |
| `region/region-<leg>-<mode>-<paragraphs>-<N>.log` | Valgrind's tail for each probe run: modes `noop`, `one`, `one-percent` × 24, 200 and 10,000 paragraphs × N ∈ {1, 3} × two legs, 36 runs. |
| `region/region-summary.json` | The per-shape timed-region instruction counts of the record's second table, `(Ir(N=3) − Ir(N=1)) / 2`. |
| `timing/timing.sh` | The paired-timing script: order A1 B1 B2 A2, `--warmup 5 --samples 50` per run, all three selectors per run, pinned to CPU 11. |
| `timing/A1.json`, `B1.json`, `B2.json`, `A2.json` | The harness reports as written, with every per-sample `elapsed_ns` value, the corpus manifests and the binary identity of the leg that produced them. |
| `timing/analyze.py`, `timing/paired-summary.json` | Pools the two runs of each leg (100 samples per leg) and produces the record's third table: p50, mean, p95, p99 per leg, the delta in both directions, and the A/A and B/B floors observed in the same window. |
| `gates.txt` | The tail of every gate run in the change worktree. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`; this batch did not edit those files. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `cleanup.json` | What was removed from the session scratchpad and from disk after this packet was assembled. |

## Provenance

- Base: `08d968f8ec7db27cf1187d01911fd08b9d014d91` (branch
  `feat/office-format-completeness`, carrying change 0587).
- Change branch: `perf/0591-docx-edit-single-scan`, worktree
  `/home/zhuhe/code/litchi-worktrees/0591`.
- Before leg: the shared read-only checkout
  `/home/zhuhe/code/litchi-worktrees/before-08d968f8e` at the base commit, built
  into its own `CARGO_TARGET_DIR`.
- Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws; rustc 1.95.0
  (the workspace pin, used for both harness legs and both probe legs);
  valgrind 3.26.0. Seven other agents were building and measuring on the same
  host throughout; the A/A and B/B floors in `timing/paired-summary.json` are
  what that contention amounted to in this window.
- Builds: `cargo build --release --locked` on both legs, identical flags.

| binary | sha256 |
| --- | --- |
| `litchi-perf-baseline` (before) | `dab64c26373072ab9d88883c64cce710df29aa8a36255d751dbcf3d72f2521a1` |
| `litchi-perf-baseline` (after) | `e177bfcf89a86e15270cadd4339c4898e6c729b7105f470dad5e548b1e55ed75` |
| `docx-edit-region` probe (before) | `559d3a4240de38f15bfd6e4a2640823fb6acd60d9c8158e9739c5ab34d860d0b` |
| `docx-edit-region` probe (after) | `8ad99f718b370bc8308b854d44ad97bb9e1b3705a864a1ae0780ab6f6239debd` |

## Replaying it

```sh
# both legs, deterministic counts
counts/counts.sh <leg-binary> <before|after> <outdir>
for f in <raw>/<leg>-*.out; do
  callgrind_annotate --inclusive=yes --threshold=100 "$f" > "incl-$(basename "$f" .out).txt"
done
python3 counts/extract.py <outdir> <raw-callgrind-dir> <before|after>

# timed region alone: substitute the leg's checkout for LITCHI_ROOT, build, then
# taskset -c 11 valgrind --tool=callgrind <probe> <mode> <paragraphs> <1|3>

# paired timing
timing/timing.sh <outdir> && python3 timing/analyze.py <outdir>
```

## What this packet does not establish

Nine scenarios on one synthetic corpus on one host with two builds. No cold
cache, no physical device, no real-producer document, no peak RSS, no allocation
profile, no concurrency scaling, no cross-platform result, and no claim about
DOCX operations other than a direct-body paragraph text rewrite and the patch
application that publishes it. Instruction counts rank work and are not latency;
the paired medians are reported beside the floor measured in the same window and
are not registered as a speedup. The raw `callgrind.out` files were deleted after
the annotations and call counts were extracted; `cleanup.json` records them.
