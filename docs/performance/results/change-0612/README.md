# Evidence: change 0612, the XLS globals skip measured and the design frozen

Change record:
[`0612-xls-skip-uninterpreted-globals-design.md`](../../0612-xls-skip-uninterpreted-globals-design.md).

Disposition: **design only.** `performance_claim: none`. **No file under
`crates/` was modified.** The density gate does fire on real fixtures — 19 of
123 modellable `.xls`/`.xlt` fixtures take at least one seek, and the flagship's
open drops from 565,201 to 82,867 read bytes, −85.34% — but the one fixture it
reaches gets **5.54% slower in native cycles on an in-memory source and 17.22%
slower on a file source**, +5.91%/+6.45% and +16.93%/+16.35% at p50 in paired
wall clock, against an A/A floor under 1% in the same window. The two fixtures
the gate never reaches are identical read-for-read and byte-for-byte and are
0.5% to 1.4% slower from the per-record gate test alone. XLS-6 is falsified and
the design is frozen unimplemented.

## Contents

| Path | What it is |
| --- | --- |
| `analysis.txt` | The exact output of `scripts/analyze.py` over `callgrind/`, `perf/` and `perf-file/`, then `model/fold_model.py`, then `scripts/fold_latency.py` over `latency-summary.json`. Every instruction, cycle, byte, read and latency figure the record cites appears here. |
| `model/globals_skip_model.py` | The schedule model. Pure standard library plus `olefile`; walks the CFB container and the BIFF8 globals framing with arithmetic and shares no code with `litchi-xls`, so it is an independent oracle for the counters rather than a restatement of them. It replays today's fill schedule byte-for-byte (the four-record exact prologue, the coupled next-header prefetch, 512-byte windows doubling to 64 KiB, the stream, `max_global_bytes` and running-minimum `BoundSheet8` clamps) and five candidate schedules. |
| `model/globals-skip.json` | Its output: one object per fixture with the record-kind composition, the touched closure, the skipped-payload size distribution, today's fills and bytes, and per gate the fills, bytes, seeks, retained bytes and deltas. |
| `model/fold_model.py` | Reprints the six model tables from `globals-skip.json` alone. |
| `counters/counters.txt` | The deterministic control: 36 cells (2 legs × 3 fixtures × 2 source modes × `open`/`list`/`one-cell`), each asserted identical across the five samples of its own child, with the harness's independent eager-parser projection of the worksheet count. |
| `callgrind/ann-<leg>-<fixture>-<op>-s{small,large}.txt` | `callgrind_annotate --threshold=99.9` **self** cost. |
| `callgrind/inc-<leg>-<fixture>-<op>-s{small,large}.txt` | `--inclusive=yes` from the same raw profile, which is what gives `parse_globals`' and the CFB reader's share of the open. Differencing the large- and small-sample child and dividing by the extra operations isolates one operation: change 0574's method as 0576, 0584, 0595 and 0608 used it. |
| `perf/perf-<leg>-<fixture>-open-s{100,1100}-r<1..5>.csv` | `perf stat -x,` cycles, instructions, branches, branch misses and task-clock on an **owned in-memory** source, isolated the same way, five repetitions per cell so the median is taken before the pair is differenced. Legs `aa` and `bb` are byte-for-byte copies of `base` and are the floor. |
| `perf-file/…` | The same on a **file** source, which is the mode that pays a `pread64` and a `statx` per read and is where the trade is worst. |
| `latency-summary.json` | The folded `A1 B1 B2 A2` wall-clock capture: 3 B legs × 3 fixtures × 2 modes × 4 rounds, each with n, p50, mean, p95, p99 and the SHA-256 of the binary that produced it. 400 samples per round after 50 warm-ups. The 54 raw round captures behind it are 18 MB of per-sample JSON and are **not** retained; `scripts/fold_latency.py --summary` reprints every latency table from this file. |
| `latency/quiescence-*.log` | Host load at both ends of each timing window. |
| `scripts/scaffold-skip.patch` | **Measurement scaffold, not a candidate.** The gate at 1 KiB applied to *which stream bytes are fetched*. The buffer keeps today's shape — one contiguous `[0, global_end)` allocation, zero-filled in full — so allocation, zero-fill, both framings and every semantic output are byte-identical to the base and the skipped bytes are zeros nothing reads. Applied, built, measured, reverted. |
| `scripts/scaffold-cursor.patch` | **Measurement scaffold.** Change 0574 opportunity 3 alone: one retained `SharedOleStreamCursor` in place of one `read_stream_range_hinted` per fill, no skip. |
| `scripts/scaffold-skip-cursor.patch` | Both together — the design in its best form. |
| `scripts/capture_counters.sh` | The deterministic-counter driver. |
| `scripts/capture_callgrind.sh` | The callgrind isolation-pair driver; takes both a self and an inclusive annotation from each profile. |
| `scripts/capture_perf.sh` | The native hardware-counter driver, with repetitions and a source-mode switch. |
| `scripts/capture_latency.sh` | The `A1 B1 B2 A2` wall-clock driver; times both source modes for every cell. |
| `scripts/analyze.py`, `scripts/fold_latency.py` | Recompute every table from this directory alone. Pure standard library. |
| `gates.txt` | The tails of `cargo fmt --all --check`, `cargo clippy -p litchi-xls --all-targets`, `cargo doc -p litchi-xls --no-deps` and `cargo test -p litchi-xls` (72 binaries, 1,390 passed, 0 failed, 1 ignored), all on the clean worktree at the base commit. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, for the coordinator to merge. |

## Provenance

- Base commit: `1e4198321` (`perf(ppt): span the retained stream instead of
  copying every record payload (0606)`).
- Branch: `perf/0612-xls-skip-uninterpreted-globals-design`.
- Working copy: a detached worktree at `/home/zhuhe/code/litchi-worktrees/0612`
  with its own `CARGO_TARGET_DIR`; the shared repository was neither built in nor
  modified. Nothing under `crates/` changed: `git diff 1e4198321 -- crates/` is
  empty, verified before the gates and before the commit.
- Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws. Seven other
  measurement agents were active on the same machine throughout; the one-minute
  load average ran from 29.1 down to 7.5, which is what both floors are for.
- Toolchain: rustc 1.95.0 (59807616e 2026-04-14), cargo 1.95.0, valgrind 3.26.0,
  perf 7.0.14.
- Measured CPU: **8**, `taskset` pinned with `setarch x86_64 -R` for every
  counted and timed run. `RAYON_NUM_THREADS=1`, `OMP_NUM_THREADS=1`.
- Legs, all `tools/perf-baseline`'s `xls_source_attribution` built
  `--release --locked --features xls-source-attribution` from the same worktree:

  | leg | sha256 | what it is |
  | --- | --- | --- |
  | `base` | `f60f3a02ffa016fb7b05fea607fb8806dbcfbba9c211097ff72f942d868b03f0` | the unmodified base |
  | `skip` | `585edbc3c7879f725503c222ea87cc679e0d9f048d223e828a1678c80863f193` | `scaffold-skip.patch` |
  | `cursor` | `c7ebdb38e382e43df6bf577d77b241a50a993d58eabc7e7ca140197838a27a40` | `scaffold-cursor.patch` |
  | `skipcur` | `c1c5995e2b8a10394e0b50290ab6556d02d317ec8732980b99279ede743a9883` | `scaffold-skip-cursor.patch` |
  | `aa`, `bb` | `f60f3a02…` | byte-for-byte copies of `base`, the floor |

- Fixtures:

  | stem | path | sha256 |
  | --- | --- | --- |
  | flagship | `test-data/ole/xls/ConditionalFormattingSamples.xls` | `d1942d857ffbd4d10ebca1745cd5d70c14af9d9f1388c91ed0a0800e31ad5ce7` |
  | cv | `test-data/ole/xls/WithCustomViews.xls` | `3c0c168f38498cc7a356ffaee82b19241ab022ca8813c4ebef286399c32cbd64` |
  | 54016 | `test-data/poi/test-data/spreadsheet/54016.xls` | `2e050f1fbb31868b097aa6d4d0fe0a16af8e39c252af82d01cd8f93c4f9a911a` |

  The model covers all 126 `.xls` and `.xlt` fixtures under `test-data`; three
  are encrypted and excluded.

## Harness scope, stated once

`xls_source_attribution` counts **logical** calls through a wrapper over a warm,
page-cached, immutable staged copy: no physical-I/O counter, no cold cache, no
remote source. `owned-readat` wraps bytes in memory and performs zero syscalls;
`file-source` wraps a staged immutable file and performs one `pread64` and one
`statx` per counted call. Timings from the two modes are never compared with
each other. The harness times **every read call it makes**, so a leg that issues
48 more reads pays 96 more `clock_gettime` pairs that production would not; that
instrumentation is named and sized in the record rather than subtracted.

## Replay

```sh
cd docs/performance/results/change-0612
python3 model/globals_skip_model.py --repo /home/zhuhe/code/litchi --out model/globals-skip.json
python3 model/fold_model.py
python3 scripts/analyze.py callgrind perf perf-file
python3 scripts/fold_latency.py --summary latency-summary.json
```

To rebuild a leg: apply the matching patch in `scripts/` to a checkout of the
base commit, build `xls_source_attribution` with
`--release --locked --features xls-source-attribution`, and revert. Each scaffold
applies to `crates/litchi-xls/src/workbook/source.rs` alone.
