# Evidence: change 0604, the OLE2 whole-stream zero-fill ceiling and the frozen appending design

Change record:
[`0604-cfb-append-reads-design.md`](../../0604-cfb-append-reads-design.md).

Disposition: **retained, design only.** `performance_claim: none`. **No file
under `crates/` was modified.** The ceiling measured here is 0.57% to 2.94% of
an open in native cycles on every fixture, inside the 4% p50 floor that change
0587 named as this item's falsification criterion, so the design is frozen and
not implemented.

## Contents

| Path | What it is |
| --- | --- |
| `zero-fill-byte-counts.txt` | Exact bytes zero-filled per open at every `u8` zero-fill site in `litchi-cfb` plus the `litchi-xls` globals buffer, for five opens on owned in-memory sources. One open per line group; produced with the instrumentation patch below, which was reverted before the gates and the commit. |
| `ceiling.txt` | `perf stat` cycles and instructions per operation for the five opens and for the `memset`, `zero+copy`, `append` and allocation-only micro-benchmarks at each fixture's byte counts, plus the isolation-pair A/A legs. Isolation pairs at 10 and 110 samples, median of 11 repetitions, `taskset -c 19`. |
| `ceiling-raw.txt` | Every individual `perf stat` cycle total behind `ceiling.txt`, both legs of every pair, so the medians can be recomputed. |
| `aa-floor.txt` | Wall-clock A/A floor: the same binary as both legs, 60 samples per leg in `A1 B1 B2 A2` order, each sample a 30-iteration mean after 3 warm-ups, pinned; p50, mean, p95, p99 per leg and the paired delta in both directions. |
| `gates.txt` | Tails of `cargo fmt --all --check`, `cargo clippy -p litchi-cfb -p litchi-core -p litchi-xls --all-targets --locked`, `cargo test -p litchi-cfb --release --locked`, the three named truncated-final-sector parity tests, and `cargo doc` on the three crates. All run on the clean worktree at the base commit. |
| `probe/` | The scratch probe: `Cargo.toml` (path dependencies on `litchi-core`, `litchi-cfb`, `litchi-doc`, `litchi-ppt`, `litchi-xls`; `<REPO-ROOT>` stands for the checkout it was built against) and `src/main.rs`. Eight cases: four opens on owned in-memory sources and four micro-benchmarks. |
| `scripts/ceiling.sh` | The isolation-pair driver that produced `ceiling.txt` and `ceiling-raw.txt`. |
| `scripts/aa-floor.sh` | The paired A/A driver that produced `aa-floor.txt`. |
| `scripts/zero-fill-count-instrumentation.patch` | The temporary `eprintln!` instrumentation at the six `litchi-cfb` zero-fill sites and the `litchi-xls` globals `resize`, used only to produce `zero-fill-byte-counts.txt`. Applied, captured, reverted; the working tree was verified clean before the gates were run. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, for the coordinator to merge. |

## Provenance

- Base commit: `f8cf7d2a1d155615ab97eda6310c3d5ced0b82ae`
  (`perf(opc): prove unchanged relationships at open, and buffer the atomic tempfile`).
- Branch: `perf/0604-cfb-append-reads-design`.
- Working copy: a detached worktree at `/home/zhuhe/code/litchi-worktrees/0604`;
  the shared repository was not built in or modified.
- Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws. Seven other
  agents were building and measuring on the same machine throughout, which is
  what the `aa-floor.txt` tails show.
- Toolchain: rustc 1.95.0 (59807616e 2026-04-14), cargo 1.95.0.
- Probe build: `--release` with `debug = 1`, own `CARGO_TARGET_DIR` outside the
  repository. Binary sha256
  `de22ca1072e97c9f69c878b9e08f871f37d093931e84b2d931d5c1dbccf82fae`
  (the clean build used for `ceiling.txt` and `aa-floor.txt`; the instrumented
  build used for `zero-fill-byte-counts.txt` is a separate binary and was
  deleted with its target directory).
- Measured CPU: 19, `taskset` pinned for every timed and counted run.
- Fixtures, all from the repository's own corpus:
  `test-data/poi/test-data/document/ca.kwsymphony.www_education_School_Concert_Seat_Booking_Form_2011-12.doc`,
  `test-data/ole/doc/FloatingPictures.doc`,
  `test-data/poi/test-data/slideshow/45543.ppt`,
  `test-data/ole/xls/ConditionalFormattingSamples.xls`.

## Reproducing

```sh
# 1. the probe, built against the checkout under measurement
#    (edit probe/Cargo.toml's <REPO-ROOT> first)
CARGO_TARGET_DIR=<somewhere outside the repo> \
  cargo build --release --manifest-path probe/Cargo.toml

# 2. the ceiling and the isolation-pair A/A
BIN=<target>/release/cfb_zero_fill_ceiling REPO=<repo> OUT=ceiling.txt \
  CPU=19 ./scripts/ceiling.sh

# 3. the wall-clock A/A floor
BIN=<target>/release/cfb_zero_fill_ceiling REPO=<repo> OUT=aa-floor.txt \
  CPU=19 N=30 ./scripts/aa-floor.sh

# 4. the byte counts (apply, build, run, revert)
git apply scripts/zero-fill-count-instrumentation.patch
#   … rebuild the probe, run each --case with --warmups 0 --samples 1,
#   aggregate the "ZF <site> <bytes>" lines …
git checkout crates/litchi-cfb/src/file.rs crates/litchi-cfb/src/shared.rs \
             crates/litchi-xls/src/workbook/source.rs
```

## What this packet does not establish

- **No speedup and no regression.** Nothing was implemented, so nothing was
  compared before against after. The ceiling is the cost of the zero-fill
  measured in isolation, an upper bound on what removing it could return.
- **Owned in-memory sources only.** Every open measured wraps bytes already in
  memory. The design is a no-op on `FileSource` and on `File`-backed readers by
  construction, and no file-backed leg was taken.
- **Five fixtures, one host, one build, one CPU.** The corpus maximum is 1.6 MB
  with no DIFAT sector and no 4,096-byte sector; no synthetic large fixture was
  built.
- The `slurp-*` micro-benchmarks run in a tight loop, so the allocator reuses a
  warm heap block after the first iterations. That affects both of their legs
  equally, but their absolute cycle figures are warm-heap numbers.
- Nothing here says anything about CFB-2 to CFB-6, XLS-6, or any other item
  change 0587 ranks near this one.
