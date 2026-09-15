# Evidence: change 0617, the length-changing OLE2 save baseline and the frozen copy-through design

Change record:
[`0617-cfb-copy-through-writer-design.md`](../../0617-cfb-copy-through-writer-design.md).

Disposition: **retained, design only.** `performance_claim: none`. **No file
under `crates/` was modified.** The container rebuild measured here is 1.07% of
a length-changing XLS save and 13.25-29.99% of a DOC one in native cycles, so
change 0587's falsification condition for CFB-2 is met on XLS and not on DOC;
the design is frozen and blocked on an ADR clarification of physical sector
layout.

## Contents

| Path | What it is |
| --- | --- |
| `counts.txt` | Exact deterministic counts for one length-changing edit-and-save on each of five cases: source and output size, stream inventory, which streams the public editor changed and by how much, and the unchanged-stream share. Read back through the ordinary public CFB parser from the actual editor output. |
| `counts-alloc.txt` | Allocation regions from the isolated `cfb_save_probe_peak` binary: allocation calls, bytes allocated and deallocated, and peak live bytes for `open`, `commit`, `container` and `container-changed-only` on each case. |
| `attribution.txt` | Callgrind isolation pairs (2 and 12 iterations, per-symbol self instructions differenced and divided by 10), aggregated by owning crate, for all four operations on all five cases. Carries its own caveat: 85% of the container leg is `rep movsb`/`rep stosb`, which callgrind counts once per byte. |
| `attrib/<case>-<operation>.txt` | The per-case, per-operation output behind `attribution.txt`, including the top 30 symbols by self instructions per operation. 20 files. |
| `cycles.txt` | Native `perf stat` isolation pairs (10 and 110 iterations, median of 11 repetitions each, `taskset -c 11`): cycles and instructions per operation, the container's share in both metrics, and the A/A control. **These are the figures to rank on.** |
| `cycles-raw/` | Every individual `perf stat` CSV line behind `cycles.txt`, both levels of every leg, so the medians can be recomputed. 44 files. |
| `determinism.txt` | `OleWriter::write_to` run 12 times in 12 separate processes for each of 1, 2, 3 and 8 explicitly created storages, with an FNV-1a digest of each output. One digest for 1 storage; twelve distinct digests for 8. |
| `ppt-target-sweep.txt` | The 40-fixture × 6-target sweep for a length-changing PPT shape-text edit, with the refusal histogram and the fixture list. Zero admissions. |
| `gates.txt` | Tails of `cargo fmt --all --check`, `cargo clippy -p litchi-cfb -p litchi-ole-common --all-targets --locked`, `cargo test` on both crates and `cargo doc` on both, plus the full `test result` lines. All run on the clean worktree at the base commit. |
| `probe/` | The scratch probe: `Cargo.toml` (path dependencies on `litchi-cfb`, `litchi-core`, `litchi-doc`, `litchi-ppt`, `litchi-xls`, `litchi-ole-common`; `<REPO-ROOT>` stands for the checkout it was built against), `src/lib.rs`, `src/main.rs`, `src/alloc_metrics.rs`, `src/bin/peak.rs`, `src/bin/determinism.rs`. |
| `scripts/attribute.sh` | The callgrind isolation-pair driver that produced `attrib/`. |
| `scripts/cycles.sh` | The `perf stat` isolation-pair driver that produced `cycles-raw/`. |
| `scripts/aggregate.py` | The callgrind parser and per-crate bucketer that turns two profiles into one attribution. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, for the coordinator to merge. |

## Provenance

- Base commit: `8fe9efa55` (`perf(zip,opc): share one Deflate decoder across an
  OOXML open's structural reads`).
- Branch: `perf/0617-cfb-copy-through-writer-design`.
- Working copy: a worktree at `/home/zhuhe/code/litchi-worktrees/0617`; the
  shared repository was not built in or modified.
- Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws. Seven other
  agents were building and measuring on the same machine throughout, which is
  what the +3.29% cycles A/A in `cycles.txt` shows.
- Toolchain: rustc 1.95.0 (59807616e 2026-04-14), cargo 1.95.0 (f2d3ce0bd
  2026-03-21). valgrind 3.26.0, `perf` available.
- Measured CPU: 11, via `taskset -c 11` on every profiled and timed process.
- Probe build: `--release` with `debug = 1`, own `CARGO_TARGET_DIR` outside the
  repository, at `/home/zhuhe/code/litchi-worktrees/targets/0617-probe`
  (removed after the evidence was copied here).
- Binary sha256:
  - `cfb_save_probe` `313d58a8cab2b63416cb6ab60f705682dc21b83ea301141253619f744b491470`
  - `cfb_save_probe_peak` `dc3f512efdcfed200e0ebac8074034e5c68327396b6268f175e078355d0dfa77`
  - `cfb_writer_determinism` `fcf23626ee958c1070186866b275676caedcd367ae0c2ed66c790e7c8a72eeac`
- Probe source sha256 (the six retained files concatenated in the order
  `src/lib.rs`, `src/main.rs`, `src/alloc_metrics.rs`, `src/bin/peak.rs`,
  `src/bin/determinism.rs`, `Cargo.toml`):
  `c3edfc4844e0bc44914ef9aa72cdddc4fbf553a95ff0c046129154a06d3f5814`.

## Reproducing

```sh
# 1. build the probe outside the repository
cd <PROBE-DIR>   # probe/ from this packet, with <REPO-ROOT> substituted
CARGO_TARGET_DIR=<TARGET-DIR> cargo build --release

# 2. the generated PPT fixture (no real .ppt can reach this edit; see ppt-target-sweep.txt)
<TARGET-DIR>/release/cfb_save_probe --format ppt \
  --emit-ppt-fixture <SCRATCH>/authored-3x2.ppt --slides 3 --shapes 2 --input /dev/null

# 3. deterministic counts
<TARGET-DIR>/release/cfb_save_probe --format xls \
  --input <REPO-ROOT>/test-data/poi/test-data/spreadsheet/54016.xls --report

# 4. allocation regions (the isolated binary that installs the counting allocator)
<TARGET-DIR>/release/cfb_save_probe_peak --format doc \
  --input <REPO-ROOT>/test-data/ole/doc/FloatingPictures.doc --report

# 5. instruction attribution and native cycles
PROBE=<TARGET-DIR>/release/cfb_save_probe REPO=<REPO-ROOT> OUT=<SCRATCH> CPU=11 \
  AUTHORED=<SCRATCH>/authored-3x2.ppt bash scripts/attribute.sh
PROBE=<TARGET-DIR>/release/cfb_save_probe REPO=<REPO-ROOT> OUT=<SCRATCH> CPU=11 \
  AUTHORED=<SCRATCH>/authored-3x2.ppt bash scripts/cycles.sh

# 6. the OleWriter determinism finding
for i in $(seq 12); do <TARGET-DIR>/release/cfb_writer_determinism 8; done | sort -u
```

## What is not here

No before/after pair, because nothing was implemented and there is no `after`
leg. No wall-clock timing series: the ranking metric here is native cycles by
isolation pair, and the A/A control is on that metric. No peak-RSS measurement:
the allocation regions are counting-allocator high-water marks, not resident set
size. No corpus-wide census of the unchanged-stream share, which the record names
as the cheapest next measurement.
