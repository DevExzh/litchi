# Evidence packet: change 0606, PPT record tree borrowed payloads

Record: [`docs/performance/0606-ppt-record-tree-borrowed-payloads.md`](../../0606-ppt-record-tree-borrowed-payloads.md).

## Contents

| path | what it is |
| --- | --- |
| `record-tree-census.txt` | Owned payload bytes, record count and depth for all 30 `.ppt` fixtures on the before leg. The after leg produces a byte-identical file, because the census is taken through the `&[u8]` entry point that this change leaves copying. |
| `allocations-before.txt`, `allocations-after.txt` | Counting-allocator report per operation and fixture: allocation count, allocated bytes, peak live bytes and live bytes still retained while the result is in scope. Produced by `driver/src/alloc.rs`. |
| `callgrind/an-<leg>-<mode>.txt` | Callgrind isolation-pair analysis (s=10 vs s=110, differenced over 100 operations): per-operation instructions, inclusive and self deltas and selected call counts. Eight files: `{before,after}` × `{eager-open, eager-slides, eager-text, owned-edit-save}`. |
| `perf-stat-final.txt` | Native `perf stat` instructions, cycles and minor faults. Per mode and leg: a 1,000-iteration loop under glibc defaults, the same loop with `MALLOC_MMAP_THRESHOLD_`/`MALLOC_TRIM_THRESHOLD_` pinned at 8 MiB, and `-r 200` single-shot processes each doing exactly one operation. |
| `timing/{A1,B1,B2,A2,AA1,AA2}.json` | `litchi-perf-baseline --case ppt_semantic_open,ppt_semantic_list_slides,ppt_semantic_full_text,ppt_semantic_one_edit_save --samples 40 --warmup 5`, in that run order. A = before, B = after; AA1/AA2 are two further before legs taken in the same window for the A/A floor. |
| `timing/fx-<cfg>-<mode>-<leg>.txt` | Raw per-sample nanoseconds for the same four operations on the real fixture `45543.ppt`, 40 samples per leg after 5 warm-ups, same A1 B1 B2 A2 AA1 AA2 order. `<cfg>` is `default` (glibc defaults) or `heap` (`MALLOC_MMAP_THRESHOLD_=MALLOC_TRIM_THRESHOLD_=8388608`). |
| `timing/summary.txt` | p50, mean, p95 and p99 for every leg above, the paired deltas in both directions and the A/A spread. |
| `oracle-digests.txt` | Per-fixture SHA-256 of the differential reader dump, before and after, for all 30 `.ppt` fixtures. |
| `driver/` | The scratch driver, retained in full: `src/main.rs` (`tree`, `profile`, `time` and `oracle` modes) and `src/alloc.rs` (the counting-allocator binary), plus its manifest. `@ROOT@` in `Cargo.toml` is the checkout the leg is built against. |
| `scripts/` | `capture-callgrind.sh`, `capture-final.sh` (callgrind plus the native counts), `capture-timing-final.sh` and `analyze.py` (the isolation-pair analyzer, taken verbatim from `results/change-0587/doc-ppt/`). `$SCRATCH` is the session scratch directory. |
| `gates.txt` | The tail of every gate run before the commit. |
| `decision.json` | Machine-readable decision record. |
| `log-sections.md` | The four log paragraphs for HOTSPOTS, GOAL_AUDIT, REPORT and ADR_COMPLIANCE; the coordinator merges them. |

## How to reproduce

1. `git worktree add -b <branch> <dir> 6c4c1469b` for the before leg and check
   out this branch for the after leg.
2. Copy `driver/` to a scratch directory, replace `@ROOT@` in its `Cargo.toml`
   with the leg's checkout path, and `cargo build --release` with a
   `CARGO_TARGET_DIR` of its own.
3. Build `tools/perf-baseline` in each leg with `cargo build --release --locked
   --bin litchi-perf-baseline`.
4. Run `scripts/capture-final.sh` (deterministic counts), then
   `scripts/capture-timing-final.sh` (timing), with `$SCRATCH` pointing at a
   scratch directory and the CPU pin adjusted.
5. `driver` `oracle <test-data root>` on each leg, then normalize the two
   `HashMap` debug renderings before diffing:
   `sed -E 's/by_slide_id: \{[^}]*\}/by_slide_id: {..}/g; s/by_persist_id: \{[^}]*\}/by_persist_id: {..}/g'`.
   Both normalized dumps hash to
   `9cdddef6e870aa35884738c3f1729c9e12b6ca7bc22b2e551a4dad586c236b7c`.

The raw 9.9 MB oracle dumps and the callgrind `.out` files are **not** retained;
`oracle-digests.txt` and `callgrind/an-*.txt` are the extracts the record cites,
and the driver and scripts regenerate them.

## Provenance

- Base commit (both legs' "before"): `6c4c1469beda47c2d44b502024a8dc680af35bfe`,
  branch `feat/office-format-completeness`.
- Change branch: `perf/0606-ppt-record-tree-borrowed-payloads`.
- Before leg source: the shared read-only checkout
  `/home/zhuhe/code/litchi-worktrees/before-6c4c1469b`; after leg: this branch's
  worktree. Both built `--release --locked` with the pinned toolchain.
- Host: AMD EPYC 9R45, 32 logical CPUs, 123 GiB, Linux 7.0.0-1012-aws,
  rustc 1.95.0, valgrind 3.26.0, glibc system allocator. Every measured process
  pinned to CPU 26 with `taskset`; seven other agents were building and
  measuring on other cores throughout.
- Binary sha256:
  - `litchi-perf-baseline` before `e2e1c59323a7bb3738e20cf279a3c0194025c0ee45d761df16a4b24e679d6103`
  - `litchi-perf-baseline` after `96b3a02d3e5ba9fde1d4c47256263a921bc1d9dd7d76232cd2072a9dbf9ec07b`
  - `ppt0606` before `70166afb0ea979d85053cb16aa8ec302e19b61c384126be0a6457309b0ec5bca`
  - `ppt0606` after `73ee5033c8a7c2f631959c6fbd5696abe19a950843f961a3a0dcfda82b68ff72`
  - `ppt0606_alloc` before `63de5a9bf1fbd3b94118b3fe5a4ee370842f4ef107eab8c5108cd56998761788`
  - `ppt0606_alloc` after `f357b4e401c48e01817615d3cb9e6c691a6c2633467da682de502898d6be4891`
- Primary fixture: `test-data/poi/test-data/slideshow/45543.ppt`, 311,524-byte
  `PowerPoint Document` stream, 286 records, depth 5, copy factor 1.63.
