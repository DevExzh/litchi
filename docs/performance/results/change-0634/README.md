# Evidence packet: change 0634, PPT per-slide re-parse borrows the retained stream

Record: [`docs/performance/0634-ppt-slide-factory-borrowed-reparse.md`](../../0634-ppt-slide-factory-borrowed-reparse.md).

## Contents

| path | what it is |
| --- | --- |
| `allocations-before.txt`, `allocations-after.txt` | Counting-allocator report per operation and fixture: allocation count, allocated bytes, peak live bytes and live bytes still retained while the result is in scope. Produced by `driver/src/alloc.rs`. Six modes: `eager-open`, `eager-slides`, `eager-text`, `eager-notes`, `owned-edit-save`, `source-open`. |
| `callgrind/an-<leg>-<mode>.txt` | Callgrind isolation-pair analysis (s=10 vs s=110, differenced over 100 operations): per-operation instructions, inclusive and self deltas and selected call counts. Ten files: `{before,after}` x `{eager-open, eager-slides, eager-text, eager-notes, owned-edit-save}`. |
| `perf-stat.txt` | Native `perf stat` instructions, cycles and minor faults, both legs. Per mode: a 1,000-iteration loop under glibc defaults, the same loop with `MALLOC_MMAP_THRESHOLD_`/`MALLOC_TRIM_THRESHOLD_` pinned at 8 MiB, and `-r 200` single-shot processes each doing exactly one operation. |
| `timing/{A1,B1,B2,A2,AA1,AA2}.json` | `litchi-perf-baseline --case ppt_semantic_open,ppt_semantic_list_slides,ppt_semantic_full_text,ppt_semantic_one_edit_save --samples 40 --warmup 5`, in that run order. A = before, B = after; AA1/AA2 are two further before legs taken in the same window for the A/A floor. Each file carries both registered PPT corpora, `ppt-tiny` and `ppt-large`. |
| `timing/fx-<cfg>-<mode>-<leg>.txt` | Raw per-sample nanoseconds for the five whole operations on the real fixtures, 40 samples per leg after 5 warm-ups, same A1 B1 B2 A2 AA1 AA2 order. `<cfg>` is `default` (glibc defaults) or `heap` (`MALLOC_MMAP_THRESHOLD_=MALLOC_TRIM_THRESHOLD_=8388608`). |
| `timing/summary.txt` | p50, mean, p95 and p99 for every leg above, the paired deltas in both directions and the A/A spread. Produced by `scripts/summarize-timing.py`. |
| `oracle-digests.txt` | Per-fixture SHA-256 of the normalized differential reader dump, before and after, for all 30 `.ppt` fixtures: 30 of 30 identical. |
| `driver/` | The scratch driver, retained in full. It is change 0606's driver (`src/main.rs` with `tree`, `profile`, `time` and `oracle` modes, `src/alloc.rs` for the counting allocator) with one addition on both legs: an `eager-notes` mode that opens, lists slides and reads every slide's speaker-notes text. `@ROOT@` in `Cargo.toml` is the checkout the leg is built against. |
| `scripts/` | `capture-counts.sh` (callgrind isolation pairs, native `perf stat`, allocation counts for one leg), `capture-timing.sh` (the paired timing), `summarize-timing.py` and `analyze.py` (the isolation-pair analyzer, taken verbatim from `results/change-0587/doc-ppt/` by way of `results/change-0606/`). `$SCRATCH` is the session scratch directory. |
| `gates.txt` | The tail of every gate run before the commit (eleven gates plus the `tools/native-resave` check). |
| `decision.json` | Machine-readable decision record. |
| `log-sections.md` | The four log paragraphs for HOTSPOTS, GOAL_AUDIT, REPORT and ADR_COMPLIANCE; the coordinator merges them. |

## How to reproduce

1. `git worktree add <dir> c7326f680` for the before leg and check out this
   branch for the after leg.
2. Copy `driver/` to a scratch directory once per leg, replace `@ROOT@` in its
   `Cargo.toml` with that leg's checkout path, and `cargo build --release` with a
   `CARGO_TARGET_DIR` of its own. Stage both binaries as
   `$SCRATCH/bin/ppt0634-{before,after}` and
   `$SCRATCH/bin/ppt0634_alloc-{before,after}`, outside any Cargo target
   directory (0627: a concurrent build relinked a timed binary mid-run).
3. Build `tools/perf-baseline` in each leg with
   `cargo build --release --locked --bin litchi-perf-baseline` and stage them as
   `$SCRATCH/bin/litchi-perf-baseline-{before,after}`.
4. `SCRATCH=<dir> scripts/capture-counts.sh before 10` and
   `SCRATCH=<dir> scripts/capture-counts.sh after 10` (deterministic counts),
   then `SCRATCH=<dir> scripts/capture-timing.sh 10` (timing), with the CPU pin
   adjusted. `python3 scripts/analyze.py <leg>-<mode>` reads the callgrind pair;
   `SCRATCH=<dir> python3 scripts/summarize-timing.py` prints `timing/summary.txt`.
5. `driver` `oracle <test-data root>` on each leg, then normalize the two
   `HashMap` debug renderings before diffing:
   `sed -E 's/by_slide_id: \{[^}]*\}/by_slide_id: {..}/g; s/by_persist_id: \{[^}]*\}/by_persist_id: {..}/g'`.
   Both normalized 9,927,532-byte dumps hash to
   `9cdddef6e870aa35884738c3f1729c9e12b6ca7bc22b2e551a4dad586c236b7c`, which is
   the digest change 0606 recorded for the same dump.

The raw 9.9 MB oracle dumps and the callgrind `.out` files are **not** retained;
`oracle-digests.txt` and `callgrind/an-*.txt` are the extracts the record cites,
and the driver and scripts regenerate them.

## Provenance

- Base commit (both legs' "before"): `c7326f68065edf6f2198ca3cb39c38c48cf00ed9`,
  branch `feat/office-format-completeness`. It contains change 0606, so the
  before leg is 0606's after leg.
- Change branch: `perf/0634-ppt-slide-factory-borrowed-reparse`.
- Before leg source: the shared read-only checkout
  `/home/zhuhe/code/litchi-worktrees/before-c7326f680`; after leg: this branch's
  worktree. Both built `--release --locked` with the pinned toolchain.
- Host: AMD EPYC 9R45, 32 logical CPUs, 123 GiB, Linux 7.0.0-1012-aws,
  rustc 1.95.0, valgrind 3.26.0, glibc system allocator. Every measured process
  pinned to CPU 10 with `taskset`; seven other agents were building and
  measuring on other cores throughout.
- Binary sha256:
  - `litchi-perf-baseline` before `2b96165d6a3cb486b5db99d9029b17e305dcea7ff0256c8026ae9d5da3a02b00`
  - `litchi-perf-baseline` after `474a3bf20fc031c1bf9a879043ef20f616d2508f0c8d5ad83d9b5a58d49fb263`
  - `ppt0634` before `6c4ee49f69d5205080c10acdaafb8da1dc9ab72a87632fcfc5768327c6685bc9`
  - `ppt0634` after `df1eeba6f382f95ced4074f65b53e606414c4e94c25aa00fd30e9791f13d2e68`
  - `ppt0634_alloc` before `b1a818079cf3fa26d78fb68b830a43e21f6ed2d87f461e34c2700b736fffb0d8`
  - `ppt0634_alloc` after `227f08eaa4b1e203dd0e698045ca516f010993e1a3a222abd24ee901ceac7e08`
- `cargo fmt` ran after the measurements and reflowed five test and source
  blocks. The `ppt0634` after binary rebuilt from the formatted source is
  byte-identical (`df1eeba6f382f95ced4074f65b53e606414c4e94c25aa00fd30e9791f13d2e68`); `litchi-perf-baseline` rebuilds to
  `927b1110c4c8c647de6ad4ce44c9b78536d60e0cccabd641c0230ec51e07af6a`, differing from the measured binary
  only outside `.text` and `.rodata`, both of which hash identically before and
  after the reformat (`.text` 39,248,742 bytes, `.rodata` 4,907,868 bytes). The
  difference is the panic-location strings rustfmt's line moves rewrote; no
  executed instruction changed, so the timings stand.
- Primary fixture: `test-data/poi/test-data/slideshow/45543.ppt`, 311,524-byte
  `PowerPoint Document` stream, 286 records, 11 slides, no notes pages. Notes
  fixture: `test-data/poi/test-data/slideshow/headers_footers_2007.ppt`,
  99,635-byte stream.
