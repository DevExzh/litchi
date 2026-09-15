# Change 0622 evidence packet

Compact per-cell source facts carried from planning into the XLSX value-only
commit, replacing the commit's second whole-sheet layout scan. The record is
[`../../0622-xlsx-compact-source-facts.md`](../../0622-xlsx-compact-source-facts.md).

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## Provenance

| | |
| --- | --- |
| base commit | `1e41983213dc378c13774ed7038c51faf231977f` |
| branch | `perf/0622-xlsx-compact-source-facts` |
| before checkout | `/home/zhuhe/code/litchi-worktrees/before-1e4198321` (shared, read-only) |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0 (59807616e 2026-04-14), valgrind/callgrind 3.26.0 |
| CPU pin | `taskset -c 15`; eight agents were building on the host throughout |
| build | `cargo build --release --locked --bins --features allocator-metrics` in `tools/perf-baseline`, per leg, each with its own `CARGO_TARGET_DIR` |

Binary sha256:

| binary | sha256 |
| --- | --- |
| before `litchi-perf-baseline` | `aeeb1cd91f90f43bed8eed2787c1c221282ea288710c407df1cc22d8bf08ac48` |
| after `litchi-perf-baseline` | `3f57a6cd74613bb8cda6101fd150eb1692fb459d907e59e1f3d89f22d6f6b375` |
| before `litchi-perf-baseline-alloc` | `261950334e7cc54f7c155091d27ef47dab6dab86d39df80816691ceddd14fb60` |
| after `litchi-perf-baseline-alloc` | `f25f306b4bc43835d120105b1258c59b78a8dd7dfaf1baded412cbbbb6e45b59` |

## Contents

| path | what it is |
| --- | --- |
| `decision.json` | the disposition, its reason codes, accepted costs and known gaps |
| `gates.txt` | the tail of every gate, plus the pre-existing `litchi-iwa` workspace failure and how it was reproduced on the untouched before checkout |
| `log-sections.md` | the four log paragraphs for the coordinator to merge |
| `counts/instruction-summary.txt` | callgrind isolation pairs: whole-sample totals and the per-symbol per-operation attribution |
| `counts/progress.txt` | the twenty isolation-pair runs and their exit codes |
| `alloc/allocation-summary.txt` | change-0538 phase allocation regions, both legs, every case and shape |
| `alloc/alloc-*.json` | the raw allocator-harness reports the summary is derived from |
| `alloc/rss-repeat.txt` | fifteen `/usr/bin/time -f %M` runs per leg on `dense-sparse` one-edit |
| `timing/*.json` | the A1 B1 B2 A2 runs and the four before-only floor runs, trimmed to the fields the summary reads |
| `timing/abba-summary.txt` | paired p50/mean/p95/p99 deltas in both directions beside the A/A floor |
| `differential/output-hashes.txt` | published-package sha256 for 32 (case, shape) pairs, before and after |
| `differential/oracle-bounded.log` | the in-crate differential oracle run, with the real-corpus funnel line |
| `differential/oracle-full.log` | the same oracle with `LITCHI_0622_FULL_ORACLE=1` (every cell of every worksheet) |
| `differential/corpus-marker-census.txt` | a lexical census of the repository's 389 real worksheet parts |
| `scripts/counts.sh` | the callgrind isolation-pair driver |
| `scripts/timing.sh` | the A1 B1 B2 A2 plus floor driver |
| `scripts/alloc-report.py` | the allocation-region differ |
| `scripts/symbols.py` | the per-symbol callgrind differ |
| `scripts/timing-summarize.py` | the paired-delta and A/A-floor summarizer |

## How to reproduce

```sh
# deterministic counts first
scripts/counts.sh                    # writes counts/*.out, then:
python3 scripts/symbols.py           # per-symbol per-operation attribution

# allocations
litchi-perf-baseline-alloc --case <case> --xlsx-cell-crud-shape <shape> \
  --warmup 0 --samples 5 --json alloc-<leg>-<case>-<shape>.json
python3 scripts/alloc-report.py

# the differential oracle
cargo test -p litchi-xlsx --lib change_0622 -- --nocapture
LITCHI_0622_FULL_ORACLE=1 cargo test --release -p litchi-xlsx --lib change_0622 -- --nocapture

# paired timing last, in one window
scripts/timing.sh timing/
```

## What this packet does not contain

No registered claim, no cold-cache or physical-I/O measurement, no range-source
or concurrency result, and no measurement of a real producer file: change 0602
established that the source-backed value editor admits none of the 95 real
`.xlsx` fixtures, and the oracle's funnel line confirms that exactly one of the
repository's 391 real worksheet parts is accepted by the value-only planning
validator at all.
