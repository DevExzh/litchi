# Change 0635 evidence packet

Two retained reductions — the XLSX fact builder's per-element and per-cell work,
and the stylesheet parse's per-event heap copies — plus one designed part that was
implemented, measured and **rejected**: carrying the facts across a snapshot chain.
The record is
[`../../0635-xlsx-facts-builder-and-chains.md`](../../0635-xlsx-facts-builder-and-chains.md).

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## Provenance

| | |
| --- | --- |
| base commit | `c7326f68065edf6f2198ca3cb39c38c48cf00ed9` |
| branch | `perf/0635-xlsx-facts-builder-and-chains` |
| before checkout | `/home/zhuhe/code/litchi-worktrees/before-c7326f680` (shared, read-only) |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0 (59807616e 2026-04-14), valgrind/callgrind 3.26.0 |
| CPU pin | `taskset -c 11`; eight agents were building on the host throughout |
| build | `cargo build --release --locked --bins --features allocator-metrics` in `tools/perf-baseline`, per leg, each with its own `CARGO_TARGET_DIR`; every timed binary staged outside a Cargo target directory (change 0627) |

Binary sha256:

| binary | sha256 |
| --- | --- |
| before `litchi-perf-baseline` | `066154fa6e31e2b5dded9aeca35944ae2278f4edd17d07433b1b6fb1f0744d9f` |
| after `litchi-perf-baseline` | `b2b81e86d6cb26b6b22d6e60c51dcf0758a88b906cd69d53425fd3214610eb99` |
| after-with-chain `litchi-perf-baseline` | `3fadc0f895b9e8a1db7676016be32a4821a785e3ec1cd6adf0e660557567a1df` |
| before `litchi-perf-baseline-alloc` | `5c3a36ca1e27a04214a89a1a9936d6a035335cad0c14f2017221d91fcc8f982b` |
| after `litchi-perf-baseline-alloc` | `5bd3e861192296ff27393c6ec4e5595bf2f2f84e5f1e407fef943e5f2ddf68e7` |

**These hashes identify the exact artifacts measured; they are not a reproducibility
claim.** The workspace release profile sets `lto = true`, and on this toolchain a
rebuild of byte-identical source does not reproduce the binary: rebuilding the
committed tree twice more produced
`08ba9d1e8f2ca5f687d54541a57986fa820a4031d3cd76521e4b473e07baa061` and
`6318afe9d6efe637567ff99b0e930bedefde23e72eb66e9489a7a3b6112bbb93`. Every number in
this packet was taken with the staged binaries named above.

"after-with-chain" is the rejected third part: the retained change plus
`patch/0635-chain-facts.patch`. It exists only so the chain part's cost could be
priced against the change that ships.

## Contents

| path | what it is |
| --- | --- |
| `decision.json` | the disposition, its reason codes, accepted costs and known gaps |
| `gates.txt` | the tail of every gate that was run |
| `log-sections.md` | the four log paragraphs for the coordinator to merge |
| `counts/instruction-summary.txt` | callgrind isolation pairs: whole-iteration totals and per-symbol per-operation attribution, all three legs |
| `counts/progress.txt` | every isolation-pair run and its exit code |
| `alloc/allocation-summary.txt` | change-0538 phase allocation regions, both legs, two cases over three shapes |
| `alloc/alloc-*.json` | the raw allocator-harness reports the summary is derived from |
| `timing/*.json` | the A1 B1 B2 A2 runs and the four before-only floor runs |
| `timing/abba-summary.txt` | paired p50/mean/p95/p99 deltas in both directions beside the A/A floor |
| `differential/output-hashes.txt` | published-package sha256 per (case, shape), before and after |
| `differential/oracle.log` | change 0622's differential oracle, unchanged, run on the after tree |
| `differential/parser-equivalence.log` | the two change-0635 tests comparing the fused reference parsers against the shared ones |
| `patch/0635-chain-facts.patch` | the rejected chain part, appliable with `git apply` on the tip of this branch |
| `scripts/counts.sh` | the callgrind isolation-pair driver |
| `scripts/symbols.py` | the per-symbol and whole-iteration differ |
| `scripts/timing.sh` | the A1 B1 B2 A2 plus floor driver |
| `scripts/timing-summarize.py` | the paired-delta and A/A-floor summarizer |
| `scripts/alloc.sh`, `scripts/alloc-report.py` | the allocation-region driver and differ |
| `scripts/output-hashes.sh` | the published-package digest comparison |

## How to reproduce

```sh
# deterministic counts first
scripts/counts.sh before                       # then: scripts/counts.sh after
COUNTS=counts python3 scripts/symbols.py before after

# allocations
scripts/alloc.sh alloc/ && python3 scripts/alloc-report.py

# value identity
scripts/output-hashes.sh differential/

# the oracles
cargo test -p litchi-xlsx --lib change_0622 -- --nocapture
cargo test -p litchi-xlsx --lib change_0635 -- --nocapture

# paired timing last, in one window
scripts/timing.sh timing/ && python3 scripts/timing-summarize.py timing/
```

## What this packet does not contain

No registered claim, no cold-cache or physical-I/O measurement, no range-source or
concurrency result, and no measurement of a real producer *file*: change 0602
established that the source-backed value editor admits none of the 95 real `.xlsx`
fixtures, and change 0622's oracle funnel confirms that exactly one of the
repository's 391 real worksheet parts is accepted by the value-only planning
validator at all. The 0601 producer-*shaped* corpora are measured, and they are the
closest reachable approximation.
