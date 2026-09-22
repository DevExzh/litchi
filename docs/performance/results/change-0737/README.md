# 0737 unchanged-owner PPT lifecycle qualification

The revised probe is diagnostic only. Its secondary legacy p50 differs from
the independently rebuilt original by +6.67%; it is not an interchangeable
ordinary-path baseline. The rejected 0735 production candidate remains absent.
See [the report](../../0737-ppt-oracle-lifecycle-controls.md).

The fixed matrix has 108 native processes (4,518 samples), 30 allocation
processes, both PPT fixtures, CPU 12, serial execution and rotated arm order.
Archive/new-legacy, duplicate legacy A/A, matched strict retained/drained,
fresh-child and zero/three-warmup allocator controls remain separate contrasts.
Every measured output and all 198 strict warmup receipts match the prior full
oracle. Compact receipt values still accumulate; witness counts are not bytes.

Offline replay, from the repository root with the recorded production source:

```sh
python3 -B docs/performance/results/change-0737/analyze.py
python3 -B docs/performance/results/change-0737/audit.py
python3 -B docs/performance/results/change-0737/negative-contract.py
python3 -B docs/performance/results/change-0737/artifact-seal.py --check
```

Exact deleted-binary identities are supplied by `cleanup.json`. The replay
checks the full current source census, constraints, frozen inputs, exact
commands, schedule, raw captures, full oracles and independently recomputed
statistics. `negative-contract.py` rejects 17 altered reports. The independent
audit shares report/custody validation but implements its own statistics.

For new native measurements, make a new packet and owned scratch paths rather
than modifying this sealed packet. Start with the source revision in
`base.json`, run the original and controls builds through `build.py`, and run
both variants through `qualify.py`. Review the oracle contract and prospective
plan, run report corruption controls, freeze the inputs with `run.py freeze`,
run `preflight.py`, then execute `run.py capture`. `build.py` records all five
quality commands, complete source/probe identity and exact native/allocation
binaries. All Cargo and native commands must run serially.

The synthetic preflight is schema/statistical integration evidence only. It
uses temporary copies of qualified reports and never writes synthetic values
to the real captures directory. Its extra check deliberately changes one
derived statistic and requires the independent audit to reject it.

The failed Clippy attempt, subsequent successful build and documentation-only
rebuild remain archived. No older packet is modified. `terminal.json` records
post-cleanup replay, and the artifact seal inventories every final packet file.
There is no new owner test, RSS, cold-I/O, concurrency or production speedup claim.
