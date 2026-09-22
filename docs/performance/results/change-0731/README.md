# 0731 public PPT instruction attribution

[Report](../../0731-ppt-public-instruction-attribution.md): production unchanged,
no speedup claim. The most important result is an interpretation limit:
Callgrind's 67.78% artifact-hash instruction fraction uses software SHA while
Valgrind masks the host's SHA feature. It cannot nominate a native bottleneck.

The frozen matrix contains three native processes (50 samples, three warmups),
three allocation processes and three Callgrind processes (one sample, no warmup).
All use the existing PPT slide-removal workflow and exact sealed 0728 oracle.
A wrapper in the packet probe is the sole collection boundary; no production
API or feature changed. `plan.json`, `hypothesis.md`, `freeze.json`, `build.json`
and the raw capture manifest bind the evidence. The post-profile feature witness
is an interpretation control, separate from the fixed measurement matrix.

Root ran all Cargo/native tools serially. `run.py build` performs five probe
gates and builds both binaries; `run.py freeze` and `run.py capture` fix and run
the matrix. The scripts deliberately refuse to overwrite existing capture roots.
Reproduction should use an isolated checkout at the recorded source revision,
with an empty owned packet/build output destination; preserve the recorded
probe and plan. Build products were removed after their hashes were recorded.

Offline replay from the repository root:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0731/analyze.py
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0731/negative-checks.py
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0731/artifact-seal.py --check
```

`analysis.json` contains every native process statistic, all allocation fields,
and full raw Callgrind self/edge arithmetic. Native timing printed by an
instrumented process is retained only as raw evidence and is never interpreted
as native latency. `source-review.md` maps the implementation; `review.md`
records independent interpretation. Cleanup and terminal receipts preserve
binary identity and post-cleanup checks. No iWork files are changed.
