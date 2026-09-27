# 0786 finite execution-budget scaling evidence

This packet measures current public low-level read sessions under finite root
worker, I/O-concurrency and CPU-task budgets. It does not change production
code. See `plan.json` for the frozen 120-case schedule and `host.json` for the
machine. The source is warm immutable memory; fresh means a newly prepared
session/package, not a physically cold source. Primed ordered Parts are a
cache-hit control. The OPC operation includes eager package opening, whereas
CFB/Parts operations start after metadata preparation; absolute route times
are not equivalent scopes.

The reusable executable lives in `tools/perf-execution`. Corpus generation,
byte/order verification and session cleanup are outside the wall interval.
CPU-time intervals surround the wall-clock reads and are slightly wider.
RSS is the entire child process, including corpus preparation, verification,
warmups and every sample. Source-counter diagnostics use a separate feature
build and are never pooled into native latency estimates.

Root runs Cargo and measured children serially. `build.py`, `quality.py`, and
`capture.py qualification|native|observer` retain logs, commands and hashes.
They refuse evidence overwrites. Local absolute paths in execution receipts
record historical custody; replay must resolve retained artifacts relative to
this packet after the temporary worktree is removed. Reproduction requires a
fresh output packet and target, not overwriting these measurements.

Offline replay, from this repository:

```sh
python3 -B docs/performance/results/change-0786/analyze.py --check
python3 -B docs/performance/results/change-0786/validate.py --require-final-seal
python3 -B docs/performance/results/change-0786/raw_audit.py --check
```

No hardware counters, allocator instrumentation, physical disk, artificial
range delay, cross-session contention or native-format CRUD is measured in
this packet. CPU utilization is process CPU time divided by wall time, not
an observed worker count. Requested-width efficiency and Amdahl fits are
scaling descriptions, not a causal decomposition or a production speedup.

The retained set has 120 qualification reports, 720 native reports, and 240
observer reports (22,200 measured samples). `raw_audit.py` independently
reconstructs the 96 member payloads across the three shapes and all 120 paired
speedup curves, including bootstrap endpoints. `compile-preflight-0` and
`quality-0` retain corrected pre-measurement failures; `quality-1` is the passing
six-gate run. No measured child was retried or excluded.
