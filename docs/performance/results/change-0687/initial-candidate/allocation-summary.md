# Allocation and retention changes

Separate instrumented probes, three identical repeats per group. Bytes are
allocator requests/live gauges, not process RSS or logical-budget guarantees.
All phases, including unchanged rows, remain in `allocation-comparison.json`.

| Case | Source | Query | Calls before → after | Requested bytes before → after | Peak live before → after | Retained delta before → after |
|---|---|---|---:|---:|---:|---:|
