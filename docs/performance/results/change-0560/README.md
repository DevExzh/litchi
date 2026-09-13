# change-0560 evidence packet

Change record: [`docs/performance/0560-xls-single-observation-freshness.md`](../../0560-xls-single-observation-freshness.md).
Disposition: retained. `performance_claim: none`; no claim-registry entry is added.

## Contents

| Path | What it is |
| --- | --- |
| `plan.json` | Frozen before capture: hypothesis, equivalence argument and acceptance gates. |
| `latency/` | The A/E/E/A children for ten XLS selectors, plus the six `xls_source_attribution` children that carry the deterministic observation counters. |
| `syscalls/` | Six `strace -f -c` children over the file-source XLS matrix. |
| `analysis.json` | Deterministic counters, syscall counts, every latency comparison in both directions, and the explicit review-trigger list with absolute nanosecond values. |

## Replay

```sh
python3 -B -c "import json;print(json.load(open('docs/performance/results/change-0560/analysis.json'))['summary'])"
```

## What is not here

No cold-cache, physical-device, remote/range-source, peak-RSS, allocation,
concurrency-scaling, real-producer or cross-platform capture. The `xls-tiny`
corpora measure below this harness clock's useful resolution; their cells are
retained and reported but carry no weight.
