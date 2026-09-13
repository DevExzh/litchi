# change-0559 evidence packet

Change record: [`docs/performance/0559-cfb-ascii-simple-uppercase.md`](../../0559-cfb-ascii-simple-uppercase.md).
Disposition: retained. `performance_claim: none`; no claim-registry entry is added.

## Contents

| Path | What it is |
| --- | --- |
| `plan.json` | Frozen before the candidate build and before any capture. |
| `run.py`, `analyze.py` | The same driver and analyzer used for change 0558, run with change 0558 as the baseline stage so this change is measured in isolation. |
| `latency/`, `syscalls/` | 48 + 12 children proving this change is neutral on the source-backed XLS path: `version()` calls, `statx`, `read_at` calls and read bytes are byte-identical. |
| `guardrail/` | 84 children across 21 OLE2 selectors, pre-0558 baseline versus the combined tree, ASLR disabled. |
| `profiles/` | Callgrind branch-simulation dumps for five OLE2 cases, baseline (`-A`) and combined tree (`-C`). |
| `collect_counters.py`, `counters.json` | The deterministic Ir/branch table built from those dumps. |
| `analyze_guardrail.py`, `guardrail-analysis.json` | Per-case both-direction comparison over the combined tree. |

## Replay

```sh
python3 -B docs/performance/results/change-0559/analyze.py \
  --packet docs/performance/results/change-0559 \
  --output docs/performance/results/change-0559/analysis.json
python3 -B docs/performance/results/change-0559/collect_counters.py \
  --profile-dir docs/performance/results/change-0559/profiles
python3 -B docs/performance/results/change-0559/analyze_guardrail.py
```

## What is not here

No cold-cache, physical-device, remote/range-source, peak-RSS, allocation,
concurrency-scaling, real-producer or cross-platform capture. No corpus in the
matrix carries a non-ASCII directory-entry name, so the cost of the extra branch
on the Unicode path is bounded by inspection rather than measurement.
