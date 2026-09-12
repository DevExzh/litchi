# 0531 whole-child hardware diagnostics

Counters cover the complete fresh child, including corpus setup, planning, publication, and output oracles.
They remain diagnostics; they do not establish latency, operation-local cost, or an isolated planning counter.

Status: **pass**.

| Repeat | Shape | Cycles change | Instructions change | IPC change | Branch-miss change |
| ---: | --- | ---: | ---: | ---: | ---: |
| 1 | dense-sparse | 4.591% | 0.802% | -3.623% | -2.405% |
| 1 | medium | 2.727% | 0.938% | -1.742% | 3.625% |
| 2 | dense-sparse | -3.996% | -4.331% | -0.349% | 0.580% |
| 2 | medium | -1.719% | -2.449% | -0.743% | 3.699% |

Every row retains the full event coverage and multiplexing record in the JSON report.

## baseline

Capture rows: **4**.

- `hardware-r1-medium`: measured; grouped coverage `complete`; multiplexed events `none`.
- `hardware-r1-dense-sparse`: measured; grouped coverage `complete`; multiplexed events `none`.
- `hardware-r2-medium`: measured; grouped coverage `complete`; multiplexed events `none`.
- `hardware-r2-dense-sparse`: measured; grouped coverage `complete`; multiplexed events `none`.

## candidate

Capture rows: **4**.

- `hardware-r1-medium`: measured; grouped coverage `complete`; multiplexed events `none`.
- `hardware-r1-dense-sparse`: measured; grouped coverage `complete`; multiplexed events `none`.
- `hardware-r2-medium`: measured; grouped coverage `complete`; multiplexed events `none`.
- `hardware-r2-dense-sparse`: measured; grouped coverage `complete`; multiplexed events `none`.
