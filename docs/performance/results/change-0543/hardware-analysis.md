# 0543 whole-child hardware diagnostics

Counters cover the complete fresh child, including corpus setup, planning, publication, and output oracles.
They remain diagnostics; they do not establish latency, operation-local cost, or an isolated planning counter.

Status: **pass**.

| Repeat | Shape | Cycles change | Instructions change | IPC change | Branch-miss change |
| ---: | --- | ---: | ---: | ---: | ---: |
| 1 | dense-sparse | -11.152% | -7.279% | 4.359% | 1.046% |
| 1 | medium | -0.725% | -3.246% | -2.539% | 3.519% |
| 2 | dense-sparse | -11.995% | -7.274% | 5.365% | 1.756% |
| 2 | medium | -3.680% | -3.541% | 0.145% | 4.322% |

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
