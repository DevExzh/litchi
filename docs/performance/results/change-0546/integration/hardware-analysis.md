# 0546 whole-child hardware diagnostics

Counters cover the complete fresh child, including corpus setup, planning, publication, and output oracles.
They remain diagnostics; they do not establish latency, operation-local cost, or an isolated planning counter.

Status: **pass**.

| Repeat | Shape | Cycles change | Instructions change | IPC change | Branch-miss change |
| ---: | --- | ---: | ---: | ---: | ---: |
| 1 | dense-sparse | -10.870% | -6.751% | 4.621% | 0.841% |
| 1 | medium | -4.213% | -3.138% | 1.122% | 0.338% |
| 2 | dense-sparse | -3.847% | -3.067% | 0.812% | 1.537% |
| 2 | medium | -4.513% | -3.658% | 0.895% | 0.016% |

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
