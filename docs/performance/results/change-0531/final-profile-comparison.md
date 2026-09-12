# 0531 conditional planning profile

This report validates the source-bound XLSX `edit_sheets` Callgrind lane.
Planning Ir is conditional mechanism evidence and is not converted to latency.

Status: **pass**.

## baseline

Status: **pass**.

| Repeat | Shape | Planning Ir |
| ---: | --- | ---: |
| 1 | medium | 125,648,982 |
| 1 | dense-sparse | 237,178,343 |
| 2 | medium | 125,610,546 |
| 2 | dense-sparse | 237,077,736 |

## final

Status: **pass**.

| Repeat | Shape | Planning Ir |
| ---: | --- | ---: |
| 1 | medium | 119,322,888 |
| 1 | dense-sparse | 225,119,198 |
| 2 | medium | 119,304,260 |
| 2 | dense-sparse | 225,067,010 |

## Comparison

| Repeat | Shape | Baseline Ir | Candidate Ir | Reduction | Gate |
| ---: | --- | ---: | ---: | ---: | ---: |
| 1 | dense-sparse | 237,178,343 | 225,119,198 | 5.084% | pass |
| 1 | medium | 125,648,982 | 119,322,888 | 5.035% | pass |
| 2 | dense-sparse | 237,077,736 | 225,067,010 | 5.066% | pass |
| 2 | medium | 125,610,546 | 119,304,260 | 5.021% | pass |

Diagnostic profile decision: **instruction-gate-passed-diagnostic-only**.
