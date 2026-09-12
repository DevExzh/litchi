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

## candidate

Status: **pass**.

| Repeat | Shape | Planning Ir |
| ---: | --- | ---: |
| 1 | medium | 121,534,778 |
| 1 | dense-sparse | 225,065,014 |
| 2 | medium | 119,278,225 |
| 2 | dense-sparse | 225,112,535 |

## Comparison

| Repeat | Shape | Baseline Ir | Candidate Ir | Reduction | Gate |
| ---: | --- | ---: | ---: | ---: | ---: |
| 1 | dense-sparse | 237,178,343 | 225,065,014 | 5.107% | pass |
| 1 | medium | 125,648,982 | 121,534,778 | 3.274% | pass |
| 2 | dense-sparse | 237,077,736 | 225,112,535 | 5.047% | pass |
| 2 | medium | 125,610,546 | 119,278,225 | 5.041% | pass |

Conditional profile decision: **retainable-conditional-lane**.
