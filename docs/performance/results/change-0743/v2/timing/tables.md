| case | corpus | before p50 ms | after p50 ms | paired p50 change | 95% CI | paired mean change | before p95 ms | after p95 ms |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `docx_semantic_full_text` | docx-semantic-large | 3.197 | 3.167 | -0.44% | [-1.43%, +0.22%] | -0.47% | 3.264 | 3.285 |
| `docx_semantic_full_text` | docx-semantic-medium | 0.065 | 0.064 | -0.81% | [-2.59%, +1.94%] | +0.68% | 0.071 | 0.070 |
| `pptx_semantic_full_text` | pptx-semantic-large | 50.688 | 27.137 | -46.47% | [-47.48%, -46.18%] | -46.43% | 51.521 | 27.539 |
| `pptx_semantic_full_text` | pptx-semantic-medium | 0.592 | 0.340 | -42.62% | [-42.83%, -41.82%] | -42.39% | 0.598 | 0.350 |
| `pptx_semantic_noop_edit_save` | pptx-semantic-large | 27.926 | 23.171 | -16.42% | [-17.36%, -15.79%] | -16.58% | 28.563 | 23.506 |
| `pptx_semantic_noop_edit_save` | pptx-semantic-medium | 0.851 | 0.807 | -5.05% | [-5.85%, -4.13%] | -5.04% | 0.858 | 0.814 |
| `pptx_semantic_one_edit_save` | pptx-semantic-large | 55.570 | 44.547 | -18.72% | [-20.47%, -17.40%] | -18.69% | 56.192 | 45.192 |
| `pptx_semantic_one_edit_save` | pptx-semantic-medium | 1.710 | 1.506 | -12.33% | [-12.59%, -9.90%] | -12.32% | 1.724 | 1.522 |
| `pptx_semantic_one_percent_edit_save` | pptx-semantic-large | 269.608 | 144.746 | -45.97% | [-46.20%, -45.59%] | -45.78% | 273.218 | 147.150 |
| `pptx_semantic_one_percent_edit_save` | pptx-semantic-medium | 1.712 | 1.507 | -12.33% | [-12.70%, -10.10%] | -12.31% | 1.724 | 1.522 |
| `pptx_semantic_open` | pptx-semantic-large | 2.170 | 2.093 | -0.83% | [-5.58%, +2.68%] | -1.86% | 2.305 | 2.291 |
| `pptx_semantic_open` | pptx-semantic-medium | 0.448 | 0.453 | +1.61% | [+0.87%, +1.98%] | +1.67% | 0.454 | 0.462 |

Phases, `pptx_semantic_opened_transaction_phases` on pptx-semantic-large (median of per-process medians, ms):

| phase | before | after | change |
| --- | ---: | ---: | ---: |
| opened_presentation_ns | 27.204 | 23.215 | -14.66% |
| snapshot_edit_ns | 0.047 | 0.045 | -4.37% |
| set_shape_text_ns | 1.193 | 0.749 | -37.23% |
| transaction_commit_ns | 25.644 | 20.821 | -18.81% |
| apply_commit_ns | 0.201 | 0.200 | -0.85% |
| publication_ns | 0.346 | 0.339 | -2.21% |
| total_ns | 54.601 | 45.484 | -16.70% |

Phases, `pptx_semantic_opened_transaction_phases` on pptx-semantic-medium (median of per-process medians, ms):

| phase | before | after | change |
| --- | ---: | ---: | ---: |
| opened_presentation_ns | 0.760 | 0.716 | -5.80% |
| snapshot_edit_ns | 0.013 | 0.013 | -0.52% |
| set_shape_text_ns | 0.128 | 0.084 | -34.21% |
| transaction_commit_ns | 0.651 | 0.535 | -17.85% |
| apply_commit_ns | 0.058 | 0.058 | -0.26% |
| publication_ns | 0.089 | 0.088 | -0.93% |
| total_ns | 1.698 | 1.491 | -12.21% |

Regression flags (pair ratio > 1.05): 7

- `docx_semantic_full_text` docx-semantic-large p95 round 0 (positions 3/2): +14.26%
- `docx_semantic_full_text` docx-semantic-medium mean round 0 (positions 0/1): +52.39%
- `pptx_semantic_one_edit_save` pptx-semantic-large p95 round 3 (positions 3/2): +101.94%
- `pptx_semantic_one_percent_edit_save` pptx-semantic-large p95 round 2 (positions 3/2): +127.74%
- `pptx_semantic_open` pptx-semantic-large p95 round 1 (positions 0/1): +5.40%
- `pptx_semantic_open` pptx-semantic-large p95 round 2 (positions 0/1): +6.78%
- `pptx_semantic_open` pptx-semantic-large p95 round 3 (positions 3/2): +17.28%
