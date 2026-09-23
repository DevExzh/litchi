| case | corpus | before p50 ms | after p50 ms | paired p50 change | 95% CI | paired mean change | before p95 ms | after p95 ms |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `docx_semantic_full_text` | docx-semantic-large | 3.156 | 3.117 | -0.97% | [-2.46%, -0.48%] | -0.72% | 3.219 | 3.227 |
| `docx_semantic_full_text` | docx-semantic-medium | 0.064 | 0.064 | -0.51% | [-1.77%, +0.47%] | -1.01% | 0.070 | 0.068 |
| `pptx_semantic_full_text` | pptx-semantic-large | 50.590 | 27.690 | -45.13% | [-45.88%, -44.08%] | -45.02% | 51.116 | 28.284 |
| `pptx_semantic_full_text` | pptx-semantic-medium | 0.589 | 0.350 | -40.68% | [-41.47%, -40.03%] | -40.36% | 0.597 | 0.358 |
| `pptx_semantic_noop_edit_save` | pptx-semantic-large | 27.780 | 23.620 | -14.55% | [-15.86%, -10.25%] | -14.69% | 28.512 | 24.040 |
| `pptx_semantic_noop_edit_save` | pptx-semantic-medium | 0.854 | 0.804 | -5.71% | [-6.09%, -5.25%] | -5.82% | 0.862 | 0.810 |
| `pptx_semantic_one_edit_save` | pptx-semantic-large | 55.656 | 26.181 | -52.55% | [-53.66%, -50.55%] | -52.34% | 56.586 | 26.710 |
| `pptx_semantic_one_edit_save` | pptx-semantic-medium | 1.709 | 1.311 | -23.40% | [-23.68%, -23.26%] | -23.36% | 1.722 | 1.323 |
| `pptx_semantic_one_percent_edit_save` | pptx-semantic-large | 272.448 | 146.112 | -46.29% | [-46.69%, -45.46%] | -46.14% | 275.756 | 147.543 |
| `pptx_semantic_one_percent_edit_save` | pptx-semantic-medium | 1.710 | 1.308 | -23.51% | [-23.76%, -23.03%] | -23.50% | 1.726 | 1.319 |
| `pptx_semantic_open` | pptx-semantic-large | 2.027 | 2.042 | +1.44% | [+0.43%, +2.67%] | +0.21% | 2.225 | 2.091 |
| `pptx_semantic_open` | pptx-semantic-medium | 0.447 | 0.451 | +1.03% | [+0.38%, +1.84%] | +1.43% | 0.452 | 0.461 |

Phases, `pptx_semantic_opened_transaction_phases` on pptx-semantic-large (median of per-process medians, ms):

| phase | before | after | change |
| --- | ---: | ---: | ---: |
| opened_presentation_ns | 27.614 | 23.568 | -14.65% |
| snapshot_edit_ns | 0.044 | 0.044 | -0.86% |
| set_shape_text_ns | 1.200 | 0.745 | -37.94% |
| transaction_commit_ns | 25.725 | 1.513 | -94.12% |
| apply_commit_ns | 0.193 | 0.194 | +0.75% |
| publication_ns | 0.329 | 0.327 | -0.80% |
| total_ns | 55.402 | 26.375 | -52.39% |

Phases, `pptx_semantic_opened_transaction_phases` on pptx-semantic-medium (median of per-process medians, ms):

| phase | before | after | change |
| --- | ---: | ---: | ---: |
| opened_presentation_ns | 0.768 | 0.719 | -6.43% |
| snapshot_edit_ns | 0.014 | 0.014 | +0.26% |
| set_shape_text_ns | 0.128 | 0.084 | -34.63% |
| transaction_commit_ns | 0.661 | 0.344 | -47.86% |
| apply_commit_ns | 0.058 | 0.058 | +0.63% |
| publication_ns | 0.089 | 0.089 | -0.74% |
| total_ns | 1.720 | 1.310 | -23.84% |

Regression flags (pair ratio > 1.05): 2

- `docx_semantic_full_text` docx-semantic-large p95 round 1 (positions 0/1): +16.30%
- `pptx_semantic_open` pptx-semantic-large p50 round 0 (positions 3/2): +5.28%
