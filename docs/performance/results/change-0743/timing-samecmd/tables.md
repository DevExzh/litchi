| case | corpus | before p50 ms | after p50 ms | paired p50 change | 95% CI | paired mean change | before p95 ms | after p95 ms |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `docx_semantic_full_text` | docx-semantic-large | 3.155 | 3.145 | -0.16% | [-1.12%, +0.58%] | -0.03% | 3.226 | 3.237 |
| `docx_semantic_full_text` | docx-semantic-medium | 0.064 | 0.065 | +0.55% | [-1.61%, +2.95%] | +0.15% | 0.070 | 0.071 |
| `pptx_semantic_full_text` | pptx-semantic-large | 50.706 | 28.153 | -44.28% | [-45.03%, -43.97%] | -44.26% | 51.176 | 28.658 |
| `pptx_semantic_full_text` | pptx-semantic-medium | 0.593 | 0.353 | -40.47% | [-40.99%, -40.04%] | -40.15% | 0.599 | 0.364 |
| `pptx_semantic_noop_edit_save` | pptx-semantic-large | 28.234 | 23.961 | -15.08% | [-16.96%, -14.37%] | -15.23% | 28.733 | 24.276 |
| `pptx_semantic_noop_edit_save` | pptx-semantic-medium | 0.858 | 0.808 | -5.64% | [-6.27%, -5.41%] | -5.61% | 0.867 | 0.815 |
| `pptx_semantic_one_edit_save` | pptx-semantic-large | 55.781 | 26.510 | -52.45% | [-53.03%, -51.75%] | -52.36% | 56.575 | 27.012 |
| `pptx_semantic_one_edit_save` | pptx-semantic-medium | 1.724 | 1.315 | -23.81% | [-24.10%, -23.46%] | -23.75% | 1.742 | 1.326 |
| `pptx_semantic_one_percent_edit_save` | pptx-semantic-large | 275.419 | 147.821 | -46.50% | [-46.70%, -46.08%] | -46.43% | 279.201 | 148.677 |
| `pptx_semantic_one_percent_edit_save` | pptx-semantic-medium | 1.724 | 1.314 | -23.86% | [-24.10%, -23.24%] | -23.79% | 1.741 | 1.325 |
| `pptx_semantic_open` | pptx-semantic-large | 2.171 | 2.206 | +1.34% | [-3.60%, +8.76%] | +3.81% | 2.273 | 2.317 |
| `pptx_semantic_open` | pptx-semantic-medium | 0.447 | 0.456 | +2.20% | [+1.59%, +2.91%] | +2.37% | 0.454 | 0.465 |

Phases, `pptx_semantic_opened_transaction_phases` on pptx-semantic-large (median of per-process medians, ms):

| phase | before | after | change |
| --- | ---: | ---: | ---: |
| opened_presentation_ns | 27.909 | 23.499 | -15.80% |
| snapshot_edit_ns | 0.051 | 0.046 | -10.10% |
| set_shape_text_ns | 1.217 | 0.750 | -38.38% |
| transaction_commit_ns | 26.226 | 1.543 | -94.11% |
| apply_commit_ns | 0.210 | 0.202 | -3.86% |
| publication_ns | 0.376 | 0.342 | -9.00% |
| total_ns | 55.852 | 26.378 | -52.77% |

Phases, `pptx_semantic_opened_transaction_phases` on pptx-semantic-medium (median of per-process medians, ms):

| phase | before | after | change |
| --- | ---: | ---: | ---: |
| opened_presentation_ns | 0.768 | 0.721 | -6.11% |
| snapshot_edit_ns | 0.014 | 0.013 | -0.04% |
| set_shape_text_ns | 0.129 | 0.084 | -34.88% |
| transaction_commit_ns | 0.659 | 0.347 | -47.41% |
| apply_commit_ns | 0.058 | 0.058 | +0.77% |
| publication_ns | 0.089 | 0.089 | +0.21% |
| total_ns | 1.717 | 1.316 | -23.36% |

Regression flags (pair ratio > 1.05): 8

- `docx_semantic_full_text` docx-semantic-large p95 round 2 (positions 0/1): +14.39%
- `pptx_semantic_open` pptx-semantic-large p50 round 1 (positions 0/1): +8.76%
- `pptx_semantic_open` pptx-semantic-large p50 round 1 (positions 3/2): +13.13%
- `pptx_semantic_open` pptx-semantic-large p95 round 1 (positions 3/2): +14.01%
- `pptx_semantic_open` pptx-semantic-large mean round 0 (positions 0/1): +6.73%
- `pptx_semantic_open` pptx-semantic-large mean round 0 (positions 3/2): +5.04%
- `pptx_semantic_open` pptx-semantic-large mean round 1 (positions 0/1): +8.07%
- `pptx_semantic_open` pptx-semantic-large mean round 1 (positions 3/2): +12.17%
