| case | corpus | before p50 ms | after p50 ms | paired p50 change | 95% CI | paired mean change | before p95 ms | after p95 ms |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `docx_semantic_full_text` | docx-semantic-large | 3.164 | 3.190 | +0.66% | [-0.36%, +3.33%] | +0.96% | 3.261 | 3.255 |
| `docx_semantic_full_text` | docx-semantic-medium | 0.065 | 0.064 | -0.04% | [-0.86%, +0.60%] | +0.00% | 0.070 | 0.071 |
| `pptx_semantic_full_text` | pptx-semantic-large | 50.212 | 50.439 | +0.39% | [-2.61%, +1.12%] | +0.66% | 50.631 | 51.151 |
| `pptx_semantic_full_text` | pptx-semantic-medium | 0.592 | 0.592 | -0.07% | [-0.48%, +0.03%] | -0.46% | 0.598 | 0.604 |
| `pptx_semantic_noop_edit_save` | pptx-semantic-large | 27.606 | 27.552 | -0.60% | [-1.36%, +2.30%] | -0.69% | 28.108 | 28.028 |
| `pptx_semantic_noop_edit_save` | pptx-semantic-medium | 0.847 | 0.852 | +0.64% | [-0.39%, +1.03%] | +0.60% | 0.854 | 0.857 |
| `pptx_semantic_one_edit_save` | pptx-semantic-large | 54.510 | 54.447 | +0.30% | [-1.38%, +0.97%] | +0.33% | 55.119 | 55.232 |
| `pptx_semantic_one_edit_save` | pptx-semantic-medium | 1.701 | 1.710 | +0.80% | [-0.61%, +0.94%] | +0.78% | 1.717 | 1.724 |
| `pptx_semantic_one_percent_edit_save` | pptx-semantic-large | 266.571 | 268.228 | +0.40% | [+0.02%, +1.70%] | +0.42% | 270.183 | 273.087 |
| `pptx_semantic_one_percent_edit_save` | pptx-semantic-medium | 1.694 | 1.704 | +0.63% | [-0.49%, +1.57%] | +0.53% | 1.707 | 1.716 |
| `pptx_semantic_open` | pptx-semantic-large | 2.102 | 2.035 | -3.27% | [-7.96%, +5.04%] | -1.83% | 2.245 | 2.243 |
| `pptx_semantic_open` | pptx-semantic-medium | 0.446 | 0.451 | +0.93% | [+0.00%, +1.78%] | +1.00% | 0.453 | 0.458 |

Regression flags (pair ratio > 1.05): 3

- `docx_semantic_full_text` docx-semantic-large p95 round 1 (positions 3/2): +13.28%
- `pptx_semantic_open` pptx-semantic-large p50 round 1 (positions 0/1): +5.04%
- `pptx_semantic_open` pptx-semantic-large p95 round 0 (positions 3/2): +13.33%
