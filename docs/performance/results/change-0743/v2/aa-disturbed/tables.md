| case | corpus | before p50 ms | after p50 ms | paired p50 change | 95% CI | paired mean change | before p95 ms | after p95 ms |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `docx_semantic_full_text` | docx-semantic-large | 3.158 | 3.184 | +1.92% | [-48.58%, +95.31%] | +2.00% | 3.446 | 3.270 |
| `docx_semantic_full_text` | docx-semantic-medium | 0.064 | 0.064 | -1.03% | [-15.04%, +0.19%] | -1.75% | 0.072 | 0.068 |
| `pptx_semantic_full_text` | pptx-semantic-large | 50.266 | 50.486 | +0.24% | [-50.13%, +1.03%] | +0.13% | 51.089 | 51.228 |
| `pptx_semantic_full_text` | pptx-semantic-medium | 0.589 | 0.592 | +0.55% | [-0.20%, +0.95%] | +0.31% | 0.597 | 0.601 |
| `pptx_semantic_noop_edit_save` | pptx-semantic-large | 41.035 | 27.475 | -0.17% | [-50.95%, +1.62%] | +0.31% | 42.875 | 28.245 |
| `pptx_semantic_noop_edit_save` | pptx-semantic-medium | 0.845 | 0.846 | +0.18% | [-0.50%, +0.31%] | +0.12% | 0.853 | 0.854 |
| `pptx_semantic_one_edit_save` | pptx-semantic-large | 81.194 | 54.396 | +0.14% | [-50.83%, +1.14%] | -4.82% | 150.700 | 56.884 |
| `pptx_semantic_one_edit_save` | pptx-semantic-medium | 1.698 | 1.689 | -0.29% | [-0.80%, -0.08%] | -0.51% | 1.723 | 1.707 |
| `pptx_semantic_one_percent_edit_save` | pptx-semantic-large | 266.267 | 269.510 | +1.09% | [-59.44%, +108.83%] | -1.51% | 362.485 | 282.909 |
| `pptx_semantic_one_percent_edit_save` | pptx-semantic-medium | 1.700 | 1.692 | -0.56% | [-0.76%, -0.21%] | -0.56% | 1.716 | 1.707 |
| `pptx_semantic_open` | pptx-semantic-large | 2.076 | 2.063 | -0.66% | [-61.33%, +160.12%] | -0.91% | 2.254 | 2.200 |
| `pptx_semantic_open` | pptx-semantic-medium | 0.445 | 0.447 | +0.37% | [-1.28%, +1.71%] | +0.46% | 0.452 | 0.453 |

Regression flags (pair ratio > 1.05): 11

- `docx_semantic_full_text` docx-semantic-large p50 round 0 (positions 0/1): +95.31%
- `docx_semantic_full_text` docx-semantic-large p95 round 0 (positions 0/1): +150.98%
- `docx_semantic_full_text` docx-semantic-large mean round 0 (positions 0/1): +97.50%
- `pptx_semantic_noop_edit_save` pptx-semantic-large p95 round 0 (positions 3/2): +227.99%
- `pptx_semantic_noop_edit_save` pptx-semantic-large mean round 0 (positions 3/2): +29.01%
- `pptx_semantic_one_percent_edit_save` pptx-semantic-large p50 round 0 (positions 3/2): +108.83%
- `pptx_semantic_one_percent_edit_save` pptx-semantic-large p95 round 0 (positions 3/2): +152.07%
- `pptx_semantic_one_percent_edit_save` pptx-semantic-large mean round 0 (positions 3/2): +124.54%
- `pptx_semantic_open` pptx-semantic-large p50 round 0 (positions 0/1): +160.12%
- `pptx_semantic_open` pptx-semantic-large p95 round 0 (positions 0/1): +148.78%
- `pptx_semantic_open` pptx-semantic-large mean round 0 (positions 0/1): +99.30%
