| case | corpus | before p50 ms | after p50 ms | paired p50 change | 95% CI | paired mean change | before p95 ms | after p95 ms |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `docx_semantic_full_text` | docx-semantic-large | 3.131 | 3.204 | +1.79% | [+0.68%, +3.62%] | +1.95% | 3.207 | 3.302 |
| `docx_semantic_full_text` | docx-semantic-medium | 0.064 | 0.065 | +0.67% | [-0.87%, +2.09%] | +0.66% | 0.071 | 0.072 |
| `pptx_semantic_full_text` | pptx-semantic-large | 50.968 | 50.534 | -0.83% | [-1.54%, +0.02%] | -0.78% | 51.886 | 50.886 |
| `pptx_semantic_full_text` | pptx-semantic-medium | 0.591 | 0.589 | -0.25% | [-0.90%, -0.06%] | +0.03% | 0.597 | 0.597 |
| `pptx_semantic_noop_edit_save` | pptx-semantic-large | 28.372 | 28.092 | -0.79% | [-1.14%, -0.41%] | -0.67% | 28.834 | 28.658 |
| `pptx_semantic_noop_edit_save` | pptx-semantic-medium | 0.857 | 0.855 | -0.19% | [-0.97%, +0.91%] | -0.36% | 0.864 | 0.863 |
| `pptx_semantic_one_edit_save` | pptx-semantic-large | 55.936 | 55.566 | -1.04% | [-4.74%, +0.79%] | -1.12% | 57.157 | 56.371 |
| `pptx_semantic_one_edit_save` | pptx-semantic-medium | 1.719 | 1.717 | +0.33% | [-1.53%, +1.07%] | +0.36% | 1.741 | 1.732 |
| `pptx_semantic_one_percent_edit_save` | pptx-semantic-large | 275.754 | 275.012 | -0.29% | [-0.44%, +0.61%] | -0.29% | 279.650 | 278.187 |
| `pptx_semantic_one_percent_edit_save` | pptx-semantic-medium | 1.720 | 1.727 | +0.29% | [-1.06%, +1.16%] | +0.33% | 1.734 | 1.745 |
| `pptx_semantic_open` | pptx-semantic-large | 2.112 | 2.112 | -0.24% | [-0.77%, +6.88%] | -0.79% | 2.274 | 2.177 |
| `pptx_semantic_open` | pptx-semantic-medium | 0.448 | 0.446 | -0.44% | [-2.01%, +1.42%] | -0.09% | 0.453 | 0.454 |

Regression flags (pair ratio > 1.05): 3

- `docx_semantic_full_text` docx-semantic-large p95 round 0 (positions 3/2): +6.93%
- `docx_semantic_full_text` docx-semantic-medium p95 round 0 (positions 3/2): +7.41%
- `pptx_semantic_open` pptx-semantic-large p50 round 1 (positions 3/2): +6.88%
