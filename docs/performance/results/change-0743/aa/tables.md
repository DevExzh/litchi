| case | corpus | before p50 ms | after p50 ms | paired p50 change | 95% CI | paired mean change | before p95 ms | after p95 ms |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `docx_semantic_full_text` | docx-semantic-large | 3.165 | 3.164 | +0.21% | [-0.51%, +1.15%] | +0.94% | 3.247 | 3.531 |
| `docx_semantic_full_text` | docx-semantic-medium | 0.064 | 0.065 | +0.77% | [+0.38%, +1.17%] | +0.99% | 0.069 | 0.072 |
| `pptx_semantic_full_text` | pptx-semantic-large | 50.558 | 50.866 | +0.39% | [-0.56%, +2.05%] | +0.32% | 51.223 | 51.267 |
| `pptx_semantic_full_text` | pptx-semantic-medium | 0.593 | 0.600 | +1.14% | [+0.75%, +1.65%] | +0.09% | 0.600 | 0.606 |
| `pptx_semantic_noop_edit_save` | pptx-semantic-large | 28.195 | 28.461 | +1.02% | [-1.20%, +1.80%] | +1.51% | 28.868 | 28.998 |
| `pptx_semantic_noop_edit_save` | pptx-semantic-medium | 0.860 | 0.858 | -0.31% | [-1.65%, +1.50%] | -0.18% | 0.870 | 0.867 |
| `pptx_semantic_one_edit_save` | pptx-semantic-large | 55.896 | 56.208 | +0.56% | [-2.00%, +4.24%] | -0.20% | 57.804 | 56.942 |
| `pptx_semantic_one_edit_save` | pptx-semantic-medium | 1.722 | 1.718 | -0.28% | [-0.93%, +0.68%] | -0.15% | 1.739 | 1.741 |
| `pptx_semantic_one_percent_edit_save` | pptx-semantic-large | 273.188 | 274.806 | +0.63% | [+0.05%, +1.64%] | +0.52% | 277.353 | 277.684 |
| `pptx_semantic_one_percent_edit_save` | pptx-semantic-medium | 1.723 | 1.720 | -0.25% | [-1.45%, +0.68%] | -0.25% | 1.741 | 1.738 |
| `pptx_semantic_open` | pptx-semantic-large | 2.045 | 2.154 | +6.92% | [-12.95%, +10.30%] | +5.66% | 2.238 | 2.260 |
| `pptx_semantic_open` | pptx-semantic-medium | 0.447 | 0.448 | -0.05% | [-0.93%, +1.26%] | -0.43% | 0.456 | 0.455 |

Regression flags (pair ratio > 1.05): 9

- `docx_semantic_full_text` docx-semantic-large p95 round 0 (positions 0/1): +11.39%
- `docx_semantic_full_text` docx-semantic-large p95 round 1 (positions 0/1): +17.09%
- `docx_semantic_full_text` docx-semantic-medium p95 round 0 (positions 0/1): +7.08%
- `docx_semantic_full_text` docx-semantic-medium p95 round 0 (positions 3/2): +6.85%
- `pptx_semantic_open` pptx-semantic-large p50 round 0 (positions 0/1): +10.30%
- `pptx_semantic_open` pptx-semantic-large p50 round 1 (positions 3/2): +9.30%
- `pptx_semantic_open` pptx-semantic-large p95 round 1 (positions 3/2): +12.62%
- `pptx_semantic_open` pptx-semantic-large mean round 0 (positions 0/1): +8.30%
- `pptx_semantic_open` pptx-semantic-large mean round 1 (positions 3/2): +9.28%
