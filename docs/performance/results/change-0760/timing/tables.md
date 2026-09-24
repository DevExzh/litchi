| case | corpus | before p50 ms | after p50 ms | paired p50 change | 95% CI | paired mean change | before p95 ms | after p95 ms |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `pptx_semantic_full_text` | pptx-semantic-large | 27.307 | 27.378 | +0.62% | [-1.89%, +2.18%] | +0.28% | 28.058 | 27.922 |
| `pptx_semantic_full_text` | pptx-semantic-medium | 0.342 | 0.344 | +0.75% | [-1.44%, +1.21%] | +0.15% | 0.353 | 0.353 |
| `pptx_semantic_noop_edit_save` | pptx-semantic-large | 24.003 | 24.108 | +0.19% | [-0.66%, +1.31%] | +0.72% | 24.240 | 24.890 |
| `pptx_semantic_noop_edit_save` | pptx-semantic-medium | 0.814 | 0.817 | +0.14% | [-0.45%, +0.99%] | +0.26% | 0.822 | 0.828 |
| `pptx_semantic_one_edit_save` | pptx-semantic-large | 45.884 | 26.774 | -41.79% | [-42.87%, -40.97%] | -41.93% | 47.066 | 27.326 |
| `pptx_semantic_one_edit_save` | pptx-semantic-medium | 1.522 | 1.332 | -12.64% | [-13.32%, -12.06%] | -12.36% | 1.535 | 1.346 |
| `pptx_semantic_one_percent_edit_save` | pptx-semantic-large | 150.345 | 150.203 | -0.09% | [-0.28%, +0.10%] | -0.15% | 152.293 | 150.938 |
| `pptx_semantic_one_percent_edit_save` | pptx-semantic-medium | 1.517 | 1.331 | -12.45% | [-12.75%, -11.89%] | -12.45% | 1.530 | 1.344 |

Phases, `pptx_semantic_opened_transaction_phases` on pptx-semantic-large (median of per-process medians, ms):

| phase | before | after | change |
| --- | ---: | ---: | ---: |
| opened_presentation_ns | 23.720 | 23.785 | +0.28% |
| snapshot_edit_ns | 0.066 | 0.062 | -5.81% |
| set_shape_text_ns | 0.783 | 0.788 | +0.68% |
| transaction_commit_ns | 21.193 | 1.607 | -92.42% |
| apply_commit_ns | 0.324 | 0.267 | -17.43% |
| publication_ns | 0.467 | 0.426 | -8.68% |
| total_ns | 46.317 | 26.952 | -41.81% |

Phases, `pptx_semantic_opened_transaction_phases` on pptx-semantic-medium (median of per-process medians, ms):

| phase | before | after | change |
| --- | ---: | ---: | ---: |
| opened_presentation_ns | 0.722 | 0.724 | +0.21% |
| snapshot_edit_ns | 0.014 | 0.014 | +1.45% |
| set_shape_text_ns | 0.086 | 0.087 | +0.42% |
| transaction_commit_ns | 0.540 | 0.346 | -36.00% |
| apply_commit_ns | 0.064 | 0.065 | +1.22% |
| publication_ns | 0.098 | 0.097 | -0.49% |
| total_ns | 1.525 | 1.336 | -12.42% |

Regression flags (pair ratio > 1.05): 7

- `pptx_semantic_full_text` pptx-semantic-large p95 round 3 (positions 3/2): +320.06%
- `pptx_semantic_full_text` pptx-semantic-large mean round 3 (positions 3/2): +68.13%
- `pptx_semantic_noop_edit_save` pptx-semantic-large p95 round 0 (positions 3/2): +6.81%
- `pptx_semantic_noop_edit_save` pptx-semantic-large p95 round 1 (positions 0/1): +5.92%
- `pptx_semantic_noop_edit_save` pptx-semantic-large p95 round 3 (positions 3/2): +158.40%
- `pptx_semantic_noop_edit_save` pptx-semantic-large mean round 3 (positions 3/2): +10.30%
- `pptx_semantic_one_percent_edit_save` pptx-semantic-large p95 round 3 (positions 3/2): +5.12%

Whole-process user instructions and cycles (perf stat; includes corpus construction and verification):

| group | before instructions | after instructions | paired change | before cycles | after cycles | paired change |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| large | 194.510 G | 187.788 G | -3.46% | 47.055 G | 45.599 G | -3.01% |
| medium | 13.681 G | 13.179 G | -3.67% | 4.358 G | 4.335 G | -0.44% |
| phases | 46.564 G | 39.815 G | -14.49% | 11.870 G | 10.398 G | -13.04% |
