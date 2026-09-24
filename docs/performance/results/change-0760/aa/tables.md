| case | corpus | before p50 ms | after p50 ms | paired p50 change | 95% CI | paired mean change | before p95 ms | after p95 ms |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `pptx_semantic_full_text` | pptx-semantic-large | 27.357 | 27.217 | -0.51% | [-5.89%, +3.67%] | -0.67% | 28.093 | 27.715 |
| `pptx_semantic_full_text` | pptx-semantic-medium | 0.343 | 0.342 | -0.28% | [-0.83%, +0.25%] | -0.36% | 0.353 | 0.352 |
| `pptx_semantic_noop_edit_save` | pptx-semantic-large | 23.662 | 23.891 | +0.10% | [-0.83%, +3.02%] | -0.06% | 24.418 | 24.642 |
| `pptx_semantic_noop_edit_save` | pptx-semantic-medium | 0.815 | 0.818 | +0.16% | [-0.59%, +0.70%] | +0.37% | 0.822 | 0.828 |
| `pptx_semantic_one_edit_save` | pptx-semantic-large | 45.830 | 46.042 | -0.18% | [-0.82%, +2.93%] | +0.14% | 46.334 | 46.736 |
| `pptx_semantic_one_edit_save` | pptx-semantic-medium | 1.520 | 1.527 | +0.25% | [-0.49%, +0.65%] | +0.23% | 1.531 | 1.536 |
| `pptx_semantic_one_percent_edit_save` | pptx-semantic-large | 149.383 | 150.083 | +0.37% | [+0.02%, +1.50%] | +0.76% | 150.269 | 152.541 |
| `pptx_semantic_one_percent_edit_save` | pptx-semantic-medium | 1.520 | 1.524 | -0.11% | [-0.40%, +0.71%] | -0.43% | 1.536 | 1.537 |

Phases, `pptx_semantic_opened_transaction_phases` on pptx-semantic-large (median of per-process medians, ms):

| phase | before | after | change |
| --- | ---: | ---: | ---: |
| opened_presentation_ns | 23.359 | 23.620 | +1.12% |
| snapshot_edit_ns | 0.061 | 0.054 | -10.58% |
| set_shape_text_ns | 0.780 | 0.770 | -1.36% |
| transaction_commit_ns | 20.908 | 21.151 | +1.16% |
| apply_commit_ns | 0.290 | 0.237 | -18.11% |
| publication_ns | 0.448 | 0.391 | -12.61% |
| total_ns | 46.039 | 46.373 | +0.73% |

Phases, `pptx_semantic_opened_transaction_phases` on pptx-semantic-medium (median of per-process medians, ms):

| phase | before | after | change |
| --- | ---: | ---: | ---: |
| opened_presentation_ns | 0.717 | 0.716 | -0.11% |
| snapshot_edit_ns | 0.014 | 0.014 | +0.98% |
| set_shape_text_ns | 0.086 | 0.085 | -0.79% |
| transaction_commit_ns | 0.538 | 0.538 | -0.01% |
| apply_commit_ns | 0.063 | 0.064 | +1.62% |
| publication_ns | 0.095 | 0.094 | -0.46% |
| total_ns | 1.511 | 1.513 | +0.16% |

Regression flags (pair ratio > 1.05): 0


Whole-process user instructions and cycles (perf stat; includes corpus construction and verification):

| group | before instructions | after instructions | paired change | before cycles | after cycles | paired change |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| large | 194.511 G | 194.513 G | +0.00% | 46.979 G | 46.817 G | -0.35% |
| medium | 13.682 G | 13.682 G | -0.00% | 4.368 G | 4.375 G | -0.18% |
| phases | 46.565 G | 46.564 G | -0.00% | 11.749 G | 11.739 G | +0.27% |
