# Sweep 0756: wave-wide before/after (009d515bef -> ddf788eb80)

Descriptive only, not a registered claim. Four processes per arm per group in the order A B B A A B B A (A = base 009d515bef, B = final ddf788eb80), each `taskset -c 20 ... --samples 9 --warmup 2`. Base/final p50 = median of the four per-process p50s; ratio = final / base (below 1 is faster); paired range = min-max of final/base over the four adjacent (A,B) process pairs. See README.md for caveats.

| Format | Case | Corpus | Grp | Base p50 (ms) | Final p50 (ms) | Ratio | Paired range | Notes |
|---|---|---|---|---:|---:|---:|---|---|
| DOCX | `docx_ordinary_save_lifecycle` | `docx-semantic-medium` | M | 5.636 | 5.672 | 1.006 | 0.998-1.011 | paired range straddles 1.0 |
| DOCX | `docx_semantic_full_text` | `docx-semantic-medium` | S | 0.06323 | 0.03878 | 0.613 | 0.607-0.615 |  |
| DOCX | `docx_semantic_full_text` | `docx-semantic-large` | S | 3.134 | 1.914 | 0.611 | 0.600-0.617 |  |
| DOCX | `docx_semantic_noop_edit_save` | `docx-semantic-medium` | S | 0.08258 | 0.03604 | 0.436 | 0.433-0.447 |  |
| DOCX | `docx_semantic_noop_edit_save` | `docx-semantic-large` | S | 3.82 | 1.506 | 0.394 | 0.391-0.402 |  |
| DOCX | `docx_semantic_one_edit_save` | `docx-semantic-medium` | S | 0.2295 | 0.1514 | 0.660 | 0.644-0.704 |  |
| DOCX | `docx_semantic_one_edit_save` | `docx-semantic-large` | S | 8.615 | 4.645 | 0.539 | 0.536-0.543 |  |
| DOCX | `docx_semantic_one_percent_edit_save` | `docx-semantic-medium` | S | 0.2321 | 0.1951 | 0.841 | 0.834-0.847 |  |
| DOCX | `docx_semantic_one_percent_edit_save` | `docx-semantic-large` | S | 8.898 | 6.903 | 0.776 | 0.768-0.781 |  |
| DOCX | `docx_semantic_open` | `docx-semantic-medium` | S | 0.06102 | 0.06258 | 1.025 | 1.013-1.037 |  |
| DOCX | `docx_semantic_open` | `docx-semantic-large` | S | 0.2069 | 0.2093 | 1.011 | 0.987-1.213 | paired range straddles 1.0 |
| DOCX | `docx_source_backed_one_edit_save` | `docx-source-backed-media` | M | 4.187 | 4.068 | 0.971 | 0.956-1.009 | paired range straddles 1.0 |
| DOCX | `docx_streaming_create` | `docx-streaming-paragraphs-medium` | S | 5.814 | 3.033 | 0.522 | 0.512-0.539 |  |
| DOCX | `docx_streaming_create` | `docx-streaming-paragraphs-large` | S | 91.48 | 47.35 | 0.518 | 0.507-0.528 |  |
| PPTX | `pptx_cross_copy_media_rich_lifecycle` | `pptx_cross_copy_media_rich_lifecycle` | M | 416.7 | 121.8 | 0.292 | 0.281-0.296 | final writes different bytes (33,599,745 vs 33,599,873; record 0742 transfers source-compressed image bytes) |
| PPTX | `pptx_cross_copy_plain_lifecycle` | `pptx_cross_copy_plain_lifecycle` | M | 8.321 | 8.087 | 0.972 | 0.971-0.979 |  |
| PPTX | `pptx_ordinary_save_lifecycle` | `pptx-semantic-medium` | M | 7.419 | 7.205 | 0.971 | 0.947-0.989 |  |
| PPTX | `pptx_semantic_full_text` | `pptx-semantic-medium` | S | 0.5895 | 0.3424 | 0.581 | 0.568-0.584 |  |
| PPTX | `pptx_semantic_full_text` | `pptx-semantic-large` | S | 50.59 | 27.36 | 0.541 | 0.524-0.544 |  |
| PPTX | `pptx_semantic_noop_edit_save` | `pptx-semantic-medium` | S | 0.8504 | 0.7994 | 0.940 | 0.932-0.949 |  |
| PPTX | `pptx_semantic_noop_edit_save` | `pptx-semantic-large` | S | 27.74 | 23.46 | 0.846 | 0.832-0.871 |  |
| PPTX | `pptx_semantic_one_edit_save` | `pptx-semantic-medium` | S | 1.711 | 1.5 | 0.877 | 0.867-0.881 |  |
| PPTX | `pptx_semantic_one_edit_save` | `pptx-semantic-large` | S | 55.59 | 45.03 | 0.810 | 0.793-0.844 |  |
| PPTX | `pptx_semantic_one_percent_edit_save` | `pptx-semantic-medium` | S | 1.711 | 1.491 | 0.871 | 0.860-0.877 |  |
| PPTX | `pptx_semantic_one_percent_edit_save` | `pptx-semantic-large` | S | 272.1 | 146 | 0.537 | 0.532-0.541 |  |
| PPTX | `pptx_semantic_open` | `pptx-semantic-medium` | S | 0.4485 | 0.4496 | 1.003 | 0.997-1.016 | paired range straddles 1.0 |
| PPTX | `pptx_semantic_open` | `pptx-semantic-large` | S | 2.03 | 2.142 | 1.055 | 1.004-1.093 |  |
| PPTX | `pptx_source_backed_cross_copy_media_rich_lifecycle` | `pptx_source_backed_cross_copy_media_rich_lifecycle` | M | 13.59 | 17.23 | 1.268 | 0.991-1.625 | per-process page-fault modes; equal per mode and no arm difference when run alone (supplementary M1 below); paired range straddles 1.0 |
| PPTX | `pptx_source_backed_one_edit_save` | `pptx-source-backed-media` | M | 6.823 | 6.845 | 1.003 | 1.002-1.009 |  |
| PPTX | `pptx_streaming_create` | `pptx-streaming-slides-medium` | S | 6.52 | 6.367 | 0.977 | 0.970-0.979 |  |
| PPTX | `pptx_streaming_create` | `pptx-streaming-slides-large` | S | 188 | 185.9 | 0.989 | 0.970-0.996 |  |
| XLSX | `xlsx_eager_cell_values_one_edit_save` | `xlsx-cell-values-medium` | M | 5.731 | 2.781 | 0.485 | 0.477-0.490 |  |
| XLSX | `xlsx_eager_cell_values_one_edit_save` | `xlsx-cell-values-dense-sparse` | M | 38.18 | 15.26 | 0.400 | 0.398-0.408 |  |
| XLSX | `xlsx_first_cell` | `xlsx-medium` | X | 0.4509 | 0.1333 | 0.296 | 0.292-0.298 |  |
| XLSX | `xlsx_first_cell` | `xlsx-dense-wide` | X | 28.42 | 7.794 | 0.274 | 0.271-0.276 |  |
| XLSX | `xlsx_full_cell_scan` | `xlsx-medium` | X | 0.4575 | 0.1354 | 0.296 | 0.294-0.299 |  |
| XLSX | `xlsx_full_cell_scan` | `xlsx-dense-wide` | X | 28.99 | 7.911 | 0.273 | 0.270-0.276 |  |
| XLSX | `xlsx_one_cell_commit_save` | `xlsx-medium` | X | 2.189 | 0.8861 | 0.405 | 0.400-0.407 |  |
| XLSX | `xlsx_one_cell_commit_save` | `xlsx-dense-wide` | X | 152.7 | 60.9 | 0.399 | 0.396-0.400 |  |
| XLSX | `xlsx_one_percent_commit_save` | `xlsx-medium` | X | 8.928 | 3.655 | 0.409 | 0.409-0.413 |  |
| XLSX | `xlsx_one_percent_commit_save` | `xlsx-dense-wide` | X | 308.2 | 123.2 | 0.400 | 0.388-0.403 |  |
| XLSX | `xlsx_open_owned` | `xlsx-medium` | X | 0.1068 | 0.1073 | 1.004 | 1.003-1.010 |  |
| XLSX | `xlsx_open_owned` | `xlsx-dense-wide` | X | 1.355 | 1.353 | 0.998 | 0.997-1.036 | paired range straddles 1.0 |
| XLSX | `xlsx_ordinary_save_lifecycle` | `xlsx-cell-values-medium` | M | 16.95 | 14.51 | 0.856 | 0.853-0.897 |  |
| XLSX | `xlsx_source_backed_cell_values_one_edit_save` | `xlsx-cell-values-medium` | M | 4.48 | 4.163 | 0.929 | 0.923-0.956 |  |
| XLSX | `xlsx_source_backed_cell_values_one_edit_save` | `xlsx-cell-values-dense-sparse` | M | 29.24 | 26.96 | 0.922 | 0.913-0.941 |  |
| XLSX | `xlsx_streaming_create` | `xlsx-streaming-create-tiny` | X | 0.09399 | 0.08997 | 0.957 | 0.947-0.968 |  |
| XLSX | `xlsx_streaming_create` | `xlsx-streaming-create-medium` | X | 10.83 | 10.47 | 0.967 | 0.959-0.969 |  |
| XLSX | `xlsx_streaming_create` | `xlsx-streaming-create-large` | X | 166.4 | 160.7 | 0.966 | 0.953-0.972 |  |
| DOC | `doc_fresh_write_to` | `doc-tiny` | W | 0.007935 | 0.00729 | 0.919 | 0.906-0.951 | sub-10us per sample |
| DOC | `doc_fresh_write_to` | `doc-large` | W | 0.24 | 0.1831 | 0.763 | 0.758-0.777 |  |
| DOC | `doc_fresh_write_to` | `doc-payload-heavy` | W | 5.918 | 1.899 | 0.321 | 0.319-0.323 |  |
| DOC | `doc_semantic_full_text` | `doc-tiny` | S | 3e-05 | 3e-05 | 1.000 | 1.000-1.000 | sub-10us per sample |
| DOC | `doc_semantic_full_text` | `doc-large` | S | 0.000505 | 0.000575 | 1.139 | 1.096-1.208 | sub-10us per sample |
| DOC | `doc_semantic_one_edit_save` | `doc-tiny` | S | 0.03937 | 0.03865 | 0.982 | 0.958-0.998 |  |
| DOC | `doc_semantic_one_edit_save` | `doc-large` | S | 0.6218 | 0.6095 | 0.980 | 0.975-0.992 |  |
| DOC | `doc_semantic_open` | `doc-tiny` | S | 0.006165 | 0.00624 | 1.012 | 0.987-1.030 | paired range straddles 1.0; sub-10us per sample |
| DOC | `doc_semantic_open` | `doc-large` | S | 0.2807 | 0.3015 | 1.074 | 1.028-1.106 |  |
| PPT | `ppt_fresh_write_to` | `ppt-tiny` | W | 0.01391 | 0.01221 | 0.878 | 0.865-0.912 |  |
| PPT | `ppt_fresh_write_to` | `ppt-large` | W | 0.1371 | 0.07193 | 0.525 | 0.516-0.532 |  |
| PPT | `ppt_fresh_write_to` | `ppt-payload-heavy` | W | 4.874 | 2.352 | 0.483 | 0.480-0.486 |  |
| PPT | `ppt_semantic_full_text` | `ppt-tiny` | S | 0.00199 | 0.002 | 1.005 | 0.990-1.026 | paired range straddles 1.0; sub-10us per sample |
| PPT | `ppt_semantic_full_text` | `ppt-large` | S | 0.0265 | 0.02689 | 1.015 | 1.007-1.023 |  |
| PPT | `ppt_semantic_one_edit_save` | `ppt-tiny` | S | 0.09049 | 0.08085 | 0.893 | 0.889-0.895 |  |
| PPT | `ppt_semantic_one_edit_save` | `ppt-large` | S | 0.2299 | 0.183 | 0.796 | 0.795-0.806 |  |
| PPT | `ppt_semantic_open` | `ppt-tiny` | S | 0.006455 | 0.006165 | 0.955 | 0.928-0.979 | sub-10us per sample |
| PPT | `ppt_semantic_open` | `ppt-large` | S | 0.01334 | 0.01341 | 1.005 | 0.998-1.020 | paired range straddles 1.0 |
| XLS | `xls_comments_eager_edit_save` | `xls-comments-opaque-heavy` | M | 27.38 | 18.84 | 0.688 | 0.674-0.724 |  |
| XLS | `xls_fresh_write_to` | `xls-tiny` | W | 0.00555 | 0.005395 | 0.972 | 0.948-0.985 | sub-10us per sample |
| XLS | `xls_fresh_write_to` | `xls-large` | W | 1.223 | 1.195 | 0.977 | 0.976-0.986 |  |
| XLS | `xls_fresh_write_to` | `xls-payload-heavy` | W | 3.903 | 1.363 | 0.349 | 0.347-0.350 |  |
| XLS | `xls_numeric_eager_rk_mulrk_edit_save` | `xls-rk-mulrk-deterministic` | M | 2.379 | 0.2879 | 0.121 | 0.120-0.122 |  |
| XLS | `xls_semantic_full_cell_scan` | `xls-tiny` | S | 0.00019 | 0.0002 | 1.053 | 1.000-1.105 | sub-10us per sample |
| XLS | `xls_semantic_full_cell_scan` | `xls-large` | S | 0.07218 | 0.07363 | 1.020 | 1.017-1.021 |  |
| XLS | `xls_semantic_one_edit_save` | `xls-tiny` | S | 0.08118 | 0.03057 | 0.377 | 0.372-0.377 |  |
| XLS | `xls_semantic_one_edit_save` | `xls-large` | S | 3.337 | 0.7648 | 0.229 | 0.227-0.231 |  |
| XLS | `xls_semantic_open` | `xls-tiny` | S | 0.01115 | 0.01078 | 0.967 | 0.948-0.974 |  |
| XLS | `xls_semantic_open` | `xls-large` | S | 1.396 | 1.3 | 0.931 | 0.920-0.979 |  |
| XLS | `xls_visibility_eager_edit_save` | `xls-visibility-opaque` | M | 24.95 | 2.786 | 0.112 | 0.111-0.115 |  |
| OLE2 common | `ole_common_one_edit_save` | `ole-common-tiny-compressible` | M | 0.02331 | 0.02354 | 1.010 | 0.986-1.018 | paired range straddles 1.0 |
| OLE2 common | `ole_common_one_edit_save` | `ole-common-tiny-incompressible` | M | 0.02297 | 0.02323 | 1.011 | 0.994-1.024 | paired range straddles 1.0 |
| OLE2 common | `ole_common_one_edit_save` | `ole-common-many-small-compressible` | M | 1.423 | 1.402 | 0.985 | 0.980-1.017 | paired range straddles 1.0 |
| OLE2 common | `ole_common_one_edit_save` | `ole-common-many-small-incompressible` | M | 1.452 | 1.39 | 0.957 | 0.942-0.999 |  |
| OLE2 common | `ole_common_one_edit_save` | `ole-common-few-large-compressible` | M | 21.35 | 21.44 | 1.004 | 0.995-1.041 | paired range straddles 1.0 |
| OLE2 common | `ole_common_one_edit_save` | `ole-common-few-large-incompressible` | M | 21.34 | 21.54 | 1.010 | 0.992-1.060 | paired range straddles 1.0 |
| OLE2 common | `ole_common_one_edit_save` | `ole-common-wide-root-compressible` | M | 14.18 | 13.9 | 0.980 | 0.951-1.001 | paired range straddles 1.0 |
| OLE2 common | `ole_common_one_edit_save` | `ole-common-wide-root-incompressible` | M | 14.19 | 13.85 | 0.976 | 0.929-1.007 | paired range straddles 1.0 |

## Unweighted geometric mean of row ratios (descriptive)

| Format | Rows | Geomean ratio |
|---|---:|---:|
| DOCX | 14 | 0.675 |
| PPTX | 17 | 0.813 |
| XLSX | 18 | 0.551 |
| DOC | 9 | 0.864 |
| PPT | 9 | 0.813 |
| XLS | 12 | 0.502 |
| OLE2 common | 8 | 0.991 |

| Group | Flags | Rows | Geomean ratio |
|---|---|---:|---:|
| S | `--semantic-shape medium,large` | 42 | 0.777 |
| X | `--xlsx-shape medium,dense-wide` | 13 | 0.509 |
| W | `--writer-shape tiny,large,payload-heavy` | 9 | 0.634 |
| M | `(default shape flags)` | 23 | 0.713 |

## Supplementary: `pptx_source_backed_cross_copy_media_rich_lifecycle` by process

Each process settles into one page-fault mode for this case and keeps it for all nine samples. Every mode seen in both arms gives the same time in both. The 12,268-fault mode (19.4-19.9 ms) appeared only in two final processes in group M, where this case runs directly after `pptx_cross_copy_media_rich_lifecycle`, whose allocation pattern record 0742 changed. M1 re-runs the case alone (default flags, same A B B A A B B A protocol): both arms show the same two modes at the same times, so the M-group ratio of 1.27 reflects which mode each process landed in rather than a slower code path.

| Run | Pos | Arm | p50 (ms) | Publication p50 (ms) | Plan p50 (ms) | Minor faults / sample |
|---|---:|---|---:|---:|---:|---:|
| M (in group) | 1 | base | 12.20 | 7.39 | 4.16 | 0 |
| M (in group) | 2 | final | 15.04 | 7.74 | 6.64 | 4,577 |
| M (in group) | 3 | final | 14.90 | 7.66 | 6.57 | 4,577 |
| M (in group) | 4 | base | 15.04 | 7.72 | 6.65 | 4,577 |
| M (in group) | 5 | base | 14.93 | 7.67 | 6.61 | 4,577 |
| M (in group) | 6 | final | 19.42 | 12.59 | 6.17 | 12,268 |
| M (in group) | 7 | final | 19.89 | 12.79 | 6.40 | 12,268 |
| M (in group) | 8 | base | 12.24 | 7.41 | 4.20 | 0 |
| M1 (alone) | 1 | base | 17.21 | 12.41 | 4.17 | 8,203 |
| M1 (alone) | 2 | final | 17.27 | 12.42 | 4.17 | 8,203 |
| M1 (alone) | 3 | final | 12.12 | 7.37 | 4.13 | 0 |
| M1 (alone) | 4 | base | 17.01 | 12.25 | 4.16 | 8,203 |
| M1 (alone) | 5 | base | 17.06 | 12.27 | 4.15 | 8,203 |
| M1 (alone) | 6 | final | 17.18 | 12.37 | 4.16 | 8,203 |
| M1 (alone) | 7 | final | 12.20 | 7.40 | 4.14 | 0 |
| M1 (alone) | 8 | base | 17.02 | 12.36 | 4.17 | 8,203 |

## Whole-process instruction counts (`perf stat -e instructions,instructions:u -x,`)

One extra process per arm per group, run after the timed rounds with the same flags (its JSON is kept as `raw/<G>-perfstat-<arm>.json.gz` and is not used in the timing table). Counts cover the whole process, including corpus generation and verification outside the timed regions.

| Group | Base instructions | Final instructions | Ratio | Base instructions:u | Final instructions:u | Ratio (:u) |
|---|---:|---:|---:|---:|---:|---:|
| S | 877,563,291,649 | 764,903,799,685 | 0.872 | 875,657,664,721 | 762,933,812,891 | 0.871 |
| X | 184,598,473,200 | 105,728,013,257 | 0.573 | 182,681,658,596 | 103,850,807,020 | 0.568 |
| W | 6,287,521,630 | 3,635,941,661 | 0.578 | 4,543,637,004 | 1,920,961,134 | 0.423 |
| M | 231,486,423,269 | 149,867,376,761 | 0.647 | 213,821,992,033 | 130,425,834,113 | 0.610 |
