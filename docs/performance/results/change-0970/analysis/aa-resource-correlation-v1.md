# 0969 selected-cell resource correlation addendum

This is a bounded descriptive analysis of retained 0969 A/A packet metadata. It launches no workload, binary, build, test, or reducer and makes no causal or host claim.

Input custody: **PASS**. The selected set contains 96 children and 7680 elapsed samples across four cells; the raw packet set is bound by `3eecb0db82fb14be38852685956dd6fa485fad2c609b23261fa5c75443befdd8`.

Per-cell ranges and correlations (r is Pearson; rho is Spearman):

| cell | children/samples | p50 range (ns) | RSS range (KiB) | whole-child wall range (ns) | output capacity | p50↔RSS r/rho | p50↔wall r/rho |
|---|---:|---:|---:|---:|---:|---:|---:|
| `docfloat/e1/default` | 24/1920 | 972389–1244070 | 13204–14268 | 614373353–664580903 | 342528 | -0.336/-0.420 | 0.578/0.424 |
| `docfloat/e32/default` | 24/1920 | 7373035–8130358 | 16420–17352 | 1265320941–1415700464 | 686592 | -0.400/-0.286 | 0.817/0.554 |
| `docnohf/e1/default` | 24/1920 | 114316–118761 | 7308–7608 | 113604843–113671722 | 27648 | 0.109/0.047 | 0.156/0.274 |
| `docnohf/e32/default` | 24/1920 | 999140–1030384 | 7556–7840 | 163688653–163842153 | 63488 | -0.190/-0.162 | 0.456/0.272 |

The pooled p50 correlations are RSS r/rho 0.821/0.834, whole-child wall 0.943/0.895, and output capacity 0.917/0.888. These pooled values include cell composition and are not process-level causal evidence.

Within each cell, output capacity is constant, so its p50 correlation is undefined; the four cell capacities differ and explain the pooled association. The fixed six-child fastest/slowest p50 quartiles show whole-child wall medians for the slow tail above the fast tail in all four cells, while RSS medians are equal or lower in the slow tail. This is an association in retained metadata, without a selected value threshold or causal interpretation.

Fixed-tail quartile medians (six children per tail; 480 elapsed samples per tail):

| cell | fast p50 / RSS / wall | slow p50 / RSS / wall |
|---|---:|---:|
| `docfloat/e1/default` | 977456 ns / 13580 KiB / 614425848 ns | 1219530 ns / 13382 KiB / 639502033 ns |
| `docfloat/e32/default` | 7855850 ns / 16804 KiB / 1365536424 ns | 8094420 ns / 16674 KiB / 1415547472 ns |
| `docnohf/e1/default` | 115141 ns / 7374 KiB / 113637242 ns | 117486 ns / 7372 KiB / 113644798 ns |
| `docnohf/e32/default` | 1000537 ns / 7746 KiB / 163717048 ns | 1025562 ns / 7628 KiB / 163752268 ns |

Analyzer: `/home/zhuhe/litchi-performance-0970/tools/analyze_aa_resource_correlation_v1.py` (SHA-256 `5e5c0dcdcbc728777fd9b04598332584a4a7140d633347ce2abd4513a59f5e83`). JSON: `/home/zhuhe/litchi-performance-0970/analysis/aa-resource-correlation-v1.json` (SHA-256 `2ee8093d29f1443eb56e38fc0cb6fb690cbb0aa97a3681ed7c9558d61d1798e0`).
