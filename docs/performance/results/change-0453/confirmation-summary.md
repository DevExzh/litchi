# Separate plain-copy tail investigation

Eight fresh processes, two fixed ABBA blocks, three warmups and thirty samples.
Primary matrix remains unchanged. Combined quantiles are descriptive; these
sample counts cannot establish a stable population p99.

| Lane | Build | p50 ms | p95 ms | p99 ms | Maximum ms |
|---|---|---:|---:|---:|---:|
| 0 | baseline | 2.214840 | 2.227011 | 2.235315 | 2.238319 |
| 1 | candidate | 2.223940 | 2.240852 | 2.270132 | 2.280630 |
| 2 | candidate | 2.224875 | 2.240345 | 2.259728 | 2.267480 |
| 3 | baseline | 2.224320 | 2.235841 | 2.239274 | 2.239440 |
| 4 | baseline | 2.218739 | 2.237399 | 2.238839 | 2.238880 |
| 5 | candidate | 2.219229 | 2.233877 | 2.240592 | 2.242300 |
| 6 | candidate | 2.223260 | 2.233530 | 2.234889 | 2.235060 |
| 7 | baseline | 2.223530 | 2.233571 | 2.234616 | 2.234799 |

| ABBA block | p50 change | p95 change | p99 change |
|---|---:|---:|---:|
| 1 | +0.239% | +0.380% | +1.508% |
| 2 | +0.087% | -0.016% | +0.001% |

Pooled descriptive changes: {"api_p50_ms": 0.13238367170471144, "api_p95_ms": 0.12222312931882484, "api_p99_ms": 1.0819075882525424}.
