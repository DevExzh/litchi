# 0827 spread and resource diagnostics

Derived from the independently checked `analysis.json`. A spread is maximum
divided by minimum across six native blocks or two observer blocks. Only
values above 1.05 appear below. Observer elapsed statistics are diagnostics.

| Lane | Format / phase | Leg | Flagged block spread ratios | Tail p99/p50 flag |
| --- | --- | --- | --- | --- |
| native | DOCX / lifecycle | before | none | none |
| native | DOCX / lifecycle | after | p99=1.345546 | none |
| native | DOCX / edit | before | p95=1.191109, p99=1.156412 | 1.317326 |
| native | DOCX / edit | after | p95=1.156447, p99=1.232175 | 1.217804 |
| native | DOCX / atomic_publish | before | none | none |
| native | DOCX / atomic_publish | after | p99=1.059025 | none |
| native | DOCX / counting_publish | before | p95=1.155465, p99=1.765231 | 1.334080 |
| native | DOCX / counting_publish | after | p95=1.109491, p99=1.607411 | 1.354139 |
| native | XLSX / lifecycle | before | p99=1.064360 | none |
| native | XLSX / lifecycle | after | p99=1.094086 | 1.062673 |
| native | XLSX / edit | before | p50=1.053348, p95=1.819188, p99=1.882664, mean=1.164046 | 1.172355 |
| native | XLSX / edit | after | p95=1.676771, p99=1.872303, mean=1.090731 | 1.736931 |
| native | XLSX / atomic_publish | before | p99=1.271258 | none |
| native | XLSX / atomic_publish | after | none | none |
| native | XLSX / counting_publish | before | p95=1.493752, p99=1.795580, mean=1.114437 | 1.843942 |
| native | XLSX / counting_publish | after | p95=1.318641, p99=1.174979, mean=1.092389 | 1.699932 |
| native | PPTX / lifecycle | before | p99=1.053233 | none |
| native | PPTX / lifecycle | after | p99=1.227302 | none |
| native | PPTX / edit | before | p99=1.067628 | none |
| native | PPTX / edit | after | none | none |
| native | PPTX / atomic_publish | before | none | none |
| native | PPTX / atomic_publish | after | none | none |
| native | PPTX / counting_publish | before | p99=1.085369 | 1.064235 |
| native | PPTX / counting_publish | after | none | 1.054891 |
| observer | DOCX / lifecycle | before | none | none |
| observer | DOCX / lifecycle | after | none | none |
| observer | DOCX / edit | before | none | 1.245412 |
| observer | DOCX / edit | after | p50=1.072046 | 1.149544 |
| observer | DOCX / atomic_publish | before | none | none |
| observer | DOCX / atomic_publish | after | none | none |
| observer | DOCX / counting_publish | before | none | 1.661242 |
| observer | DOCX / counting_publish | after | none | 1.693473 |
| observer | XLSX / lifecycle | before | none | none |
| observer | XLSX / lifecycle | after | none | none |
| observer | XLSX / edit | before | p50=1.281996, p95=1.055944, p99=1.055944, mean=1.177138 | 1.124098 |
| observer | XLSX / edit | after | p95=1.229269, p99=1.229269, mean=1.090022 | 1.206222 |
| observer | XLSX / atomic_publish | before | none | none |
| observer | XLSX / atomic_publish | after | none | none |
| observer | XLSX / counting_publish | before | p50=1.268970, p95=1.195606, p99=1.195606, mean=1.166871 | 1.222918 |
| observer | XLSX / counting_publish | after | p50=1.266821, p95=1.165007, p99=1.165007 | 1.430593 |
| observer | PPTX / lifecycle | before | none | none |
| observer | PPTX / lifecycle | after | none | none |
| observer | PPTX / edit | before | none | none |
| observer | PPTX / edit | after | none | none |
| observer | PPTX / atomic_publish | before | none | none |
| observer | PPTX / atomic_publish | after | none | none |
| observer | PPTX / counting_publish | before | none | none |
| observer | PPTX / counting_publish | after | p50=1.057605 | 1.063811 |

All RSS and positive allocation counter block spreads are at most 1.05.
The complete raw and derived allocation vectors, RSS values, process counter
vectors, and every unflagged spread ratio are retained in `analysis.json`.

| Format / phase | Native paired RSS ratio | Observer paired RSS ratio |
| --- | ---: | ---: |
| DOCX / lifecycle | 0.999985 | 1.000803 |
| DOCX / edit | 1.000831 | 0.999798 |
| DOCX / atomic_publish | 1.000627 | 1.000467 |
| DOCX / counting_publish | 0.999737 | 0.999708 |
| XLSX / lifecycle | 1.000686 | 1.000949 |
| XLSX / edit | 0.999738 | 0.999694 |
| XLSX / atomic_publish | 0.999985 | 0.999679 |
| XLSX / counting_publish | 1.000000 | 0.999533 |
| PPTX / lifecycle | 1.000117 | 0.999373 |
| PPTX / edit | 1.000029 | 1.000467 |
| PPTX / atomic_publish | 1.000160 | 1.000190 |
| PPTX / counting_publish | 1.000233 | 1.000059 |
