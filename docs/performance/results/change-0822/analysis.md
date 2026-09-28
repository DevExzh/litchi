# 0822 PPTX edit profile

Offline replay of the pinned real-file public PPTX edit transaction.

Observed reports/samples: 23 / 4549.
Planned reports/samples: 23 / 4549.
Native values are descriptive perturbation evidence; no shipping latency, historical speedup, adoption, or Amdahl fraction is claimed.

| Arm | p50 median (ns) | p95 median (ns) | p99 median (ns) | mean median (ns) | RSS median (KiB) | tail flag |
|---|---:|---:|---:|---:|---:|---|
| control | 1.42844e+06 | 1.43887e+06 | 1.45258e+06 | 1.42877e+06 | 5122 | no |
| wrapped | 1.42819e+06 | 1.44195e+06 | 1.44818e+06 | 1.4294e+06 | 4940 | no |
| fp | 1.44308e+06 | 1.45252e+06 | 1.459e+06 | 1.44399e+06 | 5142 | no |

| Paired diagnostic | median ratio | CI95 |
|---|---:|---|
| wrapped/control | 1.00014 | [0.997252, 1.00243] |
| fp/wrapped | 1.00909 | [1.00581, 1.01267] |

Perf status: available.
Perf frames retain whole-process periods, headers, lost/status/unknown diagnostics, and exact-owner leaf/callpath summaries.

The reader never invokes Cargo, the probe, perf, nm, or objdump.
