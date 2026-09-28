# 0829 PPTX edit profile

Offline replay of the pinned real-file public PPTX edit transaction.

Observed reports/samples: 23 / 4549.
Planned reports/samples: 23 / 4549.
Native values are descriptive perturbation evidence; no shipping latency, historical speedup, adoption, or Amdahl fraction is claimed.

| Arm | p50 median (ns) | p95 median (ns) | p99 median (ns) | mean median (ns) | RSS median (KiB) | tail flag |
|---|---:|---:|---:|---:|---:|---|
| control | 1.27095e+06 | 1.28511e+06 | 1.29784e+06 | 1.27347e+06 | 4976 | no |
| wrapped | 1.27147e+06 | 1.28729e+06 | 1.29953e+06 | 1.27403e+06 | 5008 | no |
| fp | 1.30278e+06 | 1.3148e+06 | 1.31776e+06 | 1.30423e+06 | 4994 | no |

| Paired diagnostic | median ratio | CI95 |
|---|---:|---|
| wrapped/control | 1.00071 | [0.996008, 1.00398] |
| fp/wrapped | 1.02375 | [1.02022, 1.02728] |

Perf status: available.
Perf frames retain whole-process periods, headers, lost/status/unknown diagnostics, and exact-owner leaf/callpath summaries.

The reader never invokes Cargo, the probe, perf, nm, or objdump.
