Paired scaling evidence for measured cf98 snapshot
=================================================

The archive retains two worker-width captures from source cf98bee37455e8e6d0e73002ff2e94299ef546cb, plus the first capture's simulated range lane. It contains ten reports, ten catalogs, exact runners, source/build/host receipts, analyses and a portable verifier. The benchmark binary is external; its SHA-256 and locked build recipe are retained. This is historical measured source, not a timing claim for the publication commit.

Root and independent opc_capture_review validation passed. The verifier executes both validators afresh and recomputes first/paired work vectors and p50 comparisons. All 16 scaling case/corpus/width combinations have matching work across captures. Superlinear observations and shared-host contention remain explicit; these results do not establish an Amdahl parallel fraction, a causal optimization benefit, or production-wide scaling.

To inspect the full report and reproduce verification:

```sh
tar -xzf publication-bundle.tar.gz
cd current-head-scaling-20260910
PYTHONDONTWRITEBYTECODE=1 python3 scripts/verify.py
```

The inner README documents optional `--binary PATH` verification. Root also verified the recorded external binary bytes. The archive has 104 regular files: 103 sealed artifacts and their hash receipt.
