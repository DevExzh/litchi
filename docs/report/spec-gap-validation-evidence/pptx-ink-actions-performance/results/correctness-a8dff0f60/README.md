# PPTX InkAction correctness capture

This directory retains the timing-free correctness gate run from clean
checkout `a8dff0f606ea671ad2d648c80e6ec1a0d3a7ba80`. The checkout was detached
at that full commit and was clean before and after capture. The semantic owner
pin is `cf6fdb8e91dd232d7d762596763d2e9d8a5b9dbd`; the production source
baseline pin is `2a2ffa1cae4e6b7070082768ce84483e5d411dc8`.

The source worktree used for the run was
`/var/tmp/pptx-ink-actions-correctness-a8dff0f60-src.Zk6bXL`. The external
capture root was
`/var/tmp/pptx-ink-actions-correctness-a8dff0f60.l7raBv`. The copied `capture/`
directory contains all 123 operator result files, including command exits,
host metadata, Cargo metadata, source manifests, binary hashes, host probe,
matrix receipt, and all 42 individual lane receipts. It contains 5,235,708
bytes. `regression/` retains the serial six-test Rust regression stdout,
stderr, and status receipt (5,208 bytes). `operator/` retains the operator
stdout, stderr, status receipt, and detached-worktree setup log.

The serial regression command was run with a separate external target:

```text
env CARGO_TARGET_DIR=/var/tmp/pptx-ink-actions-correctness-a8dff0f60.l7raBv/regression-target-parent/target CARGO_INCREMENTAL=0 LC_ALL=C cargo test --release --locked --offline --manifest-path docs/report/spec-gap-validation-evidence/pptx-ink-actions-performance/harness/Cargo.toml -- --test-threads=1
```

It reported 6 passed and 0 failed. The fail-closed correctness operator then
ran with fresh absent result and target paths:

```text
python3 operator_capture.py --expected-head a8dff0f606ea671ad2d648c80e6ec1a0d3a7ba80 /var/tmp/pptx-ink-actions-correctness-a8dff0f60-src.Zk6bXL /var/tmp/pptx-ink-actions-correctness-a8dff0f60-src.Zk6bXL/docs/report/spec-gap-validation-evidence/pptx-ink-actions-performance /var/tmp/pptx-ink-actions-correctness-a8dff0f60.l7raBv/capture-results/run /var/tmp/pptx-ink-actions-correctness-a8dff0f60.l7raBv/capture-target-parent/target
```

The operator completed preflight, release build, host probe, matrix, and all
42 individual lanes with exit 0. The matrix accepted 42 of 42 lanes, with 9
expected refusal lanes and no failed lanes. The retained receipt audit in
`correctness-verification.json` found no failures for typed errors and limits,
refusal source preservation, the five acceptance predicates, topology
observations, phase allocation balances, source and metadata stability, or
command provenance.

This is a correctness capture only. It used no `/usr/bin/time` wrapper and
makes no latency, allocation-performance, native PowerPoint, or speedup claim.
Root verified all 138 retained files against Git bytes in `8343e7489`,
checked the binary hash and both target inventories, then removed the owned
external receipts, worktree setup log, and compiled targets (710,158,528
bytes). The clean source worktree remains. A timing capture must use fresh
results and targets; none of the deleted build artifacts is required for
replaying the retained receipt audit.

Key provenance values:

- Binary SHA-256: `a47055e7e28e6a2d96b7fb89ba6a7a5e6c51c9a8c842c54ce2461bd633a8d123`
- Source manifest before/after SHA-256: `e1777b98800a21afb090dd713f5d5a9c7cdc15e2d6bc720f75b99ad075d03e6c`
- Isolated harness lock SHA-256: `3c773b848134c52df596034ffbbe8688dde9924fb741dee8bc8bce0eaa772f2e`
- Fixture helper SHA-256: `bec6baafcf735d778216fb54fe6299312707e912dcaed5f89de988f7112bb58e`
- Retained adapter generator SHA-256: `eae2c7b98970878bfcdc743e0e92b13f418fbf4a21243fac698b13f01bbd7a72`
- Receipt audit SHA-256: `c242e37819f8b6f9f761ca60b69b991c92a1c801464f1b8571e6622e7ccf80a6`
- External retention manifest SHA-256: `2268719babeccb73f46f2a0ad7f7c51aa8cdc31acf9d589727277d4cc2dc388c`
