# XLSB Custom Data lifecycle performance profile

This directory contains candidate-only characterization for the committed
`litchi_xlsb::custom_data` owner. The measurement plan is in [PLAN.md](PLAN.md).
The owner design and lifecycle correctness evidence remain the authorities for
scope and semantics; this profile does not add native-producer or Office
acceptance evidence.

Run the clean profile from the repository root with:

```sh
XLSB_CUSTOM_DATA_RESULT_DIR=docs/report/spec-gap-validation-evidence/xlsb-custom-data-performance/receipts/final-release \
  bash docs/report/spec-gap-validation-evidence/xlsb-custom-data-performance/run_profile.sh
```

The runner checks out source commit `16102fe751d7c5492042330f1bd1f49c304495f0`
into a sparse detached worktree, overlays only the standalone harness, builds
an optimized `--release` binary with the retained harness lockfile, copies the
binary into the receipt, and runs all 24 lane/class groups in fresh processes. It retains source and
harness hashes, toolchain, binary hash, deterministic fixture hashes, five raw
samples per group, `/usr/bin/time -v` sidecars, and exact commands. The source
checkout and the receipt directory are left for independent review; the
disposable Cargo target is removed after the binary is retained. The copied
receipt binaries are review-time artifacts; retain them through root review but
do not add them to the repository commit. The recorded hash and clean build
command are the reproducibility boundary. The recorded source-root path is a
disposable replay hint; the source manifest uses stable profile-relative labels,
and replay reconstructs the pinned commit if that path is unavailable.

The verifier can be rerun directly:

```sh
python3 docs/report/spec-gap-validation-evidence/xlsb-custom-data-performance/verify.py \
  --results docs/report/spec-gap-validation-evidence/xlsb-custom-data-performance/receipts/final-release
```

Replay uses the retained clean source checkout recorded by the first receipt
when available. If that disposable checkout has been removed, it reconstructs
the exact pinned commit from the repository before rebuilding:

```sh
bash docs/report/spec-gap-validation-evidence/xlsb-custom-data-performance/replay.sh \
  docs/report/spec-gap-validation-evidence/xlsb-custom-data-performance/receipts/final-release
```

To render the descriptive medians from any verified receipt:

```sh
python3 docs/report/spec-gap-validation-evidence/xlsb-custom-data-performance/summarize.py \
  --results docs/report/spec-gap-validation-evidence/xlsb-custom-data-performance/receipts/final-release \
  --output docs/report/spec-gap-validation-evidence/xlsb-custom-data-performance/report.md
```

Elapsed times include the process-local allocator observer's atomic accounting
overhead. Five-sample medians are descriptive observations, not tail-latency
or uncertainty certification. Logical peak is aggregate allocator accounting
and does not expose transient overlap hidden by a reallocating allocator. RSS
is retained separately from `/usr/bin/time -v`. Copy accounting covers only
explicit post-timer harness validation copies; internal production copy traffic
is uninstrumented.

Root independently rebuilt and replayed the final harness: [verification](root-replay-verification.json) and [raw records](root-replay/raw/). All 120 samples match the canonical run for semantic assertions, fixture identity, candidate/forward package hashes, and observed copy counts. Fifteen mixed-insert/remove samples vary in allocator totals; these totals are observations, not deterministic guarantees. The retained binary hash is the rebuild identity; executable copies are excluded from version control.
