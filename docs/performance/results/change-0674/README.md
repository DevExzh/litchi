# Change 0674 evidence packet

This packet closes the remaining harness, gate and reproducibility items from
rows 17, 19 and 20 of change 0651 on the 0664 head. It contains no production
crate change and makes no performance claim.

| Path | What it is |
| --- | --- |
| [`../../0674-performance-gate-hygiene.md`](../../0674-performance-gate-hygiene.md) | the retained change record |
| [`decision.json`](decision.json) | the disposition and accepted evidence |
| [`cleanup.json`](cleanup.json) | owned build and scratch paths to remove after merge |
| [`gates.log`](gates.log) | command outcomes for structural and strict claims validation |
| [`log-sections.md`](log-sections.md) | the four paragraphs for the coordinator's rollup |

## Replay the focused checks

From the repository root, with Rust 1.95.0:

```sh
cargo +1.95.0 fmt --manifest-path tools/perf-baseline/Cargo.toml --all -- --check
CARGO_TARGET_DIR=/home/zhuhe/code/litchi-targets/0674-harness \
  CARGO_BUILD_JOBS=2 CARGO_PROFILE_DEV_DEBUG=0 \
  cargo clippy --manifest-path tools/perf-baseline/Cargo.toml --lib --tests -- -D warnings
CARGO_TARGET_DIR=/home/zhuhe/code/litchi-targets/0674-harness \
  CARGO_BUILD_JOBS=2 CARGO_PROFILE_DEV_DEBUG=0 \
  cargo test --manifest-path tools/perf-baseline/Cargo.toml \
  --lib semantic_docx_and_pptx_tiny_corpora_are_deterministic_and_editable \
  -- --test-threads=1
CARGO_TARGET_DIR=/home/zhuhe/code/litchi-targets/0674-harness \
  cargo test --manifest-path tools/perf-baseline/Cargo.toml \
  --lib xlsb_lifecycle_selectors_are_opt_in_and_hash_stable -- --test-threads=1
CARGO_TARGET_DIR=/home/zhuhe/code/litchi-targets/0674-harness \
  cargo test --manifest-path tools/perf-baseline/Cargo.toml \
  --lib selectable_case_count_matches_current_enumeration -- --test-threads=1
python3 tools/check_perf_claims.py \
  --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural
python3 tools/check_perf_claims.py \
  --registry docs/performance/claim-registry-v1.json --repo-root . \
  --evidence-root . --mode strict
python3 -m unittest tools.test_non_iwork_gate tools.test_perf_baseline_source_policy \
  tools.test_perf_claims \
  tools.test_corpus_manifest_v2 tools.test_crud_coverage_index \
  tools.test_perf_corpus_binding
```

The full command transcript is in [`gates.log`](gates.log). Structural mode
checks the registry shape and repository references without opening retained
evidence, while the named workflow gate runs strict mode with
`--evidence-root .` and independently verifies the retained evidence.
