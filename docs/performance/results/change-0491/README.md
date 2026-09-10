# 0491: DOCX provider and cache-state baseline

This bundle measures one deterministic, synthetic DOCX full-text read lifecycle.
It expands the baseline; it is not an optimization comparison or completion of
the non-iWork performance goal. See [results-review.md](results-review.md),
[methods.md](methods.md), and [review-findings.md](review-findings.md).

## Evidence and scope

- `protocol.json` and `cold-protocol.json` freeze helpers, corpus, machine,
  arguments, roles, repeats, timing and sample counts before formal collection.
- `build-normal.json` and `build-allocator.json` identify final5 executables and
  their 7,158-file source manifest, exact release commands and terminal gates.
- `captures/provider-formal1/`: 20 processes, 600 samples, five provider arms.
- `captures/cold-formal1/`: 12 measured parents, 360 fresh-child samples, and
  four explicit prepared-query cold-ineligible controls.
- `analysis/` and `verification/` retain canonical raw-report recomputations.
- `profile-providers/profiles2/` contains whole-child perf/strace diagnostics.
  These include setup/corpus/report work, not operation-only attribution. See
  [profile-review.md](profile-review.md) for counter interpretation.
- `fincore-tool.json` binds the independently inspected fincore executable.
  The Rust report retains parsed residency/dirty/writeback observations and
  stderr hashes. Raw fincore stdout is not retained separately.
- `cleanup.json` and `seal.json` bind final custody and owned-scratch removal.
- `development/` and earlier pilot/build gates retain superseded diagnostics.
  They are not formal evidence and some intentionally record failures.

The exact original 0188 archive remains pinned. A narrow generator-only
compatibility restoration handles the one known producer namespace change;
`corpus-restoration-proof.json` proves all member bytes/order/CRCs. The aligned
copy's exact transform and raw EOCD-tail overlap remain explicit. The actual
timed text is checked after its versioned timer, rather than replaced by a
fresh successful read. Genuine borrowed input remains unimplemented; see
[borrowed-source-next.md](borrowed-source-next.md).

## Reproduce

Run from repository root with Rust 1.98.1 and the environment in `support.py`.
The retained environment uses Linux ext4, CPU 2, four Cargo jobs, no incremental
compilation, release debuginfo and frame pointers. A new machine or source needs
a new evidence directory, machine/protocol bindings and exclusive attempt names;
do not overwrite the retained bundle. The CPU lock is owned by `gate.py`.

The build commands are:

```sh
cargo build --release --locked --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline
cargo build --release --locked --manifest-path tools/perf-baseline/Cargo.toml --features allocator-metrics --bin litchi-perf-baseline-alloc
```

Use `gate.py LABEL COMMAND ...` for each build, then
`retain_build.py ROLE ATTEMPT GATE_LABEL` to retain binaries. Freeze provider
protocol using `provider_matrix.protocol_value(provider_matrix.load_builds())`
and `support.write`, and cold protocol with `cold_matrix.py freeze-protocol`.
Formal capture and analysis commands, run from repository root, are:

```sh
evidence_dir=docs/performance/results/change-0491
python3 -B "$evidence_dir/gate.py" provider-formal1 python3 -B "$evidence_dir/provider_matrix.py" capture-all --attempt provider-formal1
python3 -B "$evidence_dir/provider_matrix.py" analyze --attempt provider-formal1
python3 -B "$evidence_dir/provider_matrix.py" verify --attempt provider-formal1
python3 -B "$evidence_dir/gate.py" cold-formal1 python3 -B "$evidence_dir/cold_matrix.py" capture-all --attempt cold-formal1
python3 -B "$evidence_dir/cold_matrix.py" analyze --attempt cold-formal1
python3 -B "$evidence_dir/cold_matrix.py" verify --attempt cold-formal1
python3 -B "$evidence_dir/gate.py" profiles2 python3 -B "$PWD/$evidence_dir/profile_providers.py" profiles2
python3 -B "$PWD/$evidence_dir/profile_providers.py" --verify profiles2
python3 -B "$evidence_dir/write_results.py"
```

The wrapper runs children from the repository root. Exact actual commands and absolute paths are retained in each `validation/*.started.json`.
Analysis/verification output is exclusive; use a new attempt for a new capture.
After all jobs are terminal, run `cleanup.py --dry-run`, `cleanup.py`, and create
the seal using the selected final gates/build/verification/cleanup paths.
`python3 -B seal.py verify` rechecks the completed bundle without overwriting it.

Validation includes the complete benchmark library suite, allocator tests,
warning-denied Clippy/rustdoc, scoped rustfmt, doc tests, helper mutation tests,
and crate-boundary checks. This is not a claim that every workspace feature,
fuzz target or native Office scenario was revalidated by this benchmark batch.
