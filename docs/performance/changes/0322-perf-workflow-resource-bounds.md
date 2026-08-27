# Change 0322: performance workflow resource bounds

## Scope

This change hardens `.github/workflows/perf-baseline.yml` against concurrent
build pressure and persistent repository-local Cargo artifacts. The `smoke`
and `full` Cargo workload jobs use one Cargo build job, disable incremental
compilation, and direct `CARGO_TARGET_DIR` to an explicit `runner.temp`
location. Those workload jobs have always-on disk/memory/target-size
diagnostics and cleanup that uses `find -xdev -depth -delete` without
recursive forced removal. Every declared job, including the non-Cargo
`reference-regression` job, has a positive timeout.

Artifact uploads remain visible when an earlier step fails by using
`if: always()`. The obsolete `tools/perf-baseline/target` cache path is not
part of the workflow contract.

## Static policy

`tools/test_perf_workflow_policy.py` checks the workflow using only Python's
standard library. It deliberately uses a small indentation-aware structural
reader instead of adding a PyYAML dependency. It reads Cargo settings from
YAML `env:` maps and `always()` from step-level YAML fields, resolves the
Cargo target for each workload job from its own or the workflow-global `env`,
and checks every declared job for a timeout. Any new Cargo command must be
added to the explicit workload set before it can bypass diagnostics or
cleanup.

## Verification status

Validation passed for the non-Cargo batch:

- `python3 -m unittest tools.test_perf_workflow_policy`: 9 tests passed.
- `git diff --check` passed for the three Change 0322 files: the workflow,
  policy test, and this record.

No Cargo build, benchmark, or workflow run was performed or is claimed by this
record.

## Claims deliberately not made

This change makes no claim about lower RSS, peak heap, OOM avoidance in every
environment, throughput, latency, allocations, or benchmark performance. It
only defines CI resource controls and failure diagnostics intended to make
resource behavior bounded and observable.
