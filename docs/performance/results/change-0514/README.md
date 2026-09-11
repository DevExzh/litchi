# 0514: rejected XLSX source parser/layout fusion

This bundle retains a production prototype rejected at its predeclared pilot
admission gate. Cold same-value edit latency and temporary allocation demand
increase materially. Production and candidate tests are restored to control
revision `33a21e0f087b4ead80ca63187c7cf30d0584b1f6`; the standalone public
guard remains reusable. `performance_claim: none`; `claim_authorized: false`.

Read the [change report](../../changes/0514-xlsx-fusion-rejection.md),
[decision](decision.json), [admission review](admission-review.md), and
[follow-up options](follow-up-options.md). This batch does not complete the
OLE2/OOXML goal. ODF remains deferred until that goal is complete; iWork is
excluded.

## Replay the evidence

From the repository root:

```sh
python3 -B docs/performance/results/change-0514/verify.py
```

From this directory, `sha256sum -c SHA256SUMS` checks the retained byte
inventory. The verifier checks source, binary, corpus, sink, sample, allocator,
profile and restoration bindings, including exact candidate patch replay from
the base. Replay does not require the removed benchmark executables. Raw
profile annotations intentionally retain profiler whitespace.

## Source and measurement epochs

- `before/source-manifest.json` binds control production and harness sources.
- `source-manifest.json` binds the rejected candidate, not the restored live
  tree. `candidate.patch` and `candidate-patch.json` retain all 14 changed
  paths, including 17 new tests. `restoration.json` binds exact base recovery.
- `guard-source-manifest.json` binds the independent probe, its Cargo manifest
  and lockfile, and the unchanged canonical allocator observer/wrapper. Both
  roles use that same probe and lock. Its corpora differ from the main harness.
- Role-specific build receipts bind normal and instrumented executables. The
  main candidate build receipt is at the bundle root; guard and allocator
  build receipts are under their role directories.
- Main native control repeats use 500 samples and five warmups per row;
  guard control repeats use 100 and three. Matched candidate/control pilots
  use 20 and two. Full candidate native repeats were not admitted.
- Both roles have two allocator captures, each with 10 samples and one warmup
  per row. Instrumented elapsed time and RSS are excluded from native deltas.
- Save profiles collect three calls to `xlsx_commit_save_operation`, excluding
  fixture generation, expected-output operations, oracles and caller teardown.
  Descendant call metadata can include collection-disabled calls; collected
  instruction costs remain operation-scoped. Both Valgrind warnings are
  retained alongside successful exit status and runtime oracles.

Absolute `region_peak_live_bytes` includes entry live allocations. Compute
`region_peak_live_bytes - live_bytes_before` per aligned sample for incremental
live demand. Neither value is document peak or RSS. Main and guard scenarios
also have different result lifetimes; compare memory within matched scenarios.

## Reproduce a workload

The capture/build scripts use exclusive output files so they cannot overwrite
sealed evidence. For a new experiment, use a separate checkout and output
directory, retain this bundle's independent `guard-probe` at the same relative
path, and use the exact commands and environments recorded in the receipts.
Control production sources must match the before manifest. Apply
`candidate.patch` only for a candidate reproduction, then verify the candidate
manifest before building. Do not apply it to the restored working checkout
merely to replay the evidence verifier.

The main harness lives at `tools/perf-baseline/Cargo.toml`; the independent
probe lives at `guard-probe/Cargo.toml`. Build with the recorded Rust toolchain,
`--release --locked`, the same features, and a fresh target directory. Freeze
the normal and allocator binaries separately, then run CPU-2-pinned workloads
serially with fresh JSON destinations. [guard-probe/README.md](guard-probe/README.md)
documents its public scenarios and chronological vectors. Main harness vectors
are elapsed-sorted with explicit sample-index alignment.

## Validation and prototype history

The final candidate passes 966 XLSX owner unit tests, all-features library
Clippy with warnings denied, formatting, crate boundaries and strict existing
claim checks. This is prototype validation, not a retained production change.
Full candidate native admission and broader retention-only test lanes were
not run after the pilot rejection.

`initial-unit`, `second-unit`, `third-unit`, and `fourth-unit` preserve failed
pre-measurement test builds/runs and their corrections. They are not performance
captures. `before/guard-initial-build` similarly preserves the probe's initial
CLI borrow compilation error. Final builds and measurements bind the corrected
source manifests. The owned scratch cleanup is recorded in `cleanup.json`;
unrelated worktrees and targets are outside this batch's cleanup scope.
