# 0360el — XLSB stale replay classification and recipient admission

- Status: source, consumer, and focused replay tests complete on
  `uo2/w4b-xlsb-raw-policy`; native performance qualification remains open.
- Base: `89146628f3` (0360ek replay admission).
- Scope: stale source replay in the XLSB package and raw external-link
  families, unified-ops classification, affected assertions, and the
  recipient-image allocation preflight.

## Patch conflict classification

An exact before-image or source-closure mismatch is a patch conflict. The
primary workbook package, cell-value worksheet/package/structure paths,
calculation chain, volatile dependencies, data model, XML Maps, Theme,
web-extension bindings, shared-workbook tests, and external-link transaction
now preserve that fact as `PackageError::PatchConflict` or the family error's
`PatchConflict`. The host conversion keeps external-link conflicts typed as
the XLSB package conflict.

The facade's process-local owner guard runs before the format leaf, but it
accepts a reopened package when its current bytes match the patch base. A
different owner identity alone is not a conflict. When the current bytes do
not match, the guard reports
`Source::Transaction(transaction::Error::PatchConflict)`; the leaf tests
assert the format-level `PackageError::PatchConflict`, while the unified
cell-write test asserts the transaction-level source and its `InvalidInput`
code. These are two boundaries of the same content-mismatch classification,
not different source-change semantics.

Malformed content, post-publication invariant failures, and source-backed
freshness remain their existing errors. In particular, `SourceChanged` is
still an I/O classification; it is not converted to `PatchConflict`.

## Recipient admission before an exact-image copy

Structural replay publishes the retained exact after image through the
recipient's `SourceLimits`. A borrowed physical-reader preflight checks input,
archive/member counts, names, metadata, and declared ZIP sizes before the
patch takes ownership. The accepted image is then passed through
`from_shared_vec_with_limits`, so the exact `Arc<Vec<u8>>` is retained by the
published OPC package without a second whole-image clone. BIFF12 raw admission
still runs through `project_candidate` and complete worksheet validation before
the package is returned; the physical preflight does not replace that semantic
validation. Replay-generated images use the same one-Arc handoff after their
byte and archive preflight.

## Validation

Using the capped environment at
`/home/zhuhe/code/litchi-worktrees/.coord/w4b-xlsb-raw-policy/env.sh`, with a
worktree-local temporary directory because the shared `/tmp` inode budget was
exhausted:

- `cargo test -p litchi-xlsb --test data_model_identity --test package_cell_values --test workbook_structure_edit --test volatile_dependencies_review --test xml_maps_public`: 78 passed, 0 failed.
- `cargo test -p litchi-ops --features xlsb --test xlsb_cell_writes`: 24 passed, 0 failed.
- `cargo fmt --all` and `git diff --check`: passed.

No native timing, profiler, binary, or large evidence file is committed.
