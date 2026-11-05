# 0362uga — ODP edit-construction cost closure

- Date: 2026-10-23.
- Parent: W8 final sweep, ODP edit-construction debt.
- Status: source repair complete; coordinator integration remains required.

`litchi-odp` now keeps `Package::edit()` and flat
`flat::Presentation::edit()` as cheap source-bound handles.  The packaged
edit retains only its immutable `Package`, operation journals and accounting
scalars; its `MutablePresentation` is created on the first slide or media
operation.  That first materialization reuses the package's validated slide
projection and keeps the existing staging, source-fragment, style, policy,
limit and audit checks in the same order.  Chart, RDF, design, annotation,
layer and rich-content operations that do not need the slide draft continue to
stage against the immutable source package.

The flat edit now uses `Option<Vec<Slide>>`: reads borrow the source projection,
the first real replacement clones it, and reset drops the draft back to the
source.  A no-op replace therefore does not allocate a second slide vector,
and a no-op commit still returns the exact source bytes and an unchanged patch.

The edit boundary still rejects the established signed/encrypted policy.  The
active-processing-instruction and invalid-root refusal cases are deferred out
of `Package::edit()` and run on the first fallible operation or commit, so a
malformed source is refused before publication without adding source-sized
work to edit construction.  Generated builder packages, whose projection is
not yet cached, therefore remain fully lazy until their first operation or
commit.

The packaged edit's slide and layer projections are fallible as well: deferred
source or entry-admission errors are returned to the caller instead of being
converted into dummy values.  Successful entry admission is cached for the
transaction, so repeated read projections do not rescan the source.

Security admission remains ordered before source XML root/processing-
instruction validation, preserving the established refusal precedence while
deferring both checks.  Empty scene-splice requests also pass through
admission, so they cannot hide a signed or malformed source.

Focused unit coverage asserts that a packaged edit starts with no mutable
draft, that an exact no-op commit does not materialize one, that the first
slide operation does, and that root/PI refusal is deferred until commit.  The
existing flat replacement/no-op/inverse suite and all ODP CRUD, source-backed,
replay, limit, refusal, fidelity and streaming tests remain green.

## Validation

The isolated worktree used Rust 1.95.0, one capped Cargo job at a time,
`--offline --locked`, a 20G memory cap, debug information disabled,
incremental compilation disabled, and the workspace-local temporary
directory from the coordinator environment.  The target directory was kept
under the separate 20,000,000,000-byte disk cap.  The source validation commit was
`363a2497f9103a32d41aff47bb0d2c9bed040521`:

* `cargo test --offline --locked -p litchi-odp --all-targets --all-features` —
  passed (38 suites, 448 passed, 0 failed, 0 ignored, 0 measured).
* `cargo clippy --offline --locked -p litchi-odp --all-targets --all-features -- -D warnings` — passed.
* `cargo check --offline --locked -p litchi-ops --all-targets --all-features` — passed.
* `rustfmt --edition 2024 --check` on the changed ODP source/tests and
  `git diff --check` — passed.

The raw command logs, byte counts, SHA-256 values, toolchain preflight and
target-size observations are recorded in the compact receipt below.  The
receipt commit itself only changes documentation and therefore does not alter
the validated ODP source.

The consumer check ran with one known unrelated working-tree overlay:
`crates/litchi-xlsx/src/workbook/edit/semantic/transaction.rs`, a pre-existing
formatting-only change.  It was preserved, not attributed to this work, and
the validation does not claim an exact-HEAD-clean checkout.

No native timing, baseline build, or performance exception is claimed in this
slice.  The compact command receipt is
[`results/0362uga/validation.json`](results/0362uga/validation.json).
