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

Focused unit coverage asserts that a packaged edit starts with no mutable
draft, that an exact no-op commit does not materialize one, that the first
slide operation does, and that root/PI refusal is deferred until commit.  The
existing flat replacement/no-op/inverse suite and all ODP CRUD, source-backed,
replay, limit, refusal, fidelity and streaming tests remain green.

## Validation

The isolated worktree used one capped Cargo job at a time, with
`--offline --locked`, debug information disabled, incremental compilation
disabled, and the workspace-local temporary directory from the coordinator
environment:

* `cargo test --offline --locked -p litchi-odp --all-targets --all-features` —
  passed (236 unit tests plus all ODP integration targets; no failures).
* `cargo clippy --offline --locked -p litchi-odp --all-targets --all-features -- -D warnings` — passed.
* `cargo fmt --all` and `git diff --check` — passed.

No native timing, baseline build, or performance exception is claimed in this
slice.  The compact command receipt is
[`results/0362uga/validation.json`](results/0362uga/validation.json).
