# 0359c — Typed legacy edit budgets

- Date: 2026-10-21.
- Status: in progress. The XLS preflight and primary staging slice passes its
  targeted validation; shared OLE capture and Graph budget errors remain required work.
- Contract: 0351 §2.4 requires `Limit { kind, observed, maximum }` and
  `InvalidLimit { kind, value, maximum }` on staging/apply paths.

## XLS preflight and primary staging

Embedded-payload byte admission, publication directory/stream/storage bounds,
aggregate stream bytes, workbook/summary stream admission and new-object count
preflight now return typed budget errors. The `ole_object::LimitKind` enum names
these dimensions instead of remaining empty. Zero stream/aggregate limits
already rejected by these paths now return `InvalidLimit`. Error messages name
only fixed dimensions and numeric limits; they do not include stream names or
payload content.

The existing root `LimitKind` now also covers physical ingress and nested
OLE-object dimensions. `SourceLimitKind` is removed, and every consumer uses the
canonical kind. The old cell `StagedChangeLimit` variant becomes `Error::Limit`
with `LimitKind::Operations`; the primary edit adapter preserves its established
operation-limit error and unified code. `Error` is non-exhaustive, and configured
limit values use `u64` so byte budgets retain their width on 32-bit targets.
There are no deprecated names or compatibility aliases.

Every existing comparison, default ceiling, payload check order and successful
publication operation is retained. The checks still precede nested property-set
parsing and remain failure-atomic. Budget-error construction no longer allocates
an `InvalidData` message. Existing source-limit and staged-change display text is
preserved; the newly typed OLE messages are fixed and content-free.

The unified classifier maps the emitted `Limit` to `limit_exceeded` and
`InvalidLimit` to `invalid_argument`. Previously these XLS preflight cases used
`InvalidData`, which classified as `corrupt_file`. Tests use real payload and
snapshot entry points, check their typed dimensions and numeric evidence, and
verify unified classification rather than parsing error strings.

## Validation of this slice

All 1,313 XLS library tests pass, including the existing corpus and mutation
comparisons, as do 33 unified XLS error-classification tests. XLS all-target,
all-feature compilation and clippy pass; the XLS-enabled operations crate
passes all-target clippy. The independently locked performance harness compiles
all targets. The XLS glossary checks 97 banned names, 13 families and three
immutable owners successfully, including bans on both removed error names. Formatting, all 36 tooling tests
and the benchmark-matrix consistency check also pass.
Commands, source hashes and external log hashes are retained in
[the validation receipt](results/0359c/xls-validation.json).

This slice makes no timing claim. The integrated corpus/code-disposition and
performance checks remain required after the shared OLE/Graph migration below;
unit coverage of the newly emitted codes is not a substitute for that gate.

## Required continuation

The shared `litchi-ole-common::object` capture/editor still reports several
configured bounds through `OleError::InvalidFormat`. Its typed error migration
must cover every DOC/XLS/Graph/PPT consumer in the same change. Graph's string
resource labels and the remaining family budget paths also need review. This
record does not accept those paths as exceptions or complete the legacy error
audit. The final integrated corpus, performance and workspace gate remain
required alongside W4/W4b/W8 closure.
