# 0359c — Typed legacy edit budgets

- Date: 2026-10-21.
- Status: in progress. The XLS and Graph typed-budget slices pass targeted
  validation; shared OLE capture and the Graph owner contract remain required work.
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

## Graph typed budgets and host propagation

Graph ingress, authoring and fixed-record edits now use `Error::Limit` with a
finite `LimitKind`. `InvalidLimit` carries the requested value and supported
maximum. BIFF budget failures normalize into those same variants, preserving
the physical resource dimension and numeric evidence. Non-budget BIFF failures
retain their typed cause. All native constructors, helper parameters, tests and
PPT host consumers migrate; no old `LimitExceeded` alias remains in Graph.

The format grammar still rejects more than two axis parents as malformed
input. Its subsequent, unreachable budget check is removed; the two-parent
constraint is not a configurable resource policy. Configurable comparisons,
default ceilings and byte-copy operations are retained.

The operations layer now classifies Graph and BIFF causes directly through
both XLS and PPT errors, including PPT authoring errors. This corrects the old
blanket `Graph(_) -> corrupt_file` mapping: configured budget exhaustion is
`limit_exceeded`, invalid configuration is `invalid_argument`, conflicts are
`invalid_input`, refusals follow their kind, and allocations are `internal_error`.
Malformed chart/framing input remains `corrupt_file`. Graph and BIFF are direct
optional operations dependencies enabled only by XLS/PPT; the crate-boundary
policy and all 27 tracked current consumer lockfiles receive those path edges.
No dependency version changes.

The paired native probe in `results/0359c/graph-budget-probe.{py,rs}` compares
against `b6f5225d15` with identical dependency locks and the tracked
`WithThreeCharts.xls` fixture. It checks exact unchanged encoding, changed
output, forward replay and exact inverse before timing. Initial ABBA blocks
exposed a repeatable parse/edit/replay slowdown of about 4–5%. Outlining BIFF
error conversion alone made policy validation slower; restoring its inlining
repaired validation but left the parse slowdown. All three failed candidates'
samples are retained, rather than accepting the regression as code placement.

The final change removes duplicate chart framing scans when copying or
relimiting a stream under its unchanged, already validated policy. Borrowed
views can only be constructed after admission; internal byte edits preserve
record framing and patch checked, fixed-size payloads. A changed policy still
requires full admission, and the output-byte check always runs before copying
or returning an owned stream. Two boundary tests cover tighter record limits
and zero output allowance, including when the retained policy is unchanged.
BIFF error normalization remains cold and successful policy checks stay inline.

Three isolated, CPU-pinned ABBA blocks compare the final candidate with
`b6f5225d15`. Each block has two baseline and two candidate processes; each
process checks exact forward/inverse output before timing. Ratios are candidate
over baseline, with the median and range across the three blocks:

| Route | Median | Range |
| --- | ---: | ---: |
| Validate policy | 1.002 | 1.000–1.002 |
| Scan chart substreams | 0.994 | 0.994–0.997 |
| Parse and encode all three charts | 0.937 | 0.936–0.939 |
| Parse, edit and commit | 1.002 | 0.997–1.006 |
| Parse and replay | 0.994 | 0.991–1.004 |

The earlier sustained slowdown is removed. Error and `Result<()>` layouts stay
48 bytes, and the optimization does not change the copies or allocations on
successful paths. These timings cover the native Graph routes above, not full
facade operations or shared OLE capture. Complete samples, binary/source/lock
hashes, fixture provenance and the reproduction commands are in
[the measurement receipt](results/0359c/graph-measurements.json); rejected
samples are in `graph-rejected-{0,1,2}.json` beside it. No binaries are committed.

All 75 Graph tests, 1,313 XLS library tests, 19 PPT chart tests and 39 unified
error tests pass (1,446 total, none failed or ignored). Graph/PPT and operations
all-target clippy, the locked performance harness, four glossary profiles,
crate boundaries, formatting, 36 tooling tests and the benchmark-matrix check
pass. At this slice the Graph glossary had no owner/family registrations;
[0359d](0359d-graph-package-owner.md) subsequently registers the canonical
root family and immutable owners. Commands and
source/log hashes are in [the validation receipt](results/0359c/graph-validation.json).

## Required continuation

[0359e](0359e-object-capture-budgets.md) migrates the shared
`litchi-ole-common::object` core capture/editor budgets to typed errors with
all direct consumers, DOC callbacks, XLS sheet conversion and unified
classifiers. Raw directory, OLEDS, link and protection policy/error surfaces
remain required work; the core capture migration does not close them.

[0359d](0359d-graph-package-owner.md) consolidates Graph's standalone edit
onto `Commit<Package, Patch>`, removes the second Snapshot owner and migrates
all consumers. Its unchanged-stream sharing and nested raw chart model audit
remain required. The shared OLE object's snapshot publication needs the same
return-type audit against its raw-codec role. Zero name-inventory rows do not
prove these semantic obligations; they are not accepted exceptions.

This record does not complete the legacy API/error audit. The final integrated
corpus, performance and workspace gate remain required alongside W4/W4b/W8
closure.
