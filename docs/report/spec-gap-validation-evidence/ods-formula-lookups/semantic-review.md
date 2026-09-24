# Implementation semantic review of the ODS lookup and reference batch

Status: **source semantic PASS and final isolated gates PASS for the current
frozen candidate**. This document owns the
implementation review that was previously placed in `spec-review.md`. The
committed contract and independent contract history remain in
[spec-review.md](spec-review.md), which has been restored to its `HEAD`
content. The current source includes moved-key handling for the owning H/V path,
the scalar lookup-key error precedence fix, and the latest ADDRESS character-loop
cleanup. The source review is complete and the final isolated gate receipt
passes. Performance acceptance remains separate and pending. No production or
test edits were made by this review.

The review covers `ADDRESS`, `CHOOSE`, `HLOOKUP`, `INDEX`, `INDIRECT`,
`LOOKUP`, `MATCH`, `OFFSET`, and `VLOOKUP`. The accepted contract is
`contract.md`, SHA-256
`b112d66d687337912333f932c199f6e0ee5241fedc335aeaa95de572889e6aaf`.

## Review snapshot

The current staged candidate manifest is
[`gates/freeze.json`](gates/freeze.json), SHA-256
`53261588afdad7e1efd29ac91c0f393f41e8ddbe1032f91ca9d60dcb0c86fa60`; its
status field retains the staging value `staged; final gates not run`; the
terminal results and verification receipts below are authoritative for this
source snapshot.
The selected implementation hashes for the reviewed seams are:

| File | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation/lookup/search.rs` | `2b0920dfc73a1282a9db4272499f96be3a0a3811d08f4774f8464d5c95a384de` |
| `crates/litchi-ods/src/codec/formula/evaluation/lookup/address.rs` | `7da925466d57320d4727f9ff992edc92ab64d7276fe32442464edaacee9771d8` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/lookup/search.rs` | `3e2f001f91dd9d690730798c0a22ecae4867e427ae6b5f333ef07e6fa0b8b225` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/lookup/lifting.rs` | `d202a46463fe91173bd42027f8a383012a43a6b5ce158db231b1faa2a03d0571` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/lookup.rs` | `58bf26633bed3c443a2cadbee10a1b47eb6ec5b9683a2d56284932af7459c0be` |
| `crates/litchi-ods/tests/ods_formula_lookup_evaluation.rs` | `5537e8820144cc587c6f13c794b3d02edc12fdf0b0c988c26e3919f77faa0965` |
| `crates/litchi-ods/tests/ods_formula_lookup_limits.rs` | `16a9ff3fdf0bb3c65148f68e91ca1ae2455c26c5ba6cbda03377554efccf9899` |

The terminal `gates/results.json` receipt has SHA-256
`284149e608bc64da65347e012974d06acb46d35e2882d09f6615b6a867116237` and
records all seven required commands with exit code 0. The terminal
`gates/verification.json` receipt has SHA-256
`bdeb941ccf550e43a341df081dcd54e1ba9e70779744ecb80e785cca71739243` and
reports required checks passed with stable sources. The isolated run records
1,700 tests, zero failures, and zero ignored tests. The prior 1,699-test
receipt is retained only as diagnostic history.

The earlier diagnostic attempt 3 passed all seven gates and 1,697 tests, but
its selected-source archive SHA-256 was
`e7e8be1939b7d373f65b7d35710f69eec19211e1bf953ecf82b2803a1a88ccdc`. That
attempt is superseded: focused review found a selector/reference-key read
ordering failure. Its receipt is historical evidence only. The current
candidate includes the selector-ordering, deferred-reference, moved-key,
scalar-key-error, and character-loop fixes; the terminal receipt covers this
source snapshot.

## Resolved implementation semantics

The owning and borrowed table-lookup entry points share `table_controls` and
`finish_table_lookup`. The owning path moves its raw key; the borrowed path
projects or clones scalar slots as needed. Both validate data geometry, then
validate scalar index and range/match mode before projecting a reference-valued
key. Literal controls are classified without resolver reads; reference-valued
controls are deferred until known literal refusals have passed. For a known
invalid selector this preserves the contract's zero-key-read rule. A direct
formula-error key remains the original formula error when the selector itself
produces a formula refusal. Typed resolver, cancellation, resource, source,
and unsupported failures remain evaluation failures.

`lifting.rs` derives output shape only from scalar key, index, mode, and
selector arguments. Complete ForceArray search data does not widen the output
shape. Reference-valued scalar arguments become deferred `ScalarCell`
descriptors during mapping; their cells are resolved only by the search kernel
after that output coordinate's controls are known to be valid. The mapper
broadcasts singleton, row, and column inputs, selects each coordinate
independently, and publishes per-coordinate `#N/A` for a source coordinate
outside the broadcast operand. Invalid scalar selectors therefore produce an
error array with the selector-derived shape, while list, 3-D, and other
data-shape refusals occur before cell reads.

ForceArray data preflight is a whole-operand admission check. If scalar data,
a ReferenceList, a 3-D reference, or another invalid data geometry is rejected,
the function returns one global `#VALUE!` before matrix lifting, even when the
lookup key is array-valued; ForceArray data does not provide output shape. This
does not require an error array or any resolver cell read. A direct scalar
formula-error key is the explicit precedence exception: its original error is
returned over that data refusal with zero reads. Once data is admitted, ordinary
key-array lifting and per-coordinate key-error publication apply.

Exact search retains the first visited formula error while continuing the
required source-order scan. A later typed provider failure supersedes that
retained formula error. Approximate search uses only its bounded probes and
does not inspect unvisited cells. Selected result cells are read only after a
match. These rules preserve formula-error identity without catching typed
provider failures as formula values.

The earlier CHOOSE shape/cache seam is also cleared for semantic purposes:
selected branches are probed under the selected outer mask, computed-index
probes release their depth reservation, and complete values are retained only
for the exact projected demand. A dynamically widened selected reference may
revisit its position-sensitive selector; that is a conservative cache boundary
and does not probe an unselected branch.

## Error and shape re-review

`table_controls` checks direct formula errors in the index and range controls
before conversion. It preserves argument-order precedence, then classifies
literal controls before resolving deferred reference controls. A reference
index or mode is read only when it is needed to determine that coordinate's
control value. The lookup key is resolved last. This ordering means a typed
failure from a required control or key remains an `EvaluationFailure`, while
an admitted formula error remains a formula result and is not swallowed by a
selector refusal.

The focused mixed-control cases now cover a valid and an invalid coordinate in
the same lifted result. The invalid coordinate publishes `#VALUE!` without
reading its key or search data; the valid coordinate reads only its key,
search candidate, and selected result. The same shape and precedence hold for
literal arrays and lifted reference controls in both scalar and matrix modes.
Exact search retains the first visited formula error while continuing the
required source-order scan. A later typed provider failure supersedes that
retained formula error. Approximate search uses only its bounded probes and
does not inspect unvisited cells. Selected result cells are read only after a
match.

No source semantic blocker remains in the frozen candidate. Moved-key H/V
handling avoids cloning an owned computed Text key while retaining the same
table-control and error-ordering path. The scalar-key preflight precedence and
global invalid-data refusal are now covered by the current source scope. The
terminal isolated gates pass; this document does not turn the candidate into a
production-support or performance claim.

## Required recheck

The current handoff records 34 focused semantic cases, 9 resource cases, 127
oracle observations, and a terminal isolated receipt of 1,700 tests with zero
failures or ignored tests. Native and oracle artifacts remain separately
reviewed; no performance conclusion is drawn here.
