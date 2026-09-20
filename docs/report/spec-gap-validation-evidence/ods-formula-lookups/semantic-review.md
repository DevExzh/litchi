# Implementation semantic review of the ODS lookup and reference batch

Status: **source semantic PASS for the staged candidate**. This document owns the
implementation review that was previously placed in `spec-review.md`. The
committed contract and independent contract history remain in
[spec-review.md](spec-review.md), which has been restored to its `HEAD`
content. The current source includes moved-key handling for the owning H/V path,
and the refreshed isolated gate receipt passes with stable source hashes.
Performance acceptance remains separate and pending. No production or test
edits were made by this review.

The review covers `ADDRESS`, `CHOOSE`, `HLOOKUP`, `INDEX`, `INDIRECT`,
`LOOKUP`, `MATCH`, `OFFSET`, and `VLOOKUP`. The accepted contract is
`contract.md`, SHA-256
`b112d66d687337912333f932c199f6e0ee5241fedc335aeaa95de572889e6aaf`.

## Review snapshot

The current staged candidate manifest is
[`gates/freeze.json`](gates/freeze.json), SHA-256
`ec49b19183755117243d5e823222357338022a19b91a2328c746d38c220e9598`; its
status field is the retained historical staging value `staged; final gates not
run`; the results and verification receipts below are the actual gate outcome.
The selected implementation hashes for the reviewed seams are:

| File | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation/lookup/search.rs` | `2b0920dfc73a1282a9db4272499f96be3a0a3811d08f4774f8464d5c95a384de` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/lookup/search.rs` | `3e2f001f91dd9d690730798c0a22ecae4867e427ae6b5f333ef07e6fa0b8b225` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/lookup/lifting.rs` | `d202a46463fe91173bd42027f8a383012a43a6b5ce158db231b1faa2a03d0571` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/lookup.rs` | `ff89fa7241d504c201350936c1a35193d5a33630204e5a0a5fa317a7a2d5e643` |
| `crates/litchi-ods/tests/ods_formula_lookup_evaluation.rs` | `32dd9b97bdeda505e71102c9e5de166c60b527a8a9a2209c128fde2a81c81837` |
| `crates/litchi-ods/tests/ods_formula_lookup_limits.rs` | `16a9ff3fdf0bb3c65148f68e91ca1ae2455c26c5ba6cbda03377554efccf9899` |

The current isolated gate receipt is `gates/results.json`, SHA-256
`06a1a8752f820abf54935f5f1b54e0db701547d6555724983bd867625823cba7`;
`gates/verification.json` has SHA-256
`bdeb941ccf550e43a341df081dcd54e1ba9e70779744ecb80e785cca71739243` and
reports both required checks passed and stable sources. The receipt covers all
seven gates and 1,699 tests, with zero failures and zero ignored tests.

The earlier diagnostic attempt 3 passed all seven gates and 1,697 tests, but
its selected-source archive SHA-256 was
`e7e8be1939b7d373f65b7d35710f69eec19211e1bf953ecf82b2803a1a88ccdc`. That
attempt is superseded: focused review found a selector/reference-key read
ordering failure. Its receipt is historical evidence only. The current
candidate includes the selector-ordering, deferred-reference, and moved-key
fixes; the receipt above covers this source snapshot.

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

No source semantic blocker remains in the staged candidate. Moved-key H/V
handling avoids cloning an owned computed Text key while retaining the same
table-control and error-ordering path. The current freeze and gate receipt
cover that implementation. This document does not turn the current candidate
into a production-support or performance claim.

## Required recheck

The current handoff records 33 focused semantic cases, 9 resource cases, and
127 oracle observations. The current isolated receipt records 1,699 tests with
stable sources. Native and oracle artifacts remain separately reviewed; no
performance conclusion is drawn here.
