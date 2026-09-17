# Log sections for change 0674

The coordinator merges each block into the newest position of the named file.

## For `HOTSPOTS.md`

## 0674 — the remaining 0651 harness gates and reproducibility hygiene are closed

Record: [0674](0674-performance-gate-hygiene.md). Rows 13, 17, 19 and 20 of
0651 are closed as harness and evidence work. `pptx_semantic_opened_transaction_phases`
now separates `opened_presentation`, `Snapshot::edit`, `set_shape_text`,
`Transaction::commit`, `apply_opened_presentation_commit` and `to_bytes` under
one enclosing total, with semantic reopen, output digest and phase-sum gates.
The opt-in `xlsb_semantic_workbook_structure_edit_save` selector likewise
separates detached workbook-structure planning/commit,
`apply_workbook_structure` and save, then reopens the renamed worksheet and
checks a typed cell. The default catalog remains unchanged; the selectable
registry rises from 528 to 529. The named `docx,odt` facade gate and serial
allocator subcommand cover the two omitted test surfaces. The stale
native-resave lockfile is regenerated and the result-log negation now covers
`docs/performance/results/**/*.log`; packet owners still choose which
historical or new logs to add explicitly. The harness README documents the
release profile's `lto = true` rebuild identity limit. No performance claim is
registered.

## For `GOAL_AUDIT.md`

## 0674 — phase evidence and standing gates are scoped without changing product behavior

Record: [0674](0674-performance-gate-hygiene.md). The new selector names its
semantic tiny PPTX corpus, one selected edit, six public API stages, enclosing
total, output digest and reopen oracle; package construction and validation
remain outside the clocks. It is opt-in and leaves `Case::DEFAULT` and the
catalog hash unchanged. The DOCX/ODT gate names its exact feature closure, and
the allocator gate names its isolated binary and one test thread. These are
reproducible harness and gate facts, not generalized latency or allocation
claims. Structural mode validates ten claims without opening retained evidence;
the named strict claims gate validates the same ten claims and independently
checks their evidence root. Strict mode still requires an evidence root.

## For `REPORT.md`

## 0674 — an opened PPTX edit can be attributed by phase, while omitted test surfaces become standing gates

Record: [0674](0674-performance-gate-hygiene.md). The phase selector emits
aligned vectors for the six transaction stages and `total_ns`, verifies the
edited package after reopening, and requires one deterministic output digest
across retained samples. It reports attribution evidence only; no timing
number, speedup, allocation ratio, physical-I/O result or claim-registry entry
is published. The XLSB structure selector also retains phase vectors, output
digest and typed reopen checks without a performance claim. The facade
polyglot command runs the `docx,odt` feature pair and the harness command runs
the process-global allocator tests serially. The packet's `gates.log` retains
the focused outcomes for both claims modes.

## For `ADR_COMPLIANCE.md`

## 0674 — no production boundary or safety defense moved

Record: [0674](0674-performance-gate-hygiene.md). Nothing under `crates/`
changed, so no public API, error identity, malformed-input defense, output
contract or ADR boundary moved. The harness remains in `tools/` and retains
`#![forbid(unsafe_code)]`; allocator isolation changes test scheduling rather
than library behavior. The checked PPTX phase selector preserves semantic
reopen and deterministic-output gates, and its checked basis-point division
keeps the existing overflow defense while satisfying clippy. The native-resave
lock refresh is metadata hygiene. The result-log negation is scoped after the
broad rules, so result-packet logs are visible while callers still select what
to add. Structural claims validation now succeeds without retained evidence;
strict evidence verification remains the named workflow gate.
