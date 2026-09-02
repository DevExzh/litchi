# Change 0376: XLSB Single Cell Tables lifecycle

## Scope

Change 0376 closes the opened-document CRUD gap for canonical XLSB
`tableSingleCells` owners. Existing `xml_maps::Transaction` entry points can
now add the first single-cell binding to a worksheet that has no Single Cell
Tables part and can remove the owner after its final binding is deleted. No
public API shape changes.

This implements the first-create and final-remove cases required by
[GOAL.md](../../GOAL.md) and the
[CRUD scenario checklist](../../CRUD_Scenario_Checklist.md) without extending
the slice to ordinary-table part lifecycle, XML instance import/export,
schema or XPath evaluation, connection refresh, or formula execution.

## Architecture and behavior

`litchi-xlsb` remains the semantic and BIFF12 owner. First creation selects a
collision-free canonical `/xl/tables/tableSingleCells{N}.bin` using
URI-equivalent comparison, serializes the canonical dependency, and adds one
internal relationship from the owning worksheet. Content-type and dependency
state are staged with the existing package transaction.

Final removal records and deletes the exact owning worksheet relationship,
checks the source for lossless removal, and scans package-root plus all part
relationships for resolved URI-equivalent inbound references before deleting
the dependency. Shared, orphan, foreign, external, malformed, case-aliased
dangling, opaque/FRT, or noncanonical changed forms return typed refusal. An
untouched existing empty owner remains byte-preserved.

Publication remains clone-staged and swaps only after a complete workbook
reparse and postcondition check. Source lineage/freshness, limits, exact no-op
bytes and signatures, effective-change signature invalidation, failure
atomicity/retryability, and owned payload/topology inverse behavior remain
part of the existing `Commit`/`Patch` contract. Inverse application does not
restore invalidated signatures.

## Verification

All Cargo work used one process at a time, `CARGO_BUILD_JOBS=1`, one dedicated
target, disabled incremental state, one test thread, an available-memory
launch guard, and at most an 8 GiB per-process virtual-memory cap. These are
operational safeguards, not OOM-prevention evidence.

The focused `xml_maps_public` target passed `30/30` executed tests. Coverage
includes canonical first creation and reopen, final removal and inverse,
collision allocation around case-equivalent names, absolute-equivalent owner
targets, root and internal inbound-reference refusal, foreign/unowned source
refusal, opaque lossless-removal refusal, exact existing-empty no-op,
failure atomicity and retry, source identity, signatures, and exact payload
restoration. One exact pre-existing ordinary-table expectation was excluded:
`malformed_ordinary_table_vectors_fail_in_the_base_workbook_layer`.

The broader locked XLSB library/test gate passed 15 suites and `726/726`
executed tests with exactly two pre-existing tests in unmodified code
excluded: `malformed_ordinary_table_vectors_fail_in_the_base_workbook_layer`
and `checked_in_unique_standard_drawing_corpus_transfers_every_anchor`.
The former expects stale base-parser context strings. The latter expects six
anchors while the checked-in corpus produces five; the test and all four
fixtures are byte-identical to `HEAD`. Neither failure is a Change 0376
regression.

Strict production-library Clippy passed with `-D warnings` and no allowances.
The crate-boundary gate passed for 64 workspace packages and 240 internal
dependency declarations with 14 explicit existing debt entries. Independent
topology, publication-safety, and test reviews accepted the bounded
implementation.

## Claim boundary

`performance_claim: none`; `claim_authorized: false`. This batch proves the
specific canonical first-create/final-remove lifecycle, topology,
preservation/refusal, atomicity, source-identity, signature, and inverse
invariants exercised above. It includes no benchmark, latency,
allocation-volume, RSS, physical-I/O, cold-cache, throughput, fixed-memory,
broad XLSB, or system-level OOM measurement. No general OOM-prevention claim
follows.
