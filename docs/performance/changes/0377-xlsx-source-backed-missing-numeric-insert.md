# Change 0377: XLSX source-backed missing numeric insert

## Scope

Change 0377 closes the create verb for the guarded XLSX source-backed
cell-value owner. `CellValueEdit::Insert`, selector-first
`SheetCellValueEdit::insert`, and the single/multi-sheet convenience methods
create a finite numeric cell only where no physical `<c>` owner exists.
`Set`, `SetFormula`, `SetSharedFormula`, `Clear`, and
`Remove` retain their existing-owner meanings.

The operation covers missing positions in existing rows and missing rows,
with deterministic physical ordering, read-after-create, and expansion of an
existing worksheet dimension. It does not create missing strings, dates,
formulas, shared strings, styles, relationships, worksheet Parts, or arbitrary
cell types.

## Architecture and behavior

`litchi-xlsx` remains the semantic owner. `Snapshot` proves that the
target has no physical entry and is outside every merge, array/data-table, or
shared-formula range. The existing guarded source closure continues to reject
tables, worksheet relationships, MCE/opaque markup, unknown cells, cell
metadata, and unsupported protection/encryption. Existing explicit-empty and
style-only cells are owners and therefore cannot be inserted over.

`Number` keeps its validated finite lexical representation. Its write
validation now applies the existing 32,767-character SpreadsheetML cell
payload ceiling, matching the raw worksheet reader rather than adding a
source-local magic limit. Atomic batches retain `MAX_BATCH_EDITS` and
reject duplicate or later-invalid requests before staged state changes.

The existing raw `Action::set` path writes missing cells/rows in coordinate
order and expands an existing dimension. Every effective insertion uses the
existing calculation-property invalidation and safe calculation-chain removal
path. Publication remains source-bound, clone-staged, fully reparsed,
failure-atomic, and signature-aware. Forward patching proves the absent-to-
numeric transition; inverse application restores the exact original package
bytes and physical cell absence.

## Verification

All Cargo work used one process at a time, `CARGO_BUILD_JOBS=1`, one
dedicated target, disabled incremental state, one test thread, an
available-memory launch guard, and a 6 GiB per-process virtual-memory cap.
These are operational safeguards, not OOM-prevention evidence.

The focused `source_backed_cell_values` integration target passed
`59/59`. It covers existing-row holes, missing rows, row/cell ordering,
dimension expansion, single/multi-sheet selectors and publication, exact
numeric length boundaries, physical-owner and formula-range refusal, merge
source refusal, duplicate/later-invalid atomicity, calculation invalidation,
deterministic output, untouched package topology, signed mutation refusal,
exact no-op behavior, and forward/inverse restoration of physical absence.

The broader locked XLSX library/integration gate passed 58 suites and
`1213/1213` executed tests with exactly four pre-existing tests in
unmodified `source_backed_row_visibility` code excluded:

- `changed_publication_reuses_matched_provenance_without_selected_reload`
- `managed_changed_publication_preserves_unknown_members_and_releases_budget`
- `managed_signature_noop_and_changed_protection_contracts_remain_fail_closed`
- `unsafe_xml_macros_and_signatures_fail_closed`

The first conflates one semantic replacement check with one physical
`ReadAt` callback. The next two omit the established physical-member lookup
reservation and expect later preservation/signature outcomes rather than the
earlier resource limit. The fourth expects scalar-formula worksheets to be
rejected even though the shared cell-value closure intentionally accepts
them. Static causal review confirmed that none reaches `Insert` or the new
absence helper.

XLSX doctests passed `2/2`. Production-library Clippy passed with one
named pre-existing `clippy::useless_asref` allowance in unmodified
hyperlink snapshot code. The crate-boundary gate passed for 64 workspace
packages and 240 internal dependency declarations with 14 explicit existing
debt entries. Independent production/API, safety, test, and final static
reviews accepted the bounded implementation.

## Claim boundary

`performance_claim: none`; `claim_authorized: false`. This batch proves the
specific numeric absent-owner create/read/inverse lifecycle, range and source
refusals, lexical limit, ordering/dimension, calculation invalidation,
atomicity, source identity, signature, and preservation invariants exercised
above. It includes no benchmark, latency, allocation-volume, RSS,
physical-I/O, cold-cache, throughput, fixed-memory, broad XLSX, or system-level
OOM measurement. No general OOM-prevention claim follows.
