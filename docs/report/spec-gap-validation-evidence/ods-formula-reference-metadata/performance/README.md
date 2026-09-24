# ODS reference-metadata performance harness

This directory owns the before/after process harness for `AREAS`, `COLUMN`,
`COLUMNS`, `ISREF`, `ROW`, `ROWS`, `SHEET`, and `SHEETS`. The baseline commit,
independent contract inputs, and native observations are owned by the parent
batch directory. The retained performance lock is the isolated gate copy whose
SHA256 is recorded in `baseline.json`.

The case matrix has 28 matched controls and 108 metadata cases; the candidate
lane repeats the 28 matched controls for 136 timed cases. Each child
records elapsed time, allocator calls and bytes, peak live bytes, execution
work, resolver cell reads, input/output bytes, result checksum, and external
RSS. The candidate preflight validates exact metadata values and projected
coordinates before any timed row. Metadata functions are expected to inspect
reference descriptors or provider sheet metadata and therefore produce zero
cell reads; a resource cap or cancellation is recorded as a typed evaluator
failure. The computed `SHEET(ABS(range))` lane records the one scalar-
intersection read performed by `ABS` before the metadata operation.
Source-qualified references, selected source branches, and `IFERROR`/`IFNA`
source preservation and refusal are checked with zero cell reads; source
arithmetic remains a typed unsupported-reference failure. Failed-first-operand
`IFERROR`/`IFNA` fallbacks preserve source policy, while selected array source
branches remain typed unsupported failures in either cell order. Singleton
`ReferenceList` refusal is covered for all six functions whose syntax requires
a direct `Reference`. Projected computed `IF`, `IFERROR`, and `IFNA` cases cover
SHEET text arrays and the ROW/COLUMN value-array refusal with exact shapes and
read counts. The six computed ROW/COLUMN lanes each read exactly two selected
child cells before the value-array refusal.

Preparation includes no timing capture. Capture ownership is singular: after
root sends the frozen source handoff, run the command in `PLAN.md` once and
retain raw child receipts. If the run fails, preserve its output under a
`diagnostic-*` directory before any retry.
