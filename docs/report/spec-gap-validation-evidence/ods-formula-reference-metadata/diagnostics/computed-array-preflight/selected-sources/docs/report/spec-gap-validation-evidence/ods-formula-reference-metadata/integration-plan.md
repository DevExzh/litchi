# Reference-metadata implementation and validation plan

Baseline: `049c09cdde3978593149079c4257df047a3fa419`.
Scope: AREAS, COLUMN, COLUMNS, ISREF, ROW, ROWS, SHEET, SHEETS.
The normative contract is recorded independently in `contract.md`.

## Ownership and implementation

The shared private catalog owns names and arities. Scalar evaluation admits
only semantics for which it has context; it must not invent a workbook or
current position. The value evaluator uses the existing explicit Resolver and
Position. Its reference-metadata module owns descriptor admission and output.
No ambient services, global cache or new public workbook metadata model is
introduced. Unrelated working-tree changes remain outside the batch.

Reference operations preserve runtime identity before implicit intersection.
AREAS counts retained reference records, not expanded sheet planes. ISREF
inspects the runtime kind before value coercion. Both accept reference lists;
the six other functions reject list identity even when intersection leaves
only one retained record. Dimension and sheet functions
consume reference geometry/order without reading cells. ROW/COLUMN output uses
checked array capacity and per-output work. Explicit references return the first
generated axis element in scalar mode; omitted arguments use the fixed formula
position. Computed arguments retain the enclosing calculation mode and complete
arguments where required by projected matrix branches. Cache classification
must prove position independence; MUNIT scalar descendants remain sensitive to
the current position. Formula errors and typed provider/resource/source failures
remain distinct.

## Required evidence

- Independent specification contract and semantic review.
- Focused semantic tests for all eight functions, arities, formula errors,
  current coordinates, 2-D/3-D references, reference lists, arrays and composed
  scalar/matrix/lazy expressions.
- Focused resource tests for zero cell reads, bounded descriptor/output storage,
  output/work limits, cancellation, metadata provider failures and source fences.
- Independent executable oracle and native fixture observations; compatibility
  differences documented without replacing the normative contract.
- Frozen selected source hashes and isolated seven-gate run: package tests,
  strict all-target Clippy, strict rustdoc, package and batch formatting,
  crate boundaries and diff checks.
- One coordinated before/after performance capture with immutable sources,
  preflight semantic/read assertions, matched existing controls, new metadata
  cases, raw samples, allocation/work/read metrics and process RSS. Review each
  regression flag and report limitations; do not claim universal improvement.
- Final source/evidence verification, owned temporary-file cleanup and batch
  commit. Update the broader audit without staging unrelated report content.

Preparation files do not constitute a source freeze or a passing gate receipt.
