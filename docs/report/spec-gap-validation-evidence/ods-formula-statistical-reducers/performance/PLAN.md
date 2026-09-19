# ODS statistical reducer performance plan

The profile follows ADR 0005's matched-control measurement contract. It keeps
the production source closure, harness source and lockfile, toolchain identity,
typed fixture formulas, raw child receipts, source manifests, and cleanup
receipts independently checkable.

The baseline lane captures only functions present in both revisions:
arithmetic, `SIN`, `IMSUM`, `DSUM`, `SUM`, and `SUMIFS`. The candidate lane
captures the same controls plus the nine statistical reducers. Each statistical
function has scalar, literal-array, reference, ordered-list, 3-D, empty,
formula-error, 64-row nested-projection, and typed resource-refusal rows.
`AVERAGE`, `COUNTA`, and `COUNTBLANK` add 256-row and 1024-row nested
projections for scaling evidence, producing 87 statistical rows alongside 13
controls.

Reference cases use a borrowing resolver with deterministic numeric, mixed,
empty, and formula-error lanes. Resolver read counts are compared with the
known rectangle/list/3-D/nested bounds, so a repeated projected reducer is
visible in the receipt. The independent 576-row Fraction oracle and retained
native cached-results/provenance receipts are included in the selected source
inputs.

Each final group uses three untimed warmups and fifteen fresh measured child
processes in both `evaluate` and `parse-evaluate` phases. The root agent must
freeze the selected source closure and authored profile inputs before capture;
no timing result is valid if those hashes change.
