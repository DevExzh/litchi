# ODS conditional aggregate performance evidence

This directory retains the reproducible process profile for the six
reference-only ODS conditional aggregate functions `SUMIF`, `SUMIFS`,
`COUNTIF`, `COUNTIFS`, `AVERAGEIF`, and `AVERAGEIFS`.

The baseline is the committed revision
`5b125e9ea5870a56b2648ad02917894a5549c571`. It supplies nine matched
arithmetic, `SIN`, `IMSUM`, `DSUM`, and `SUM` controls. The candidate lane
contains those controls, three additional array controls, and the conditional
rows, nested projected reducers, and bounded formula-error rows.
The conditional cases are evaluated through the resolver-aware value API;
constant range arguments are retained only as a deliberate error case.

Every final group uses three untimed warmups and fifteen fresh measured child
processes in both `evaluate` and `parse-evaluate`. Receipts retain elapsed time,
GNU-time maximum RSS, allocator calls and bytes, balanced live bytes, peak live
bytes, evaluator work, retained execution-budget memory, resolver cell reads,
and a result or formula-error checksum. The resolver is an immutable borrowing
fixture with deterministic numeric and text criterion lanes.

The 33-case matrix covers omitted and explicit destination ranges, one and two
criteria, exact text criteria, ordered reference lists, one 3-D geometry, a
one-cell anchored destination clipped at the sheet extent, empty selection,
mismatched geometry, a reference-cell refusal, and 64/256/1024-row projected
`SUMIFS` reducers under an outer array `IF`. The projected rows expose whether
the scalar reducer and its selected destination are evaluated once per formula
or rebuilt for every output cell.

Final baseline and candidate captures require the root agent's quiet
source-freeze window. Preparation, build, and preflight are allowed before that
window. This profile makes no save, recalculation, cache-publication, native
producer, cold-filesystem, host-wildcard, or cross-platform timing claim.
