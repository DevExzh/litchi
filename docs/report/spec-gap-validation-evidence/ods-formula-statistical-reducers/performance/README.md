# ODS statistical reducer performance evidence

This directory retains a reproducible process profile for the nine bounded
OpenFormula statistical reducers `COUNT`, `COUNTA`, `COUNTBLANK`, `AVERAGE`,
`AVERAGEA`, `MIN`, `MAX`, `MINA`, and `MAXA`.

The matched baseline is the conditional reducer revision
`f7fe857007b7bcb65e0512c5b5caf9135fe74a5a`. It supplies arithmetic, `SIN`,
`IMSUM`, `DSUM`, `SUM`, and `SUMIFS` controls. Statistical rows are candidate
only because that baseline does not implement these nine functions.

The candidate matrix has 13 matched controls and 87 statistical rows. Each
reducer has scalar and inline-array arguments, borrowing references, ordered
reference lists, one 3-D reference, empty selections, formula-error cells, a
64-row nested scalar projection, and a zero-reference-cell resource refusal.
`AVERAGE`, `COUNTA`, and `COUNTBLANK` also have 256-row and 1024-row nested
projections to expose scaling. Every timed child performs an untimed
correctness preflight against the deterministic typed fixture, then records
elapsed time, peak RSS, allocator calls and bytes, live-byte balance, evaluator
work, retained budget memory, resolver reads, and a result checksum.

Final baseline and candidate captures require the root agent's explicit quiet
source-freeze window. Preparation, locked release builds, and bounded
preflight may run before that window. The profile makes no save, recalculation,
cache-publication, native-producer, cold-filesystem, or cross-platform timing
claim.
