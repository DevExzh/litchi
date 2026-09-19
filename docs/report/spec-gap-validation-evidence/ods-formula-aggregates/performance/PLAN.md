# ODS aggregate evaluator performance profile

This profile measures the bounded read-only evaluator batch for the seven
numeric sequence and sum-of-squares functions `SUM`, `PRODUCT`, `SUMSQ`,
`SUMPRODUCT`, `SUMX2MY2`, `SUMX2PY2`, and `SUMXMY2`.

The candidate lane is built from the source freeze selected by the root agent.
The baseline lane is committed `2aeeb8d2f`, which already contains the scalar,
matrix, complex, database, rounding, trigonometric, and elementary evaluator
paths but does not implement these seven aggregate functions.  Baseline timing
is therefore limited to matched controls that exist in both revisions.  A
baseline refusal for a new aggregate is not timed or presented as a regression
comparison.

Each lane uses a fresh release child process for every sample, three warmups
inside that process, and fifteen measured samples per case and phase.  The
runner records elapsed time, process RSS, allocator calls and bytes, peak live
bytes, evaluator work and retained execution-budget memory, resolver reference
reads, and a checksum of the result.  Repeats are fixed by shape so the report
normalizes scalar and array work without hiding the raw receipt.

The workload has four purposes:

* scalar calls and small literal-array calls exercise the ordinary aggregate
  kernels, including the scalar definitions of all seven functions;
* streamed local references exercise the resolver-backed sequence path without
  materializing a separate cell object per input;
* 16, 64, 256, and 1024-row by four-column references test scaling and the
  large-reference resource boundary;
* a three-matrix `SUMPRODUCT` is included at 4×4 literal and 1024×4 reference
  geometry so the K>2 scaled-product accumulator is exercised without
  expanding the whole matrix;
* nested projected-invariant calls such as
  `SUMPRODUCT(range+SUM(range);range)` and
  `SUM(IF(range;SUMPRODUCT(range+1;range);0))` test that invariant scalar and
  branch aggregates are projected once rather than rebuilt for every outer
  cell.  These rows are scaling evidence only; they do not claim worksheet
  recalculation or production adapter behavior.

The direct oracle in the harness computes the expected finite `f64` result from
the fixture values.  It validates the selected evaluator result and shape
before timing; it is not an independent floating-point implementation.  The
profile does not measure save, recalculation, native producer acceptance,
cross-platform bit identity, or a language-level resident-memory guarantee.
