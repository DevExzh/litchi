# Matrix-function performance plan

This plan covers `MDETERM`, `MINVERSE`, `MMULT`, `MUNIT`, and `TRANSPOSE` under
the contract in `ods-formula-matrix-functions/contract.md`. Matrix arguments
must use the functions' `ForceArray` semantics. The corpus will use numeric
elements first, with logical/text/empty conversion cases added only after the
implementation policy is recorded; the specification leaves those conversions
partly implementation-defined. Singular `MINVERSE`, shape errors, and
resource refusals remain distinct typed outcomes.

| family | bounded corpus | scaling or oracle |
|---|---|---|
| square cubic | `MDETERM`, `MINVERSE`, and square `MMULT` at `n=1,2,4,8,16`; `n=32` stress only when the work/storage limits allow | Dense diagonal-dominant integer matrices with known determinant or residual check. Arithmetic is expected to expose O(n³); `MINVERSE` retains O(n²) output cells. |
| output quadratic | `MUNIT` and `TRANSPOSE` at `n=1,4,16,64` (`64²=4096`); rectangular transpose at 2×8, 8×2, and 16×64 | Check shape, every element, checksum, output reservation, and O(n²) traversal/materialization separately from evaluator setup. |
| rectangular multiply | `(1×n)(n×1)`, `(n×1)(1×n)`, `(4×8)(8×2)`, and `(16×32)(32×4)`; square controls use the cubic family | Validate `m×k · k×p → m×p`, with operation count O(mkp), output cells O(mp), and deterministic row-major values. |
| errors and laziness | Singular 2×2 and 8×8 `MINVERSE`; non-square `MDETERM`/`MINVERSE`; incompatible `MMULT`; `MUNIT(0)`, negative, and non-integer inputs; each function behind an unselected `IF(FALSE(); expensive; 0)` branch | Formula errors must not be reported as resource failures. Selected/unselected pairs must prove that the skipped matrix does not consume reads, work, output reservations, or allocations. |

Use exact small matrices for differential oracles and diagonally dominant
integer matrices for larger cases. For inverse results, verify bounded
`A·A⁻¹` residuals and shape rather than comparing formatted decimal text.
Record the chosen singular and `MUNIT` non-integer policies before timing.
Include `IF(TRUE(); expensive; 0)` selected controls, plus `IFERROR`/`IFNA`
unselected error branches where supported, so branch laziness is measured
with both successful and failing operands.

Each case needs separate `setup`, `parse`, `evaluate`, and `parse-evaluate`
lanes. Build the expression and matrix fixture outside the evaluate timer;
retain a parse-evaluate lane for caller-visible cost. Use the existing value
harness accounting shape: p50/p95/p99 batch time, explicit repeat, work used,
allocator calls and requested/released bytes, live and peak-live deltas,
retained execution-budget bytes, output element count, checksum, typed failure,
and `/usr/bin/time -v` maximum RSS. Add copied-byte accounting where the
implementation exposes it; heap counters alone must not be called proof of
zero-copy behavior. The result reservation for `MINVERSE`, `MMULT`, `MUNIT`,
and `TRANSPOSE` should be reported separately from temporary arithmetic
storage.

For each family add bounded refusals with maximum array cells, maximum work,
and maximum storage below the successful case. Add pre-cancelled and
mid-operation cancellation cases for cubic loops, and verify cancellation
checks stop before unbounded output allocation. Shape validation should occur
before allocation; record whether a refusal used zero output cells and its
work/accounting counters.

The current before implementation reports `UnsupportedFunction` for these
calls, so it cannot serve as a semantic speed baseline. Capture candidate-only
absolute results with the no-op/constant harness controls and explicit source,
binary, compiler, CPU-6, environment, and corpus hashes. Once a comparable
implementation exists, replay the identical corpus in serial AB/BA/AB
processes, archive every raw row, and apply the GOAL review trigger for any
approximately 5% latency or peak-RSS regression or material scaling loss.
Report every individual result; do not use unsupported-call timings or a
geometric mean to imply a matrix-function speedup.
