# ODS paired statistics and simple regression

This batch implements `CORREL`, `COVAR`, `PEARSON`, `RSQ`, `SLOPE`,
`INTERCEPT`, `STEYX`, and `FORECAST` in the bounded read-only OpenFormula
evaluators. It does not add dependency recalculation or publish workbook
formula caches. The baseline is recorded in [baseline.json](baseline.json).

The [contract](contract.md) records the local ODF 1.4 source hashes, exact
function scope, argument shapes, aligned omission, error precedence, and
resource rules. In particular, it documents the contradictory INTERCEPT
cross-reference and explicitly selects regression with an included constant.
It also distinguishes RSQ's `#N/A` count mismatch from a `#VALUE!` orientation
or pseudotype refusal. Modern spreadsheet aliases are outside this scope.

Paired reference data stream without an input-cell vector. Fixed-width dyadic
sums retain exact centering and cancellation; only final publication rounds
to binary64. FORECAST retains an exact invariant fit while evaluating each
scalar query at its own projected position. Formula errors remain values;
provider, allocation, budget, cancellation, and source failures remain typed
evaluator failures. A shared dyadic helper also gives descriptive reducers
direct rounding at the subnormal target quantum.

The evidence is separated by purpose:

- [Independent numeric oracle](oracle-review.md), [generator](numeric_oracle.py),
  and [retained goldens](numeric-goldens.json).
- [Pinned native observations](native/README.md) and
  [extraction reproduction](native-reproduction.json). Native caches are
  corroborating observations, not the semantic authority. One retained STEYX
  cache differs by 2,005 ULP; its [independent proof](native/steyx-deviation-proof.json)
  preserves the host observation and verifies the exact mathematical reference.
- [Semantic and numerical review](review.md) and
  [resource/cache review](resource-review.md).
- [Frozen inputs](gates/freeze.json), [isolated gate results](gates/results.json),
  and [source-stability checks](gates/verification.json).
- [Performance scope and receipts](performance/README.md), including matched
  existing descriptive reducers affected by the shared helper.
- [Batch verification](verification.json) and
  [owned scratch cleanup](scratch-cleanup.json).

The isolated gates use the retained [Cargo.lock](gates/Cargo.lock); the ambient
workspace lock is not substituted for that frozen input. Reproduce the
independent corpus with `python3 numeric_oracle.py --check` from this evidence
directory. The four focused Rust targets are `ods_formula_paired_evaluation`,
`ods_formula_paired_limits`, `ods_formula_paired_oracle`, and
`ods_formula_paired_native`; the full isolated package gates also cover the
existing descriptive-statistics consumers of the shared helper.
