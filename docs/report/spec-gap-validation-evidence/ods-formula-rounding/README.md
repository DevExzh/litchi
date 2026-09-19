# OpenFormula rounding evaluation

The ODS scalar and value evaluators implement the eight OpenFormula 1.4
Part 4 §6.17 functions: `CEILING`, `INT`, `FLOOR`, `MROUND`, `ROUND`,
`ROUNDDOWN`, `ROUNDUP`, and `TRUNC`. The value evaluator uses the same
kernels for elementwise arrays, broadcasting, and local references.

Callers use the existing `evaluation::evaluate_scalar` or
`evaluation::value::evaluate` APIs with explicit execution contexts and limits.
There is no new public numeric type, ambient service, dependency, or workbook
mutation. Evaluation does not update formula caches, spill cells, or recalculate
dependencies. Complete OpenFormula conformance remains open.

## Conversion and resources

The finite `f64` profile retains the existing locale-independent Number
conversion and truncation-toward-zero Integer conversion. `ROUND` accepts a
Number exponent; `ROUNDDOWN`, `ROUNDUP`, and `TRUNC` convert their digit
arguments to Integer. Fractional `ROUND` exponents follow the specification's
typed signature and explicit power expression; this is an interpretation of
the conflict with its claim that nonpositive digits always yield an integer.
Supplied formula errors propagate before conversion errors. Missing optional
syntax receives the function's default; an Empty referenced cell receives
ordinary Number conversion to zero.

Integer digit counts quantize the input's shortest round-trip decimal
representation with checked coefficient arithmetic, then convert the result
back to `f64`. This keeps ordinary decimal multiples stable and distinguishes
adjacent floating-point inputs without epsilon snapping. Fractional `ROUND`
exponents use binary floating-point scales; significance-based rounding also
retains binary arithmetic. These are explicit numerical profile choices.

The rounding arithmetic uses fixed-size scalar state. Existing evaluator
stacks, array outputs, text conversion, and reference reads retain their work,
memory, geometry, and cancellation checks. The arithmetic kernels do not
allocate, but a complete evaluator call may allocate its ordinary VM state.
No claim of zero-allocation end-to-end evaluation is made.

`INT` rounds toward negative infinity. `ROUND` rounds ties away from zero;
`MROUND` chooses the greater numerical multiple on a tie. `CEILING` and
`FLOOR` preserve the specification's sign-compatible significance and optional
mode behavior. Zero `MROUND` multiples produce `#DIV/0!`; selected results
outside the finite Number domain produce `#NUM!`.

## Evidence

The [contract](contract.md) and [independent review](review.md) identify the
normative source and numerical profile. The public integration suites are
`ods_formula_rounding_evaluation` and `ods_formula_rounding_arrays`.
The [performance profile](performance.md) retains its harness, raw measurements,
and replay commands. Measurements apply only to their named corpus and build;
they do not establish native producer acceptance or whole-workbook performance.

Final root validation passed 1,168 ordinary ODS tests across 70 targets and
five doctests, including all 28 focused rounding tests. Strict all-target
Clippy, warning-denied rustdoc, and crate formatting passed. The
[gate receipt](gates/receipt.json) records exact commands, toolchain, log hashes,
and unchanged ODS source hashes before and after all four gates. The independent
[workspace boundary check](gates/boundaries.log) passed for 65 packages and
241 internal dependency declarations; this batch adds no dependency.

This completes the bounded rounding family batch, not the full specification
gap audit or the workspace-wide performance program.
