# ODS descriptive-statistics evidence

This batch implements `AVEDEV`, `DEVSQ`, `GEOMEAN`, `HARMEAN`, `KURT`,
`SKEW`, and `SKEWP` in the explicit scalar and value formula evaluators.
The baseline and unrelated tracked edits are recorded in `baseline.json`.
Formula caches and document CRUD remain inert; this is not a dependency
recalculation engine or a complete OpenFormula evaluator.

The [contract](contract.md) records normative signatures, scalar/reference
conversion, list admission, cardinality, signed zero, and numerical choices.
`DEVSQ` and `SKEWP` reject reference lists before resolver reads. Other
reference sequences are streamed in source order, with typed provider,
resource, cancellation, and source-version failures preserved.

Centered reducers use fixed-size exact dyadic power sums. `DEVSQ`, `KURT`,
`SKEW`, and `SKEWP` finish in one reference pass; `AVEDEV` replays borrowed
sequence descriptors with cumulative cell and work charges. The geometric
kernel uses normalized products and a wide exponent ledger. Both mean
functions admit signed inputs according to their real-valued domains.
Harmonic evaluation first accumulates bounded reciprocal significands and an
outward error bound. Uncertain cancellation replays into exact rational state;
reservations follow actual limb growth, including transient allocation peaks,
and work charges follow occupied precision. No reducer retains input cells.

The independent [numeric oracle](numeric_oracle.py) uses exact rational
centered moments and high-precision Decimal transcendental calculations.
Its [retained corpus](numeric-goldens.json) covers extreme finite inputs,
cancellation, subnormal skew, typed cells, and list admission. Ordinary exact
rows and sensitive finite rows retain their declared zero- or eight-ULP bounds;
the subnormal skew regression also checks exact result bits.

The [native evidence](native/README.md) retains 39 finite observations across
all seven functions from seven pinned LibreOffice FODS files. The root-owned
[reproduction receipt](native-reproduction.json) records byte-for-byte
regeneration, source hashes, and temporary-tree cleanup. These cached values
are independent fixture observations, not a native recalculation performed
by this batch.

The isolated scripts under `gates/` stage selected production modules, tests,
oracles, and contracts before recording a freeze manifest. They run seven
integration gates and record source identity before and after execution.
The gate lock comes from the completed order-statistics batch; the ambient
root lock remains untouched. The [performance plan](performance/PLAN.md)
compares 24 matched controls and 91 descriptive cases, with three warmups and
15 fresh process samples for each evaluator phase.

Run `python3 docs/report/spec-gap-validation-evidence/ods-formula-descriptive-statistics/verify.py`
to verify retained acceptance evidence. It independently checks gate commands
and logs, frozen inputs, oracle reproduction, pinned native reproduction, raw
performance samples, read counts, source custody, and reported medians.
Acceptance requires all checks to succeed; implementation or corpus presence
alone does not establish a passing gate or performance result.
