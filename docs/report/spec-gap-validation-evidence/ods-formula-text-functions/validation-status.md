# ODS ordinary text functions

This batch implements the 26 ordinary text functions in OpenFormula 1.4
§6.20. The seven byte-position functions in §6.7 remain a separate open
family. Evaluation is read-only: it neither recalculates dependencies nor
publishes workbook formula caches.

The [contract](contract.md) owns the function signatures, Unicode and
formatting choices, argument conversions, matrix behavior, and formula-error
rules. The [baseline](baseline.json) identifies the preceding committed
implementation and the isolated dependency lock used for validation.

The implementation uses the existing scalar evaluator and value bridge.
Matrix calls retain reference geometry and select borrowed cells for each
output coordinate, rather than materializing an input range. Output arrays,
owned strings, and search state remain subject to the caller's storage and
work limits. Provider, resource, cancellation, and source-version failures
remain typed evaluator failures.

Unicode properties are generated from pinned official Unicode data; the
[generator](unicode-data/generate.py) and [provenance](unicode-data/provenance.json)
record reproducible inputs and licensing. The [native evidence](native/README.md)
retains producer observations separately from the normative contract.
The [independent text oracle](oracle-review.md) covers all 26 functions,
including positional numeric placeholders and improper versus mixed fractions.
The [fraction verifier](verify_fraction.py) separately compares the exact
bounded-denominator helper with Python `Fraction` across 5,474 deterministic
cases; its [receipt](fraction-reproduction.json) pins both source and verifier.
This helper check does not establish formatter integration or a universal
numerical proof.
The [performance plan](performance/PLAN.md) defines matched existing controls,
new-function workloads, allocation/read accounting, and source custody.

The first frozen candidate passed seven isolated gates and 1,592 tests, but
independent review subsequently found three uncovered formatter grammar
cases. Those [superseded receipts](diagnostics/pre-grammar-fix/README.md) are
retained separately. The corrected grammar and 244-observation oracle now pass
the [gate verifier](gates/verify.py) under a new source freeze: all seven
gates and 1,592 tests, with zero failures or ignored tests.
[Root follow-up](grammar-closure.md) records the fixes and review attribution.

Validation is complete for this batch; see [completion.md](completion.md)
for the final disposition, review attribution, performance costs and limits.
The ambient workspace Cargo.lock must not replace the retained isolated
[gate lock](gates/Cargo.lock).
