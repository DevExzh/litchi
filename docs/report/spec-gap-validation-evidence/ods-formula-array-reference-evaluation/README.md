# Bounded OpenFormula value and reference evaluation

The opt-in `codec::formula::evaluation::value` API extends the constant scalar
profile with local references and arrays. This work follows the ODS findings
in the specification audit; it does not complete OpenFormula or add workbook
recalculation. The [public module documentation](../../../../crates/litchi-ods/src/codec/formula/evaluation/value.rs)
contains the compiled worksheet/owned-result example. The
[feature matrix](../../../../crates/litchi-ods/docs/FEATURE_MATRIX.md) distinguishes
this implemented profile from the remaining calculation features.

## Public contract

- `evaluate_scalar` remains the resolver-free constant scalar API.
- `value::evaluate` accepts a parsed expression, immutable resolver, caller
  position, mode, execution context, and finite limits. Matrix mode preserves
  first-class references and broadcasts consuming operators over arrays;
  scalar mode uses implicit intersection.
- Local range, intersection, and union operators retain reference/list kinds,
  ordering, and duplicate semantics. Lazy conditionals evaluate selected demand
  cells without resolving unselected references. Formula errors remain distinct
  from cancellation, budget, provider, and unsupported-capability failures.
- `worksheet::formula::Resolver` indexes physical repetition runs without
  expanding the grid. The caller supplies finite dimensions and may reuse the
  borrowed index across evaluations. Formula caches are never authoritative.
- `Evaluated::to_owned` copies result storage under a separate caller budget;
  it does not read or evaluate reference contents. Borrowed results remain
  available when the copy fails.

Names and labels, external sources, unimplemented function families,
dependency scheduling, recursive formula-cell evaluation, cache publication,
spilling, volatile/data-table behavior, and workbook writes remain unsupported.
A first-class 3D reference/list view is not a promise that every consumer can
flatten it into a rectangular array. Whole-row/column references use the explicit
finite resolver extent, not an invented universal spreadsheet size.

## Validation and review

The latest runtime implementation checks are retained in
[scalar-cell-candidate-03.json](gates/scalar-cell-candidate-03.json), with the
exact 455-file source closure and adjacent log archive. All five scoped gates
passed, including the 66 array/reference integration tests. The
[scalar-cell review](gates/scalar-cell-review.md) covers deferred cell reads,
publication/copy guards, immutable-source fences, and demand specialization.
The later [documentation gate](gates/value-documentation-01.json) includes the
public example and current source hashes.

The integration suite also covers borrowed text, physical worksheet order,
list/array distinction, lazy branch exclusion, finite admission, cancellation,
source-version changes, retained reservations, and owned conversion. Historical
failed drafts are retained as diagnostics rather than presented as passing
production evidence.

## Performance disposition

[Two paired captures of the specialized scalar path](performance/scalar-cell-specialized-analysis.md)
retain 162 comparisons with no deterministic mismatches. The 4,096-cell scalar
chains reduced allocations from 12,304 to 16 and requested bytes from 2,163,192
to 524,792 per measured batch. These are bounded in-memory evaluator scenarios,
not an end-to-end Office workflow claim. Individual latency and process-RSS
review flags remain visible; performance acceptance is still open.

The [capture runner guide](performance/README.md) describes the broader scalar,
value, and worksheet lanes. Raw evidence is archived before loose intermediate
files are deleted. Retained baseline/candidate executables stay on SSD for
reproducible comparisons; repository or build copies are not created in tmpfs.
