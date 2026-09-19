# Bounded OpenFormula database evaluation

The value evaluator implements the twelve OpenFormula §6.9 functions:
`DAVERAGE`, `DCOUNT`, `DCOUNTA`, `DGET`, `DMAX`, `DMIN`, `DPRODUCT`,
`DSTDEV`, `DSTDEVP`, `DSUM`, `DVAR`, and `DVARP`. The resolver-free scalar
evaluator does not implement this family. The [contract](contract.md) identifies
the local normative source and distinguishes specification requirements from
the following repository choices.

## Inputs and selection

Database and criteria arguments retain their complete rectangular shape in
scalar and projected matrix contexts. Inline arrays and single contiguous
single-sheet references are supported. Reference lists and 3-D tables remain
typed unsupported inputs. The first database row contains headers, followed by
zero or more records. Criteria require a header and at least one body row.

Positive integer field selectors address physical columns, including columns
with duplicate, non-text, Empty, or Error headers. Text selectors compare
Unicode lowercase character sequences without allocating normalized strings;
missing or ambiguous matches produce `#VALUE!`. A unique empty Text header is
selectable. Criteria headers follow the same field-selector rules; a genuinely
Empty header is rejected rather than silently becoming a match-all condition.

Criteria within a row are AND-ed in column order; rows are OR-ed. False clauses
and successful alternatives short-circuit. Each required record field is read
at most once per record and reused across clauses. A selected aggregate field
is read only after the record matches. Criterion Error values propagate during
preparation; errors in inspected record criteria propagate during matching.
Uninspected record fields do not affect the query. Provider Unsupported values
remain typed evaluation failures; cached formulas are never recalculated.

Text criteria use case-sensitive, whole-cell literal matching. Whitespace is
significant. Wildcards, regular expressions, substring matching, ambient locale,
and textual date conversion are disabled. Only non-empty Text criteria with an
explicit operator prefix (`=`, `<>`, `<`, `<=`, `>`, or `>=`) use the existing
locale-independent finite-number parser. A bare numeric-looking Text criterion
remains Text and therefore matches Text candidates rather than Number
candidates. Already numeric date/time serials compare as Numbers. An Empty
criterion reference converts to numeric zero; explicit `"="` matches Empty
records, and `"=0"` does not match Empty records.

The shared database/conditional matcher regression for this distinction is
recorded in the [criterion text profile evidence](criteria-text-profile.md).

## Results and arithmetic

Numeric aggregates omit Text, Logical, and Empty selected cells and propagate
selected errors. `DCOUNT` counts Numbers and ignores selected errors; `DCOUNTA`
counts all nonempty cells, including errors and empty Text. Both accept an
explicit missing middle field (`DCOUNT(database;;criteria)`) and the two-argument
omitted-field form; either counts matching records.

`DGET` requires exactly one matching record and returns Number according to
§6.9.5: Empty becomes 0, Logical becomes 0 or 1, finite numeric Text is parsed,
and unconvertible Text or Complex becomes `#VALUE!`. Zero or multiple matches
also produce `#VALUE!`.

Empty selections return 0 for `DSUM`, `DMIN`, and `DMAX`, and 1 for `DPRODUCT`.
`DAVERAGE` and population statistics require one Number; sample statistics
require two. Insufficient counts produce `#VALUE!`. Nonfinite final numeric
results produce `#NUM!`.

Sums use fixed-size exact binary accumulators and round once at the result
boundary; averages divide that accumulator before rounding. Products retain a
scaled mantissa and exponent. Variance uses compensated moments around the
original first value, scaled to avoid overflow; standard deviation can remain
finite when the corresponding variance overflows. Each function updates only
its own numeric kernel. These choices preserve finite cancellation, subnormal
means, and adjacent large floating-point differences without per-record
allocation; they do not promise exact real arithmetic for all statistics.

## Resource and ownership boundaries

Headers and compiled criteria borrow input text. Only headers, criteria clauses,
required-column indexes, and one row of optional cached values use temporary
vectors, with fallible capacity reservations held until after each allocation
is dropped. Provider text obeys the configured text-length ceiling. Work,
geometry, cumulative reference reads, and storage remain bounded by the caller's
evaluation/execution limits. Provider cancellation checks remain active during
reads; errors do not publish a partial result.

Coordinate-independent results can be cached inside projected lazy expressions.
The public result is a scalar Number/Error; existing owned-result conversion
therefore needs no new database-specific public type or retained table copy.

See the [evidence status](README.md) for current tests and remaining performance
acceptance. This family does not establish complete Small Group conformance or
dependency recalculation.
