# ODF 1.4 paired-statistics and simple-regression evaluator contract

This contract defines the bounded implementation profile for the eight paired
statistics and simple linear regression functions in OpenFormula 1.4 Part 4:
`CORREL`, `COVAR`, `PEARSON`, `RSQ`, `SLOPE`, `INTERCEPT`, `STEYX`, and
`FORECAST`. It is the semantic boundary for the resolver-free scalar
evaluator, the resolver-backed value evaluator, and their independent
validation. It does not authorize formula-cell recalculation, cached-result
use, source refresh, or publication of a changed cell.

The normative source is the repository-local ODF distribution:

| Source | SHA-256 |
| --- | --- |
| archive `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| member `part4-formula/OpenDocument-v1.4-os-part4-formula.html` | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |

The directly relevant sections are §§3.2.3, 3.3, 3.6–3.7, 4.6–4.11.13,
5.6, 6.1–6.3, 6.13.5, 6.13.30, 6.18.17–6.18.18, 6.18.28, 6.18.38,
6.18.41, 6.18.56, 6.18.66, 6.18.69, and 6.18.76. The evaluator also
retains the bounded conversion, resource, source-fence, cache, and
typed-failure rules selected by ADR 0004, ADR 0005, ADR 0006, ADR 0008, and
ADR 0024 and recorded by the aggregate, statistical-reducer, dispersion,
descriptive-statistics, and matrix-function contracts.

## Normative function scope and signatures

These are the only functions in this contract. The local ODF 1.4 Part 4
archive contains no normative entries for `COVARIANCE.P`, `COVARIANCE.S`, or
`FORECAST.LINEAR`; those names are not aliases in this profile. A host may
provide a separate compatibility extension, but an extension must not be
presented as an ODF result or silently change the behavior of these eight
names.

The signatures retain the specification's pseudotypes. `ForceArray` is an
argument attribute, not a value type. Each function returns one `Number` in
scalar evaluation. A semicolon separates formula arguments.

| Function | Part 4 signature | Operation | Cardinality and geometry |
| --- | --- | --- | --- |
| `CORREL` | `CORREL(ForceArray Array N1; ForceArray Array N2)` | Pearson correlation coefficient | Equal rows and columns; at least one admitted numeric pair. |
| `COVAR` | `COVAR(ForceArray Array N1; ForceArray Array N2)` | Population covariance | Equal rows and columns; at least one admitted numeric pair. |
| `PEARSON` | `PEARSON(ForceArray Array IndependentValues; ForceArray Array DependentValues)` | Pearson correlation coefficient | Equal rows and columns; at least one admitted numeric pair. It is identical to `CORREL` in Part 4. |
| `RSQ` | `RSQ(ForceArray Array ArrayY; ForceArray Array ArrayX)` | Square of the Pearson product-moment correlation coefficient | Equal rows and columns. Part 4's empty or different-data-point case is `#N/A`; an equal-cell-count orientation mismatch is `#VALUE!`; a zero variance is `#DIV/0!`. |
| `SLOPE` | `SLOPE(ForceArray Array Y; ForceArray Array X)` | Slope of the linear regression line | Equal rows and columns; at least one admitted numeric pair and nonzero x variance. |
| `INTERCEPT` | `INTERCEPT(ForceArray Array Data_Y; ForceArray Array Data_X)` | y-intercept of the linear regression line | Equal rows and columns; at least one admitted numeric pair and nonzero x variance. |
| `STEYX` | `STEYX(ForceArray Array MeasuredY; ForceArray Array X)` | Standard error of the predicted y value | Equal rows and columns; at least three admitted numeric pairs and nonzero x variance. |
| `FORECAST` | `FORECAST(Number Value; ForceArray Array Data_Y; ForceArray Array Data_X)` | Regression prediction at query x `Value` | Data arrays have equal rows and columns; the data set must have at least one pair and nonzero x variance. |

The strict fixed arity is enforced. A zero-argument or missing required
argument is generated `#VALUE!`; it is not an empty data set. `FORECAST` has
no optional precision, significance, order, or interpolation parameter.

Part 4 explicitly says that `RSQ` returns `#N/A` when its arrays are empty or
have a different number of data points. The repository applies that explicit
result before scanning when two admitted rectangular descriptors have
different checked cell counts. If the checked cell counts are equal but the
row or column orientation differs, the separate `COLUMNS`/`ROWS` constraint
is a shape refusal and is `#VALUE!` before a resolver scan. A valid equal-
shape pair with no admitted numeric pair is the explicit generated `#N/A`
case. A rejected `ReferenceList` or 3-D pseudotype remains `#VALUE!`. The
other no-pair and undersized cases use the profile errors below because Part 4
describes them as constraints or mathematical errors without fixing a
subtype.

## Pair construction, orientation, and array admission

`ForceArray` evaluates each data argument in complete non-scalar array mode.
It prevents implicit intersection and preserves the argument's rows and
columns. A scalar value is a one-by-one Array for a ForceArray parameter. An
inline rectangular Array keeps its declared shape. One rectangular 2-D
Reference keeps its full shape and is traversed by row and then column. A
3-D cuboid or a multi-area `ReferenceList` is not one rectangular Array in
this bounded profile and is a pseudotype/shape refusal. It returns `#VALUE!`
before any resolver cell is read when its descriptor is available. A computed
expression may have to produce its descriptor first; the reducer itself never
scans a rejected list.

The two data arrays must have exactly the same row count and column count.
There is no scalar-to-matrix broadcast, truncation, reshape, transpose, or
implicit intersection. A `1 × N` row and an `N × 1` column are different
shapes even when they contain the same number of cells. A pair is formed at
the same `(row,column)` position in the two arrays. The repository's
row-major order is used for deterministic formula-error precedence and read
receipts; the numerical result does not depend on that order when no error is
present.

For every aligned position, the profile is:

| Left or right member at the position | Pair action |
| --- | --- |
| Finite `Number` on both sides | Admit one `(x,y)` pair. `0` is an ordinary admitted number. |
| `Empty`, `Text`, or `Boolean` on either side | Ignore this position, including the member on the other side. Numeric-looking Text is not parsed for paired data arrays. |
| Formula `Error` on either side | Retain the formula error; it is not an ignorable non-number. If its aligned partner is Empty/Text/Boolean, the error still propagates because the specification's ignore rule names only those three member types. |
| Missing, unsupported, complex, or non-finite numeric member | Generate the selected conversion/domain error, unless a retained formula error has precedence. |

The Empty/Text/Boolean rule is the rule stated by each paired Part 4
function. It differs from the NumberSequence conversion used by aggregate and
descriptive reducers: a Text or distinguished Logical cell in a paired
Reference is omitted rather than parsed or converted to `0`/`1`. A scalar
Text or Logical supplied directly to a ForceArray data parameter is therefore
a one-cell nonnumeric member and yields no pair, rather than a parsed pair.
The rule is a positional omission rule, not a blanket “ignore every
non-number” rule: Formula Error, Missing, unsupported, complex, and non-finite
members remain errors. The direct scalar `Value` query of `FORECAST` uses the
separate Number conversion below.

References are not recalculated. A formula cell's already-provided Number or
formula Error is observed as such; a cached formula result is never imported
to replace a missing or unsupported provider value. In the resolver-free
evaluator, a reference or Array that cannot be retained by its API is the
existing typed `UnsupportedKind::Reference` or `UnsupportedKind::Array`
failure. The resolver-backed evaluator handles only the admitted rectangular
descriptors.

## Equations and selected error profile

For the admitted pairs `(x_i,y_i)`, `i=1..n`, define

```text
x̄   = (1/n) Σ x_i                 ȳ   = (1/n) Σ y_i
Sxx = Σ (x_i - x̄)^2               Syy = Σ (y_i - ȳ)^2
Sxy = Σ (x_i - x̄)(y_i - ȳ)
```

The function results are:

```text
COVAR(N1,N2)    = Sxy / n
CORREL(N1,N2)   = Sxy / sqrt(Sxx Syy)
PEARSON(N1,N2)  = CORREL(N1,N2)
RSQ(ArrayY,X)   = CORREL(ArrayX,ArrayY)^2
SLOPE(Y,X)      = Sxy / Sxx
INTERCEPT(Y,X)  = ȳ - SLOPE(Y,X) x̄
FORECAST(v,Y,X) = ȳ + SLOPE(Y,X) (v - x̄)
STEYX(Y,X)      = sqrt((Syy - Sxy^2/Sxx)/(n - 2))
```

`COVAR` is population covariance and divides by `n`; there is no sample
covariance alias in this contract. `CORREL` and `PEARSON` are identical, and
`RSQ` is the square of that correlation. The stable centered form of `STEYX`
is algebraically the same as the Part 4 raw-sum equation

```text
sqrt((n Σ y_i² - (Σ y_i)²
      - (n Σ x_i y_i - Σ x_i Σ y_i)²/(n Σ x_i² - (Σ x_i)²))
     /(n(n - 2)))
```

The selected subtypes are:

| Condition | Result |
| --- | --- |
| Wrong arity, missing required slot, or rejected data/query list, shape, or pseudotype | `#VALUE!` (`ScalarError::Value`). This structural refusal occurs before resolver cell reads when the descriptor is already known. |
| Malformed or missing scalar `FORECAST` query after its ordinary Number conversion | `#VALUE!` (or `#NUM!` for a non-finite query). It is retained as a formula error while an otherwise admitted data pair is scanned, so a later typed failure can supersede it. |
| No admitted pair for `CORREL`, `COVAR`, `PEARSON`, `SLOPE`, `INTERCEPT`, or `FORECAST` | `#VALUE!`. |
| Two admitted rectangular RSQ arrays have different checked cell counts, or have no admitted pair after positional omission | `#N/A` (`ScalarError::NotAvailable`), as explicitly stated by Part 4; the cell-count case is decided before resolver reads. |
| Two RSQ arrays have equal checked cell counts but different rows or columns | `#VALUE!` shape refusal before resolver reads. |
| Fewer than three pairs for `STEYX` | `#VALUE!`. |
| `Sxx = 0` for `CORREL`, `PEARSON`, `RSQ`, `SLOPE`, `INTERCEPT`, `STEYX`, or `FORECAST`, or `Syy = 0` for correlation/RSQ | `#DIV/0!` (`ScalarError::DivisionByZero`). |
| Non-finite admitted Number, non-finite query, or non-finite final result | `#NUM!` (`ScalarError::Number`). |
| A genuine negative value under the `STEYX` square root, after exact/stable arithmetic has ruled out roundoff, | `#NUM!`. |

The constraints in Part 4 use the generic term `Error` except for RSQ's
explicit `#N/A` condition. This table is the repository's deterministic
subtype profile. A source formula Error already present in an admitted member
is returned as that formula error and has precedence over a later generated
empty, count, variance, or arithmetic result, subject to typed-failure
precedence below.

`STEYX` requires three pairs even though its residual expression can be
written for fewer values. A perfect-fit residual is canonical `+0`; a small
negative residual caused by finite roundoff is clipped to `+0` only when the
stable/exact state establishes that the mathematical residual is zero. An
unclassified negative radicand is `#NUM!`, not a silently complex value.

### INTERCEPT's `LINEST` ambiguity

Section 6.18.38 says that `INTERCEPT` follows
`LINEST(Data_Y,Data_X,FALSE())`. Section 6.18.41 independently and
unambiguously defines `Const=FALSE` to set the model constant `a` to zero.
Taking that token literally would make `INTERCEPT` identically zero and would
contradict the same section's summary (“returns the y-intercept”) and the
ordinary regression used by `SLOPE`, `FORECAST`, and the paired formulas.
This contract records the conflict and selects the coherent y-intercept
profile `a = ȳ - b x̄`, equivalent to the regression with an included
constant (`Const=TRUE`). The literal `FALSE()` token must not silently turn
the implementation into a through-origin regression. A future conformance
decision that treats the apparent token as authoritative would be a deliberate
contract change, not an incidental implementation detail.

## Direct `FORECAST` query and matrix lifting

`Data_Y` and `Data_X` are complete ForceArray arguments. The fit is computed
from all admitted aligned pairs and has one shape descriptor. `Value` is the
only ordinary scalar parameter. The selected finite Number bridge for a
direct query is:

* finite Number contributes itself;
* Logical contributes `0` or `1`;
* scalar Empty contributes `0`;
* finite locale-independent decimal Text is parsed;
* malformed, missing, complex, or non-finite input is generated `#VALUE!` or
  `#NUM!` according to the bridge above.

In scalar mode, a multi-cell `Value` uses the ordinary §3.3 scalar projection
rule: an inline Array supplies its `[0,0]` element and a Reference uses scalar
implied intersection. A `ReferenceList` is not a scalar Number and is a
structural pseudotype refusal: it is rejected with `#VALUE!` without reading
its cells, and the reducer does not begin a data scan once that known refusal
has been established. By contrast, an ordinary scalar query that is already a
Formula Error or converts to a formula `#VALUE!`/`#NUM!` (for example
malformed Text, Missing, or a non-finite Number) is retained while an accepted
data pair is scanned. That scan is required so a later typed provider,
resource, cancellation, or source failure can supersede the query's formula
error. If no typed failure occurs, the query error is published for the scalar
result after the fit attempt.
In matrix mode, because `FORECAST` returns a scalar Number, a rectangular
non-scalar `Value` argument is iterated and produces a result of the same
shape. Each query position uses its own Number conversion and can produce its
own generated conversion error; the complete data fit is reused when the data
arguments are coordinate-independent, and one bad query position does not
prevent the other positions from being evaluated.

The §3.3.2.2.1 exception applies to a function that returns an Array when it
is evaluating that function's own scalar input. Thus a direct `MUNIT({2;3})`
uses its `[0,0]` size input and produces one array. If that already-produced
Array is subsequently supplied as `FORECAST`'s ordinary scalar `Value`, the
scalar-returning `FORECAST` may lift over the produced elements in matrix
mode. The MUNIT result is not collapsed merely because it was produced by an
array-returning function. MUNIT's own scalar criterion remains
position-sensitive and is excluded from full-argument demand-cache
propagation.

## Formula errors, typed failures, and source fences

Arguments are evaluated in source order. Within each admitted data array, the
conceptual cell order is row-major; paired positions are inspected together.
The first formula Error in that order is retained. A nonnumeric Empty/Text/
Boolean pair is ignored, but an Error is not ignored even if its partner is a
non-numeric member.

The reducer continues an accepted reference scan after retaining a formula
Error. This lets a later typed resolver, provider, resource, cancellation, or
source-version failure supersede the retained formula result. A typed
`Unsupported`, `ResourceLimit`, `Allocation`, `Cancelled`, `SourceChanged`,
or `SourceVersionAvailabilityChanged` failure remains an `EvaluationFailure`;
it is never converted into a formula Error and is never caught by
`IFERROR`/`IFNA`. If no typed failure occurs, the retained source formula
Error wins over generated empty, cardinality, denominator, or final-result
errors.

For `FORECAST`, the query is the first formula argument. A scalar query
Formula Error or conversion error therefore precedes a Formula Error found in
the later data arguments, although the accepted data scan still completes for
typed-failure precedence. In matrix query mode this ordering is per output
position: a query error belongs to its own position, while a later data
Formula Error applies to every position that has no earlier query error.

Shape, list-kind, and arity gates are performed before resolver reads whenever
their descriptors are already available. They do not permit a rejected
ReferenceList or mismatched matrix to be partially scanned. An upstream
computed reference expression may itself need evaluation to reveal its
descriptor; that evaluation is separate from the paired reducer's zero-read
guarantee.

The value evaluator checks the source version and cancellation before the
calculation and again after all fit/query work, before publication. A source
fence or cancellation failure therefore supersedes any formula result,
including a retained source formula Error.

## Numeric precision and signed zero

Part 4 specifies the equations rather than a particular finite-precision
algorithm. The profile uses a fixed-size centered paired state (or an
equivalent exact/compensated binary state) containing counts, means, and
`Sxx`, `Syy`, and `Sxy`. It must avoid deciding a finite result from an
intermediate overflow or catastrophic uncentered cancellation. No arbitrary
semantic epsilon may turn a mathematically valid nonzero variance, covariance,
or residual into a domain error. The state may use the repository's fixed
dyadic moment helper, but it must not retain a per-cell input vector.

`Sxx` and `Syy` are nonnegative. Exact zero results for correlation,
`RSQ`, `STEYX`, and other mathematically zero nonnegative quantities publish
canonical `+0`. A signed result (`COVAR`, `CORREL`, `PEARSON`, `SLOPE`,
`INTERCEPT`, or `FORECAST`) that is mathematically nonzero but underflows in
the final binary64 conversion may retain `-0` when its sign is negative. An
exact algebraic zero is always `+0`; a negative underflow is not rewritten as
positive merely because it compares equal to zero. `RSQ` is nonnegative and
canonicalizes an exact zero to `+0`.

The implementation may use compensated means, scaled centered products, exact
finite-binary64 sums, or the fixed-width dyadic helper. The helper's checked
width is an implementation capacity/resource boundary, not a semantic
precision cap or an arbitrary tolerance. NaN and infinity never escape the
evaluator. Ordinary last-bit differences from the chosen finite binary64
algorithm use the repository's documented finite-number tolerance; the
equations and error/domain boundaries remain exact.

## Resource and memory boundaries

The resolver-backed evaluator streams admitted paired References in lockstep.
It does not materialize a range into a cell vector or sort/re-scan one range
per output query. It must:

* validate rectangular geometry, list/ref kind, and checked row/column/cell
  products before a reference scan;
* reserve only AST arrays, bounded reference metadata, and fixed-size paired
  reducer state. Inline arrays are already bounded AST values; no range-sized
  `Vec` or per-cell error/value buffer is permitted;
* charge AST work, argument conversion, every physically inspected cell,
  retained Text bytes, and every reducer operation to the caller's
  hierarchical budget;
* call `charge_cell_work` before each `read_reference_cell`, enforce one
  cumulative `max_reference_cells` ceiling across both arrays and all
  accepted occurrences, and check cancellation before and after each read;
* use `read_to_element` with borrowed provider Text. Text bytes are charged
  even when the aligned pair is then ignored; the reducer must not clone a
  borrowed string for each observation;
* retain formula Errors as fixed-size state and continue the admitted primary
  scan so later typed failures can supersede them; and
* release every temporary reservation on formula and typed failure paths.

All paired data inputs share the cumulative cell/read/work limits; the limit
is not reset or applied independently per array. If a numerical implementation
uses a replay to complete a centered or exact state, every replayed cell,
operation, and cancellation check is charged again and the source fence
surrounds the complete sequence of passes. A normal paired reducer needs one
lockstep pass. `FORECAST` computes one fit and reuses it for an invariant
matrix of query values; it does not re-read or recompute the data pair for
each query. A position-dependent data expression can require reevaluation at
the positions selected by the evaluator's cache classifier, and that required
reevaluation is not prohibited by this contract.

## Matrix shape, demand cache, and publication

The seven data-only reducers (`CORREL`, `COVAR`, `PEARSON`, `RSQ`, `SLOPE`,
`INTERCEPT`, and `STEYX`) consume complete paired ForceArray arguments. Their
full reference/array descriptors may be propagated through a projected lazy
branch when the classifier proves them coordinate-independent. They are
cacheable as scalar Number or formula-error payloads only after both complete
data arguments, their shapes, and their source/context identity are part of
the cache key. A typed failure is never cached as a formula Error.

`FORECAST` has the same complete data arguments plus a scalar query. The fit
may be cached/reused across query positions only when the complete data pair
is invariant. `Value` remains a position-sensitive scalar criterion and is not
promoted to a full-argument cache merely because `FORECAST` is a statistical
reducer. A nested `MUNIT` scalar criterion remains excluded from full-argument
propagation, while its produced Array follows the ordinary matrix lifting rule
once consumed by `FORECAST`.

No scalar paired reducer publishes a matrix merely because it ran in matrix
mode. `FORECAST`'s matrix result is the enclosing scalar-parameter iteration;
the other seven functions publish one scalar result per invocation. Any
formula-error payload is published only after the source/cancellation fence,
resource accounting, and typed-failure checks have completed.

## Native and validation boundary

Native spreadsheet-host behavior is compatibility evidence only and cannot
add `COVARIANCE.P`, `COVARIANCE.S`, or `FORECAST.LINEAR` to the normative
surface. It also cannot replace the selected INTERCEPT ambiguity resolution,
pairwise Text/Logical/Empty omission, error subtype profile, signed-zero
policy, or resource rules.

Validation must cover every normative function through direct scalars, inline
Arrays, empty and mixed References, shape/orientation mismatches, formula
Errors, and rejected ReferenceLists. It must include:

* one-pair and no-pair cases; covariance's population denominator; correlation
  and RSQ zero-variance cases; SLOPE/INTERCEPT/FORECAST x-variance failures;
  STEYX's two/three-pair boundary; and perfect fits;
* paired Empty, Text, numeric-looking Text, Boolean, zero, formula Error,
  Missing, complex, and non-finite members, including a nonnumeric member on
  only one side of an aligned position;
* `1 × N` versus `N × 1` orientation, scalar-to-matrix refusal, multi-area
  and 3-D reference refusal, and zero resolver reads for known invalid shape or
  list descriptors;
* exact cancellation, extreme finite magnitudes, avoidable intermediate
  overflow, centered residual accuracy, exact `+0`, negative underflowed
  `-0`, final `#NUM!`, and the documented finite-number tolerance;
* FORECAST query scalar conversion, scalar projection, matrix lifting,
  invariant fit reuse, mixed per-query errors, and MUNIT's producer-input
  `[0,0]` rule versus its produced-array consumption;
* first formula-error order, completion of an admitted scan after an error,
  typed resolver/resource/cancellation precedence, source-version fences,
  borrowed Text, cumulative read/work limits, and no per-output data re-read;
  and
* projected lazy branches that preserve complete paired data descriptors,
  cache scalar results/errors only with full source/context identity, reuse an
  invariant FORECAST fit, and leave a nested MUNIT scalar criterion
  position-sensitive.

The independent oracle must evaluate the equations with wider or exact
intermediates where needed. Native caches can corroborate ordinary finite
results but cannot redefine the eight-function ODF scope, the recorded
INTERCEPT ambiguity decision, or the selected resource and error behavior.
