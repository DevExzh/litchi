# ODF 1.4 order and rank statistics evaluator contract

This contract defines the bounded implementation profile for the eight order
and rank functions in OpenFormula 1.4 Part 4: `MEDIAN`, `MODE`, `LARGE`,
`SMALL`, `PERCENTILE`, `PERCENTRANK`, `QUARTILE`, and `RANK`. It is the
semantic boundary for the resolver-free scalar evaluator, the resolver-backed
value evaluator, and their independent validation. It does not authorize
formula-cell recalculation, cached-result use, source refresh, or publication
of a changed cell.

The normative source is the repository-local ODF distribution:

| Source | SHA-256 |
| --- | --- |
| archive `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| member `part4-formula/OpenDocument-v1.4-os-part4-formula.html` | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |

The directly relevant sections are §§3.2.3, 3.3, 3.6–3.7, 4.10–4.11.12,
5.6, 6.1–6.3, 6.17.2, 6.17.6–6.17.7, and 6.18.40, 6.18.47, 6.18.50,
6.18.57–6.18.58, 6.18.64–6.18.65, and 6.18.70. The implementation also
uses the bounded conversion, resource, source-fence, and typed-failure rules
selected by ADR 0004, ADR 0005, ADR 0006, ADR 0008, and ADR 0024 and already
recorded by the aggregate, statistical-reducer, and dispersion contracts.

## Signatures and result shapes

The signatures below preserve the specification's pseudotypes. `+` means one
or more supplied arguments in the common function template (§6.2); square
brackets denote an optional argument, and a semicolon separates formula
arguments. `#VALUE!` below means `ScalarError::Value`, `#NUM!` means
`ScalarError::Number`, and `#N/A` means `ScalarError::NotAvailable`.

| Function | Part 4 signature | Result | Constraint or shape rule |
| --- | --- | --- | --- |
| `MEDIAN` | `MEDIAN({ NumberSequenceList X }+)` | Number | At least one supplied sequence argument; the middle ordered value is returned, or the average of the two middle values. |
| `MODE` | `MODE({ ForceArray NumberSequence N }+)` | Number | At least one supplied forced-array sequence; a value must occur at least twice. |
| `LARGE` | `LARGE(NumberSequenceList List; Number\|Array N)` | Number or Array | `ROUNDUP(N;0)=N` and every selected rank is between 1 and the number of values. An Array `N` returns a same-shaped result Array. |
| `SMALL` | `SMALL(NumberSequenceList List; Integer\|Array N)` | Number or Array | `ROUNDDOWN(N;0)=N` and every selected rank is between 1 and the number of values. An Array `N` returns a same-shaped result Array. |
| `PERCENTILE` | `PERCENTILE(NumberSequenceList Data; Number X)` | Number | `COUNT(Data)>0` and `0≤X≤1`; interpolation uses the sample rank from the specification. |
| `PERCENTRANK` | `PERCENTRANK(NumberSequenceList Data; Number X [; Integer Significance=3])` | Number | `COUNT(Data)>0`, `MIN(Data)≤X≤MAX(Data)`, and `Significance` is a positive integer. |
| `QUARTILE` | `QUARTILE(NumberSequence Data; Integer Quart)` | Number | `COUNT(Data)>0` and `0≤Quart≤4`; it is `PERCENTILE(Data;Quart/4)`. |
| `RANK` | `RANK(Number Value; NumberSequenceList Data [; Number Order=0])` | Number | `Value` must occur in `Data`; `Order=0` means descending and any nonzero value means ascending. |

The repository profile reports `#VALUE!` for invalid arity, an omitted
required sequence, a supplied missing slot, a pseudotype or shape mismatch,
an unsuccessful scalar conversion, an empty sequence, a violated rank/domain
constraint, a mode with no repeated value, and a `RANK` value absent from its
data. Part 4 specifies these cases as an unspecified `Error`; this subtype
choice follows the existing sequence-reducer profile. The local LibreOffice
fixtures use `#VALUE!` for no-mode and missing-value cases and their generic
`Err:502` invalid-argument result for rank/domain cases. This repository maps
both invalid-argument classes to `#VALUE!` because its public scalar error
model has no separate invalid-argument subtype. A
non-finite input or final numeric result is `#NUM!`. A formula `#N/A` already
present in an admitted argument remains that formula error; a missing `RANK`
value is the generated `#VALUE!` profile error rather than a generated
`#N/A`.

All eight calls with zero supplied arguments return `#VALUE!`. A call with an
explicit missing slot, such as `MEDIAN(;)` or `PERCENTRANK(A1:A3;2;)`, has a
supplied argument and observes a generated `#VALUE!`; it is not an omitted
argument or an optional default. `Order` and `Significance` defaults apply only
when the optional argument is absent from the syntax. A syntactically empty
array has no admitted Number values and therefore follows the empty-sequence
rule.

The scalar evaluator publishes one Number for a scalar result. In matrix mode
the ordinary §3.3 scalar-parameter iteration can publish an Array of results;
the value evaluator preserves the checked rectangular shape and existing
`#N/A` out-of-shape behavior. `LARGE` and `SMALL` have an additional explicit
Array result: when `N` is an Array, each rank produces one output at the same
row-major position and the output shape is exactly the shape of `N`.

## Sequence pseudotypes and reference admission

The pseudotype distinction is observable and is retained in both evaluators.

* `MEDIAN`, `LARGE`, `SMALL`, `PERCENTILE`, `PERCENTRANK`, and `RANK` consume
  `NumberSequenceList`. A scalar Number, Text, Logical, or Empty value forms a
  one-member sequence under the scalar Number conversion profile. A logical
  rectangular Reference contributes Number and formula Error cells in
  reference order, omitting referenced Empty, Text, and distinguished Logical
  cells. An ordered `ReferenceList` is admitted and flattened in list
  occurrence order; each area is traversed by increasing sheet, row, and
  column. Overlaps are retained once per occurrence. A sequence reducer never
  applies implicit intersection to an admitted reference.
* `MODE` consumes `ForceArray NumberSequence`. `ForceArray` evaluates each
  argument in non-scalar array mode, so an inline Array or a logical
  rectangular Reference keeps its complete shape. The subsequent
  `NumberSequence` conversion still uses the sequence filtering above for
  referenced cells. A `ReferenceList` is not a single `NumberSequence` or a
  rectangular Array in this profile; it is a pseudotype/shape mismatch and
  returns `#VALUE!` before any resolver cell is read.
* `QUARTILE` consumes `NumberSequence`, not `NumberSequenceList`. A logical
  rectangular Reference, scalar value, or inline Array is admitted using the
  same sequence conversion, but an explicit `ReferenceList` is a
  pseudotype/shape mismatch and returns `#VALUE!` before any resolver cell is
  read. A three-dimensional cuboid remains one logical Reference and is
  traversed in sheet order.
* An inline rectangular Array is an already-valued sequence. Its elements are
  visited row-major and use the scalar conversion profile: finite Number is
  included, Logical is `0` or `1`, Empty is `0`, finite numeric Text is parsed,
  malformed or non-finite Text is generated `#VALUE!`, Missing is generated
  `#VALUE!`, and a formula Error is retained. Complex values are generated
  `#VALUE!` under this finite real profile. This Array rule is separate from
  the omission of Text, Empty, and distinguished Logical cells in a cell
  Reference.
* A direct scalar Text uses the repository's locale-independent finite decimal
  bridge. A malformed, NaN, or infinite Text produces generated `#VALUE!`.
  Scalar Logical values convert to `0`/`1`; scalar Empty converts to `0`;
  explicit Missing remains `#VALUE!`. Formula Errors are retained and are
  never converted.

The scalar evaluator rejects references and arrays when its resolver-free API
cannot retain their descriptors, using its existing typed
`UnsupportedKind::Reference` or `UnsupportedKind::Array` refusal. The value
evaluator is the resolver-backed path for valid references and arrays. Neither
evaluator recalculates a formula cell or uses a cached formula result.

## Scalar parameters and matrix evaluation

Section 3.3 says that a non-scalar value passed to a function parameter that
expects a scalar is evaluated by iteration in matrix mode. This rule applies
to the `Number` and `Integer` parameters in this batch; they are not
silently rejected as database `Field` or conditional `Criterion` values.

* `PERCENTILE` lifts `X`; `PERCENTRANK` lifts `X` and, when supplied, the
  `Significance`; `QUARTILE` lifts `Quart`; and `RANK` lifts `Value` and, when
  supplied, `Order`. `LARGE` and `SMALL` lift a scalar `N` in the same way.
  The output shape is the maximum rows and columns of these scalar arguments,
  with the §3.3 singleton, row-vector, column-vector, and in-range matrix
  projection rules. A scalar sequence argument is repeated for every output
  position.
* In scalar mode, a scalar parameter uses the current-position implicit
  intersection rules. A multi-cell Reference in a scalar parameter is not
  silently flattened. A `ReferenceList` is not a Number or Integer and
  returns `#VALUE!` without resolver reads. In a projected lazy branch, a
  computed scalar parameter is evaluated at the requested output coordinate,
  so its position dependence remains visible.
* `LARGE` and `SMALL` additionally admit a value that is already an Array for
  their `Number|Array` or `Integer|Array` parameter. That explicit Array path
  is consumed in full and returns one rank result per input element, with the
  same shape. It is distinct from generic matrix iteration of a scalar `N`.
  A bare Reference in `N` is evaluated as the scalar Number/Integer path
  unless the expression explicitly enters array context; the absence of a
  `ForceArray` marker does not authorize implicit full-range materialization.
* An Array result from the explicit `LARGE`/`SMALL` `N` path survives scalar
  publication, parenthesization, and a scalar-condition `IF` as an Array
  value, including the one-cell `1×1` case. It is not implicitly projected or
  iterated again merely because the enclosing evaluator is in scalar mode. A
  consumer whose parameter specifically demands a Number may project that
  Array under the ordinary scalar conversion rules; a matrix-capable consumer
  may retain its shape.
* `MODE`'s `ForceArray` marker applies to its sequence arguments only. The
  §3.3.2.2.1 exception applies when an Array-returning function is itself
  being evaluated in matrix mode: a non-scalar input to that function is not
  implicitly iterated, and its `[0,0]` input element is used. Thus a direct
  `MUNIT` call with an Array-valued Unit argument uses the first Unit element,
  including when the call occurs in a projected `IF`; this is the scope shown
  by the specification's `SUM(INDIRECT({"A1";"A2"}))` example. Once MUNIT has
  produced an Array, an ordinary scalar-parameter consumer applies §3.3.2 to
  that result: `PERCENTILE(Data;MUNIT(2))` may lift over MUNIT's returned
  positions and publish a correspondingly shaped result Array. A separate
  position-sensitive MUNIT case occurs in a scalar conditional `Criterion`
  context, where the scalar projection is evaluated at the current criterion
  position. Neither case is promoted into full-argument demand caching.

Sequence arguments are complete descriptors in every mode. Data and List
references therefore remain whole ranges under a projected `IF`; they are not
implicitly intersected at the output cell. Scalar parameter values may be
coordinate-dependent even when the sequence is invariant.

## Ordered operations and tie rules

The implementation compares finite binary64 values numerically, with `-0` and
`+0` equal for ordering and rank comparisons. A value selected directly from
the ordered data (an odd MEDIAN, MODE, LARGE/SMALL result, or percentile
endpoint) is published as canonical `+0`; this is the same zero policy as the
existing extrema reducers. A computed even-MEDIAN midpoint retains the sign of
its final rounded result, including a negative underflow `-0`. A percentile or
quartile interpolation with a nonzero fractional weight follows the same
computed-result rule; an endpoint with zero weight is a direct selection and
is canonical `+0`. Duplicate occurrences are retained for all rank counts.

### MEDIAN and MODE

`MEDIAN` sorts the concatenated sequence. For an odd count `n`, it returns the
value at one-based position `(n+1)/2`; for an even count it returns the
arithmetic mean of positions `n/2` and `n/2+1`. The selected mean uses the
repository's fixed numeric average profile: the represented operands are
combined without an overflowing intermediate and one final finite binary64
value is published.

`MODE` selects the value with the largest occurrence count. If several values
have that count, the smallest value is returned. A sequence with no value
occurring at least twice is undefined under §6.18.50 and returns the selected
`#VALUE!` profile. Equality is binary64 numerical equality, so both signed
zeros belong to one mode. Formula Errors are not mode candidates; they remain
formula errors under the precedence rules below.

### LARGE and SMALL

`LARGE` sorts the admitted values descending for rank selection, and `SMALL`
sorts ascending. Rank `1` is the first value in that order. Duplicate values
occupy separate rank positions. Each `N` must be finite, positive, and an
exact integer according to its stated constraint: `LARGE` uses the
`ROUNDUP(N;0)=N` condition and `SMALL` uses `ROUNDDOWN(N;0)=N`; the profile
does not silently truncate a fractional rank. A rank outside `1..count` is
generated `#VALUE!`.

For Array `N`, each element goes through the same constraint and produces its
own Number or formula Error at the corresponding output position. The data is
sorted once and reused for all rank elements. A formula Error in the data
sequence applies to every output position after argument/error precedence is
resolved.

### PERCENTILE and QUARTILE

For sorted ascending values `y₁ … yₙ`, `PERCENTILE(Data;X)` uses the exact
sample-rank algorithm in §6.18.57:

```text
r = 1 + X * (n - 1) = I + D
I = floor(r)
D = r - floor(r)
result = y_I + D * (y_(I+1) - y_I)
```

The one-based endpoints are used when `D=0`; `X=0` returns the minimum and
`X=1` returns the maximum. Interpolation is an overflow-safe convex
combination of the two finite operands, and the final result is rounded to
finite binary64. The implementation must not form an overflowing
`y_(I+1)-y_I` merely to interpolate two opposite-sign finite values. The
profile's stable binary64 result is compared using the existing finite-number
tolerance where a platform's intermediate rounding is observable.

`QUARTILE(Data;Quart)` delegates to the same algorithm with `X=Quart/4`.
`Quart` must already be an integer in `0..4`; it is not silently rounded or
truncated. Quartiles `0`, `1`, `2`, `3`, and `4` are respectively the minimum,
25th percentile, median, 75th percentile, and maximum.

### PERCENTRANK

After sorting the data ascending, an exact `X` receives the rank equal to the
number of data values strictly less than `X`. This gives duplicate values the
same lowest (first-occurrence) rank. For `n>1`, the unrounded result is
`rank/(n-1)`. If `X` lies strictly between adjacent distinct data values `Y`
and `Z`, let `rY` be the rank of `Y`; the fractional rank is:

```text
rx = rY + (X - Y) / (Z - Y)
result = rx / (n - 1)
```

The interpolation is performed with finite, overflow-safe differences. When
`n=1`, only the one data value is valid and the result is `1`, as specified by
§6.18.58. `X` outside the inclusive data minimum and maximum violates the
constraint and returns generated `#VALUE!`.

`Significance` defaults to `3` only when omitted. It must be a positive exact
integer; the profile rounds the result to that many decimal places using the
existing `ROUND` (nearest, ties away from zero) kernel. Thus the native-style
examples `PERCENTRANK({1;2;3;4};3)` and
`PERCENTRANK({1;2;3;4};3;4)` publish `0.667` and `0.6667`, respectively.
The rounded result remains in `[0,1]`. A fractional, zero, negative, or
non-finite significance is generated `#VALUE!`.

### RANK

`RANK(Value;Data;Order)` uses the exact binary64 value after Number
conversion. When `Order` is omitted or zero, the rank is

```text
1 + count(data value > Value)
```

When `Order` is nonzero, the rank is:

```text
1 + count(data value < Value)
```

Equal values receive the same rank, and the next distinct value's rank skips
all tied occurrences (competition ranking). `Order` is a Number, not an
Integer: any finite nonzero value selects ascending order. A `Value` absent
from `Data` returns generated `#VALUE!` under this profile.

## Formula-error precedence and typed failures

Arguments are evaluated eagerly in source order. Within a sequence, the
conceptual order is argument order, ReferenceList occurrence order, sheet
order for a three-dimensional Reference, and row-major cell order. The
evaluator retains the first formula Error in that order while continuing the
bounded admitted scan. This allows a later typed resolver, source, resource,
cancellation, or allocation failure to supersede the retained formula error.

If no formula Error is encountered, the first generated formula error in
conceptual order is returned. Generated conversion, mode, empty-sequence,
rank/domain, and interpolation failures use the `#VALUE!`/`#NUM!` profile
above. A formula Error in a sequence is not silently omitted merely because a
different member would provide enough numbers. For scalar-parameter matrix
iteration, each output coordinate has its own scalar-argument formula-error
result; a data-sequence formula Error is common to every coordinate.

Typed `Unsupported`, cancellation, source-version changes, resource limits,
and allocation failures remain `EvaluationFailure` values. They are never
converted into formula Errors and are never caught by `IFERROR` or `IFNA`.
The evaluator must continue an admitted scan after retaining a formula Error
long enough to perform its required work, read, cancellation, and source-fence
checks.

## Precision, memory, and cache boundaries

Some order statistics require retained numeric values for sorting, while
scalar rank queries can remain streaming. Neither path may retain a parallel
vector of `RuntimeElement`, cell text, or resolver records. The value evaluator
streams every admitted reference cell,
charges cell work before each read, charges borrowed Text bytes before
conversion, and retains only finite numeric operands plus bounded formula-error
state. Numeric-vector capacity is checked against `max_reference_cells`, array
limits, the hierarchical storage budget, and fallible reservation before
growth. For variadic functions, the numeric-observation bound is cumulative
across all arguments and ReferenceList occurrences; an inline array's per-array
cell limit must not be reset and reused as a separate allowance for every
argument. Duplicate ReferenceList occurrences remain visible in the numeric
vector; they are not deduplicated.

`MEDIAN`, `MODE`, `LARGE`, `SMALL`, `PERCENTILE`, and `QUARTILE` need ordered
data. They may retain one checked numeric vector and sort it once per
invocation; `MODE` then performs one ordered frequency pass and
`LARGE`/`SMALL` answer all rank elements from that sorted vector. A scalar
`RANK` query does not require a vector or sort: it may stream the data once
while counting strictly greater or smaller values and checking exact presence.
A scalar `PERCENTRANK` query likewise may stream once while retaining its
count, boundary, and adjacent-bracket state. When invariant data serves
multiple rank or percent-rank queries, the implementation may retain one
sorted vector or one bounded query-state pass and reuse it. If the sequence
expression is position-dependent under a projected matrix demand, it must be
reevaluated for each required position; the cache rules below determine when
that work can be reused. Sorting comparisons, stream comparisons, and each
selected/interpolated result are charged evaluator work. No unnecessary
per-output re-sort or re-read of an invariant sequence is permitted, but the
resource contract does not forbid required reevaluation of a position-dependent
expression.

Non-finite input Numbers and non-finite final results are generated `#NUM!`;
NaN and infinity never escape. Stable interpolation must retain finite
results when an intermediate subtraction would overflow. Directly selected
zero operands are canonical `+0`; computed zero results retain the signed
underflow behavior described above. The numerical model is finite binary64
with round-to-nearest-even for ordinary conversion; where an operation's
internal rounding is not invariant, validation uses the repository's
documented finite-number tolerance rather than claiming host-independent
decimal exactness.

Complete `NumberSequence`/`NumberSequenceList` data descriptors are eligible
for demand caching inside a projected lazy branch because their result is
position-independent. Scalar Number/Integer parameters remain cacheable only
when their complete expression is proven coordinate-independent. Computed
reducers under a projected lazy `IF` remain conservative. Nested statistical
criterion propagation may carry complete sequence arguments, but `MUNIT`
remains excluded from full-argument propagation and stays position-sensitive.
Cache entries contain only a scalar Number or formula-error payload plus the
source/context identity. An Array result from `LARGE` or `SMALL` is not stored
in the scalar payload cache, and a typed failure is never cached as a formula
Error.
