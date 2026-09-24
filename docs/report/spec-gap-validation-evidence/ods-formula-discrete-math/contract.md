# ODF 1.4 discrete mathematical function contract

This contract defines the bounded implementation profile for the eleven
functions in OpenFormula 1.4 Part 4 §6.16.  The set is `COMBIN`, `COMBINA`,
`FACT`, `FACTDOUBLE`, `GCD`, `LCM`, `MULTINOMIAL`, `EVEN`, `ODD`, `DELTA`, and
`GESTEP`.  It applies to the resolver-free scalar evaluator and to the
resolver-backed value evaluator.  It does not authorize recalculation of a
worksheet formula cell or use of a cached formula result.

The normative source is the local distribution:

- archive: `3rdparty/specs/OpenDocument-v1.4-os.zip`
- archive SHA-256: `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`
- member: `part4-formula/OpenDocument-v1.4-os-part4-formula.html`
- member SHA-256: `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`
- relevant general sections: §§3.2.3, 3.3, 3.6–3.7, 4.3, 4.11.5,
  4.11.12, 5.6, 6.1–6.3
- function sections: §§6.16.16–6.16.17, 6.16.26, 6.16.30,
  6.16.32–6.16.38, 6.16.43–6.16.44

The archive and member digests are part of this contract.  A later source
change requires a fresh normative review rather than silently reusing these
decisions.

## Signatures and constraints

The table preserves the pseudotypes in Part 4.  A `+` means one or more
supplied parameters in the §6.2 common function template.  The semicolon is
the formula argument separator.  `#VALUE!` below means the formula Error value
`ScalarError::Value`; `#NUM!` means `ScalarError::Number`.

| Function | Part 4 signature | Result and operation | Normative constraint or default |
| --- | --- | --- | --- |
| `COMBIN` | `COMBIN(Integer N; Integer R)` | Number, the binomial coefficient `C(N,R)` | `N >= 0`, `R >= 0`, `R <= N`; each parameter is truncated with `INT` before use |
| `COMBINA` | `COMBINA(Integer N; Integer M)` | Number, combinations with repetition, displayed as `C(N+M-1,N-1)` | `N >= 0`, `M >= 0`, `N >= M`; each actual argument is truncated with `INT` before use; the standard permits an evaluator extension for `N >= 0`, `M >= 0`, but this profile does not enable `N < M` |
| `FACT` | `FACT(Integer F)` | Number, `F!`; `0! = 1! = 1` | `F >= 0`; Integer conversion is `INT` in this profile |
| `FACTDOUBLE` | `FACTDOUBLE(Integer F)` | Number, `F * (F-2) * ... * 1` or the corresponding even product; `0!! = 1!! = 1` | `F >= 0`; Integer conversion is `INT` in this profile |
| `GCD` | `GCD({NumberSequenceList X}+)` | Number, largest integer dividing every selected `INT(a)` | Every selected `INT(a) >= 0`; at least one selected `INT(a) > 0`; if all selected values are zero the standard permits either Error or `0`, and this profile selects `0` |
| `GESTEP` | `GESTEP(Number X [; Number Step = 0])` | Number `1` when `X >= Step`, otherwise `0` | `Step` is `0` only when the optional argument is omitted; supplied values use the Number conversion profile |
| `LCM` | `LCM({NumberSequenceList X}+)` | Number, smallest non-negative integer multiple of every selected value | Every selected `X` satisfies `INT(X) = X` and `X >= 0`; this profile enforces that constraint before the LCM operation |
| `MULTINOMIAL` | `MULTINOMIAL({NumberSequence A}+)` | Number, `FACT(a1+...+an) / (FACT(a1) * ... * FACT(an))` | The source lists no additional constraint.  This profile preserves the expression's raw sum and applies Integer conversion at each nested `FACT` call; see below |
| `EVEN` | `EVEN(Number N)` | Number, an even integer with the sign of `N` and absolute value at least `abs(N)` | No source constraint; “up” is away from zero |
| `ODD` | `ODD(Number N)` | Number, an odd integer with the sign of `N` and absolute value at least `abs(N)` | No source constraint; `ODD(0) = 1`; “up” is away from zero |
| `DELTA` | `DELTA(Number X [; Number Y = 0])` | Number `1` when `X = Y`, otherwise `0` | `Y` is `0` only when omitted |

An invalid fixed or variadic arity is `#VALUE!` after the supplied argument
expressions have been evaluated.  A supplied empty slot, such as
`DELTA(1;)`, is a supplied Missing value and is not an omitted optional
argument; it is `#VALUE!`.  A supplied range that contains no admitted
sequence values is a separate empty-sequence case, described below.

`COMBINA(0;0)` is selected as the empty-multiset identity `1`.  The displayed
binomial formula has `-1` in both positions at this boundary, so it does not
define that case directly.  The choice is explicit profile policy within the
otherwise satisfied `N >= M` constraint.  `COMBINA(0;M)` for `M > 0` violates
`N >= M` and is `#NUM!`; `COMBINA(N;0)` for `N > 0` is `1`.  An evaluator must
not silently enable the standard's optional `N < M` extension.

## Number, Integer, and sequence conversion

The evaluator uses the existing finite binary64 Number bridge.  A finite
Number is used as is.  Logical `FALSE` and `TRUE` convert to `0` and `1`.
Text is parsed by the repository's locale-independent finite decimal parser;
malformed, NaN, and infinite text are `#VALUE!`.  A scalar Empty value in the
value bridge converts to `0`; explicit Missing remains `#VALUE!`.  Formula
Error values remain formula Errors and are never converted.  Complex and
unsupported values are `#VALUE!` or the existing typed capability failure,
respectively.

Part 4's `Integer` pseudotype is a Number equal to `INT(X)` (§4.11.5), and
`INT` means floor toward negative infinity (§6.3.6 and §6.17.2).  This is
different from truncation toward zero.  The selected profile therefore has
the following conversion order:

| Operation | Conversion and check order |
| --- | --- |
| `COMBIN`, `COMBINA` | Number conversion, `INT`, then non-negative and relation checks, then the binomial operation |
| `FACT`, `FACTDOUBLE` | Number conversion, `INT`, then the non-negative check, then the factorial product |
| `GCD` | NumberSequenceList admission, `INT` for every selected value, then non-negative and “one positive” checks |
| `LCM` | NumberSequenceList admission, retain the Number value, require exact `INT(value) = value` and `value >= 0`, then compute LCM; a fraction is `#NUM!` and is not silently truncated |
| `MULTINOMIAL` | NumberSequence admission, retain each raw Number value, sum those raw values, then apply `INT` through the numerator `FACT`; apply `INT` independently through each denominator `FACT` |
| `EVEN`, `ODD`, `DELTA`, `GESTEP` | Number conversion only; no Integer conversion |

The order is observable for negative fractions.  For example,
`FACT(-0.5)` and `FACTDOUBLE(-0.5)` floor to `-1` and fail their domain
constraint; they do not become `FACT(0)`.  `GCD(0.5;0)` floors to two zero
integers and therefore takes the selected all-zero result `0`.  `LCM(0.5;2)`
fails the exact-integer constraint even though the text of its semantics later
says that `INT` is applied.  The explicit constraint is applied first.

### NumberSequence and NumberSequenceList

`NumberSequence` and `NumberSequenceList` are not interchangeable.

* A scalar Number, Text, or Logical contributes one value after Number
  conversion.  A scalar Empty follows the scalar bridge above.
* A rectangular cell Reference consumed as a sequence contributes Number and
  formula Error cells in reference occurrence order.  Referenced Empty and
  Text cells are omitted.  A distinguished Logical cell is omitted, as
  required by §6.3.7 and §6.3.8.  A formula Error cell remains an Error item.
* A `NumberSequence` accepts one or more sequence arguments but does not admit
  a `ReferenceList`.  Therefore an explicit multi-area reference list is a
  pseudotype mismatch and is `#VALUE!` for `MULTINOMIAL`.  This structural
  rejection happens before provider traversal: the evaluator does not read
  cells in the rejected areas to discover hidden errors.  Already-materialized
  direct scalar or inline-array formula Errors in other arguments still have
  their normal source-order precedence; an unread provider cell has no place
  in that precedence order.
* A `NumberSequenceList` has the same cell filtering and additionally admits a
  `ReferenceList`.  Areas are visited in reference-list occurrence order and
  cells in each area are visited row-major.  This is the profile's consistent
  choice of the row/column traversal order left open by §4.11.12.  It is used
  by `GCD` and `LCM`.
* An inline general Array is a value rather than a cell Reference.  This
  profile visits it row-major.  Each element uses the scalar Number bridge:
  finite Number is included, Logical is `0`/`1`, Empty is `0`, finite numeric
  Text is parsed, malformed or non-finite Text is `#VALUE!`, Missing is
  `#VALUE!`, and a formula Error propagates.  This inline-array rule is a
  repository profile; it does not change the reference omission rule.
* A sequence function does not apply scalar implicit intersection to a
  rectangular Reference.  The whole admitted sequence is consumed.

For a sequence argument whose reference contains no admitted values, the
selected identities are:

* `GCD` returns `#NUM!`, because its “at least one positive” constraint cannot
  be met.  If it has admitted values and all are zero, it returns `0`, using
  the explicit §6.16.36 permission.
* `LCM` returns `#NUM!` when the supplied sequence union is empty; there is no
  source-defined empty LCM identity.  A non-empty sequence containing one or
  more zero values returns `0`.
* `MULTINOMIAL` returns `1` for an admitted empty sequence: the empty sum is
  `0`, the empty factorial product is `1`, and `FACT(0) / 1 = 1`.  This is
  distinct from a syntactic call with no sequence argument, which violates
  `+` and is `#VALUE!`.

The value evaluator applies the existing scalar bridge to the scalar Number
functions in this batch.  A rectangular reference or inline array is
evaluated elementwise with the established dimensions and scalar broadcast
rules; a `ReferenceList` in a scalar Number position is `#VALUE!`.  An
incompatible shape is `#VALUE!`; an out-of-shape broadcast position remains
the existing `#N/A` value.  The resolver-free scalar API rejects references
and arrays with its existing typed `UnsupportedKind::Reference` or
`UnsupportedKind::Array` result.  These API boundaries do not change the
sequence rules above.

## Function semantics and conversion order

`COMBIN` computes `C(N,R)` after `INT` conversion.  `C(N,0)` and `C(N,N)` are
`1`; a relation or domain failure is `#NUM!`.  `COMBINA` computes the displayed
repetition formula after the same conversion and strict relation checks, with
the boundary policy given above.  Implementations may use a divide-before-
multiply recurrence, but the resulting exact integer is converted to Number
only once.

`FACT` and `FACTDOUBLE` compute their finite integer products after the
`INT`/domain step.  The double factorial includes its input and decrements by
two.  The mathematical result is converted once to finite binary64; an exact
integer that cannot be represented by a finite Number is `#NUM!`.

`GCD` applies `INT` independently to every selected value and computes the
greatest common divisor of the resulting non-negative integers.  It must not
convert an already computed large GCD through an inexact binary64 arithmetic
step.  `LCM` requires every selected Number to already equal its floor and to
be non-negative; it computes the exact least common multiple.  Zero is
absorbing (`LCM(0;42) = 0`), and the all-zero case is also `0`.  The identity
or empty-input choices are not inferred from a host application's cached
result.

`MULTINOMIAL` is deliberately nested according to the standard's equation:

```text
raw = (a1, a2, ..., an)
numerator = FACT(INT(a1 + a2 + ... + an))
denominator = PRODUCT(FACT(INT(a1)), ..., FACT(INT(an)))
result = numerator / denominator
```

The `INT` on the numerator applies after the raw sum, while each denominator
has its own `INT`.  Thus `MULTINOMIAL(1.5;1.5) = FACT(3)/FACT(1)/FACT(1) = 6`.
Pre-truncating to `(1,1)` and returning `2` changes the specified expression
and is not this profile.  A negative denominator argument floors to a negative
Integer and produces `#NUM!`.  A raw sum that is not finite or whose exact
Integer/result cannot be admitted by the finite numeric profile produces
`#NUM!` after formula-error precedence is resolved.

The raw sum is the exact sum of the already admitted binary64 values, before
the `INT` step.  A conforming bounded implementation decomposes each finite
binary64 value into its sign, integer significand, and power of two, aligns the
terms in a checked fixed-point or equivalent exact accumulator, and applies
floor to that exact represented sum.  It must not add the values in ordinary
binary64 and then floor the rounded accumulator: for example, the represented
values `1.4` and `0.6` have an exact represented sum below `2`, so
`MULTINOMIAL(1.4;0.6) = 1` under this profile even if a particular rounded
addition happens to produce `2.0`.  The accumulator must retain enough signed
fractional residue to distinguish an integer boundary and must report a
checked overflow or a proven final `#NUM!`; it must not silently discard low
bits.  A common implementation is a fixed-point state with the binary64
minimum exponent as its fractional unit and a width derived from the finite
operand range plus the permitted argument count.  An exponent-aligned limb
state with the same exactness is equivalent.  This representation detail is
part of the numeric profile; it is not an arbitrary decimal precision or an
argument-value cutoff.

`EVEN(N)` and `ODD(N)` compute `ceil(abs(N))`, increase by one when its parity
does not match the requested parity, and restore the sign of `N`.  “Away from
zero” applies at exact integers too: `EVEN(2) = 2`, `ODD(2) = 3`,
`EVEN(-2) = -2`, and `ODD(-2) = -3`.  `EVEN(0) = +0` and `ODD(0) = 1`;
negative zero is canonicalized to positive zero before the explicit ODD zero
rule.

`DELTA(X;Y)` compares the converted finite binary64 values for exact equality;
the omitted `Y` is `0`.  It returns numeric `1` or `0`.  `GESTEP(X;Step)`
compares the converted values with `>=`; the omitted `Step` is `0`.  The
§6.16.37 sentence that a non-Number argument is an Error is read together
with §6.2's implicit conversion and the `Number` signature: values that
successfully convert to Number are accepted, while conversion failures are
`#VALUE!`.  This explicitly selects numeric Text, Logical, and Empty-reference
conversion.  It does not permit malformed Text, Missing, formula Error, or
complex values to be treated as zero.

## Formula errors, eager evaluation, and precedence

All arguments are evaluated eagerly in source order under §3.2.3.  Only the
separate lazy functions (`IF`, `IFERROR`, and `IFNA`) may avoid evaluating an
unselected branch.  The optional default is inserted only when the argument
is omitted; it is not a value expression and does not mask an earlier Error.

The evaluator records formula Errors while consuming all supplied arguments so
that the first Error in the conceptual traversal wins.  The traversal order
is:

1. supplied function arguments in source order;
2. sequence areas in `ReferenceList` occurrence order;
3. cells within each area row-major;
4. inline Array elements row-major.

The implementation may read or materialize in another order as long as the
same first formula Error is returned.  A formula Error has precedence over a
generated conversion, domain, shape, arity, or numeric error encountered
later.  If no formula Error exists, the first generated formula error in that
conceptual traversal is returned.  Typed `Unsupported`, `ResourceLimit`,
`Allocation`, `Cancelled`, `SourceChanged`, and
`SourceVersionAvailabilityChanged` failures remain host-level evaluation
failures and are not converted to a formula Error or caught by `IFERROR`.

For example, a malformed Text in an early sequence element followed by a
later `#N/A` returns `#N/A`; a later malformed Text cannot hide an earlier
formula `#DIV/0!`.  A negative domain value and a later formula Error likewise
return the later formula Error after the eager scan.  Arity errors are used
only when no supplied argument formula Error has precedence.

## Finite binary64 and exact integer policy

Part 4 specifies the mathematical equations and does not mandate a host
numeric representation (§3.6 and §6.2).  This repository profile admits only
finite IEEE-754 binary64 Number values at the evaluator boundary.  It uses
fixed-size exact integer state for integer products, GCD/LCM, and combinatorial
intermediates, and converts the final exact integer once with round-to-nearest,
ties-to-even.  This preserves represented integers above `2^53` as inputs and
does not pretend that every integer result is exactly representable in
binary64.  A final result that converts to a finite binary64 Number is a Number
even if it is rounded; a result that cannot convert to finite binary64 is
`#NUM!`.  NaN and infinity never escape the evaluator.

The state size is a resource/profile implementation detail, not a semantic
argument threshold.  It must be selected from the finite binary64 operand and
result requirements and checked with fallible arithmetic.  The contract does
not authorize a guessed decimal cutoff, a blanket `2^53` rejection, or a
fixed factorial limit presented as if it were in ODF.  A coder may use an
exact bounded limb state, logarithmic prechecks, or an equivalent proof of
final representability, provided that a finite result is not rejected merely
because an intermediate factorial or product would overflow binary64.

In particular:

* `C(N,R)` and `COMBINA` should cancel factors before multiplication.  A
  factorial-sized intermediate is not a reason to return `#NUM!` when the
  final binomial is finite.
* `MULTINOMIAL` should use an exact or cancellation-preserving construction.
  It must retain the raw-sum/independent-denominator conversion order above.
  An intermediate representation overflow is a `#NUM!` result only when the
  final mathematical result cannot be admitted; otherwise it is a typed
  resource failure if the caller's work or storage budget is exhausted.
* GCD must remain exact.  LCM must reduce by GCD before multiplication and
  return `#NUM!` only when the exact final LCM cannot be converted to finite
  binary64.  A later zero may mathematically reset an overflowing earlier
  LCM intermediate to zero, but formula-error precedence and bounded work
  still apply; the implementation must not publish a stale overflow.
* `EVEN` and `ODD` must test the exact mathematical output's parity before
  binary64 conversion.  If the required parity integer is not exactly
  representable and conversion would round it to an integer with the wrong
  parity, return `#NUM!` rather than silently returning a parity-violating
  value.  Do not encode this policy as a guessed `2^53` cutoff: derive it from
  the representability and parity of the actual candidate.  For example, an
  unrepresentable odd result at a large binary64 input cannot be rounded to
  the neighboring even value and still claim to implement `ODD`.
* `DELTA` and `GESTEP` compare the already converted binary64 values exactly;
  they do not introduce an epsilon.  Their `0`/`1` results are exact.

Intermediate arithmetic must not leak a non-finite value.  A final finite
subnormal is a valid Number.  Exact zero results are canonicalized to `+0` in
this profile.  The integer-state profile converts every exact result once, so
the retained 361-row numerical oracle compares expected binary64 bits exactly
for all functions whose row has a finite expected result, including
combinatorial products.  Rows with `COMBINA`'s strict-domain or `ODD`'s
unrepresentable-parity profile error expect `#NUM!`; they do not accept the
rounded reference bits retained for audit.

## Evaluation and resource boundary

The public APIs remain the existing scalar and value evaluator entry points.
The functions do not resolve names, read external sources, recalculate
dependencies, update formula caches, or call host services.  Resolver reads
are bounded by the existing cell, work, storage, and cancellation limits.
The sequence reducer may stream selected values and must not retain a
reference or allocate an unbounded vector.  Failure to obtain a required
resource is a typed failure; it must not be converted to `#NUM!` merely to
make a formula complete, and no partial owned result may be published.

## Required validation

The implementation evidence must include both evaluator APIs and every
function above.  The semantic matrix must cover:

* exact arity, omitted optional defaults, explicit Missing, scalar Empty, and
  formula Error values;
* Number, Text, Logical, Empty, Missing, and Error conversion for scalar and
  inline Array inputs;
* NumberSequence versus NumberSequenceList, ReferenceList rejection for
  `MULTINOMIAL`, row-major references, reference omission of Text/Empty/
  Logical, and list order;
* negative fractions for `FACT`, `FACTDOUBLE`, `GCD`, and `MULTINOMIAL`, strict
  fractional rejection for `LCM`, and the raw-sum example
  `MULTINOMIAL(1.5;1.5) = 6`;
* `COMBINA(0;0) = 1`, `COMBINA(0;1) = #NUM!`, `COMBINA(1;0) = 1`, and strict
  rejection of `N < M`;
* all-zero GCD, zero-containing LCM, empty sequence policies, large represented
  integers, final finite overflow, cancellation, and parity-safe EVEN/ODD;
* formula Error precedence over generated conversion/domain errors, eager
  argument evaluation, lazy surrounding `IF`/`IFERROR`/`IFNA`, cancellation,
  resource refusal, and reservation release.

The retained Python oracle uses bigint arithmetic only for integer-valued
observations and is independent of the production kernels.  Its rows are
targeted evidence rather than proof for all finite binary64 inputs.  Native
LibreOffice cached values are corroboration only; host choices such as
`COMBINA(0;0)=0` or fractional `LCM` coercion do not override this contract.
