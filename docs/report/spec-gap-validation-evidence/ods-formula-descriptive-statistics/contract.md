# ODF 1.4 descriptive-statistics evaluator contract

This contract defines the bounded implementation profile for the seven
descriptive-statistics functions in OpenFormula 1.4 Part 4: `AVEDEV`,
`DEVSQ`, `GEOMEAN`, `HARMEAN`, `KURT`, `SKEW`, and `SKEWP`. It is the
semantic boundary for the resolver-free scalar evaluator, the resolver-backed
value evaluator, and their independent validation. It does not authorize
formula-cell recalculation, cached-result use, source refresh, or publication
of a changed cell.

The normative source is the repository-local ODF distribution:

| Source | SHA-256 |
| --- | --- |
| archive `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| member `part4-formula/OpenDocument-v1.4-os-part4-formula.html` | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |

The directly relevant sections are §§3.2.3, 3.3, 3.6–3.7, 4.6–4.11.13,
5.6, 6.1–6.3, 6.13.6, 6.16.46, 6.18.2–6.18.3, 6.18.20, 6.18.34,
6.18.36, 6.18.39, 6.18.67–6.18.68, 6.18.72, and 6.18.74. The evaluator
also retains the bounded conversion, resource, source-fence, and typed-failure
rules selected by the aggregate, statistical-reducer, and dispersion
contracts.

## Signatures and constraints

The signatures retain the ODF pseudotypes. `+` means one or more supplied
arguments in the common function template; a semicolon separates formula
arguments. All seven functions return one `Number` in scalar evaluation.
`#VALUE!` below means `ScalarError::Value`, `#DIV/0!` means
`ScalarError::DivisionByZero`, and `#NUM!` means `ScalarError::Number`.

| Function | Part 4 signature | Mathematical result | Additional constraint or empty profile |
| --- | --- | --- | --- |
| `AVEDEV` | `AVEDEV({NumberSequenceList N}+)` | Mean absolute deviation from the mean | No explicit constraint; no admitted Number follows `AVERAGE(N)` and returns `#DIV/0!`; one Number returns `+0`. |
| `DEVSQ` | `DEVSQ({NumberSequence N}+)` | Sum of squared deviations from `AVERAGE(N)` | No explicit constraint beyond the sequence pseudotype; no admitted Number follows `AVERAGE(N)` and returns `#DIV/0!`; one Number returns `+0`. |
| `GEOMEAN` | `GEOMEAN({NumberSequenceList N}+)` | Geometric mean | No explicit constraint; an empty admitted sequence has zero count and returns `#DIV/0!`; a zero product returns `+0`; a negative product has a real root only for odd `n` and otherwise returns `#NUM!`. |
| `HARMEAN` | `HARMEAN({NumberSequenceList N}+)` | Harmonic mean | No explicit constraint; an empty sequence or an exact zero reciprocal denominator returns `#DIV/0!`; every admitted Number must be nonzero, so a zero input also returns `#DIV/0!`. Signed nonzero Numbers are admitted. |
| `KURT` | `KURT({NumberSequenceList X}+)` | Bias-corrected sample excess kurtosis | `COUNT(X) >= 4` and sample `STDEV(X) != 0`; a violation returns `#VALUE!`. |
| `SKEW` | `SKEW({NumberSequenceList Sample}+)` | Bias-corrected sample skewness | Part 4 requires at least three Numbers; a count below three or a zero sample standard deviation is the selected `#VALUE!` profile. |
| `SKEWP` | `SKEWP({NumberSequence Population}+)` | Population skewness | Part 4 requires at least three Numbers; a count below three or a zero population standard deviation is the selected `#VALUE!` profile. |

The specification describes constraints as an `Error` and does not require a
particular subtype for these invalid or undefined cases. This repository
selects `#VALUE!` for arity, pseudotype, conversion, count, and degenerate
standard-deviation failures. `AVEDEV` and `DEVSQ` explicitly derive their
center from `AVERAGE`; the existing `AVERAGE` profile is `#DIV/0!` when no
Number is admitted, so that subtype is retained for their empty sequence.
`GEOMEAN` uses `1/n` and `HARMEAN` divides by a reciprocal sum; an empty set
therefore uses the selected division-by-zero profile. A zero `HARMEAN` member
is an undefined reciprocal and has the same `#DIV/0!` result. A non-finite
input or non-finite final numeric result is `#NUM!`.

The `+` arity is enforced. A zero-argument call such as `KURT()` is a formula
`#VALUE!` arity result. A syntactically supplied missing slot, such as
`SKEW(;)`, is a supplied argument and is normalized to a formula `#VALUE!`
conversion error; it is not an omitted zero-argument call. An admitted empty
Reference or an array with no Number members is different from an omitted
argument and follows the function row above. Formula Errors already present
in an admitted argument take precedence over the generated empty, count,
domain, or arithmetic result.

The following examples make the signed geometric and harmonic choices
explicit:

* `GEOMEAN(-8)` is `-8`; `GEOMEAN(-2;-8)` is the positive square root of 16;
  a negative product with odd `n` uses its real odd root. A negative product
  with even `n` has no real result and is `#NUM!`. Any zero factor makes the
  product zero and publishes canonical `+0`, regardless of the signs of the
  other factors.
* `HARMEAN(-2;-4)` is `-8/3`. `HARMEAN(3;6;-2)` has the exact reciprocal
  denominator `1/3 + 1/6 - 1/2 = 0` and is `#DIV/0!`; no arbitrary near-zero
  tolerance may turn a valid nonzero denominator into an error. Any zero
  member is `#DIV/0!`.

## Equations and numeric profile

For the admitted finite Number sequence `x₁, …, xₙ`, let
`x̄ = (1/n) Σᵢ xᵢ`. The normative equations are:

```text
AVEDEV(N) = (1/n) Σᵢ |xᵢ - x̄|
DEVSQ(N)  = Σᵢ (xᵢ - a)², where a = AVERAGE(N)
GEOMEAN  = (Πᵢ xᵢ)^(1/n)
HARMEAN  = n / Σᵢ (1/xᵢ)
```

For `KURT`, `s` is the sample standard deviation and the result is:

```text
KURT(X) = n(n+1) / ((n-1)(n-2)(n-3))
          * Σᵢ ((xᵢ - x̄) / s)^4
          - 3(n-1)² / ((n-2)(n-3))
```

For sample skewness, `s` is the sample standard deviation and:

```text
SKEW(Sample) = n / ((n-1)(n-2))
               * Σᵢ ((xᵢ - x̄) / s)^3
```

For population skewness, `σ` is the population standard deviation and:

```text
SKEWP(Population) = (1/n) * Σᵢ ((xᵢ - x̄) / σ)^3
```

`KURT` therefore uses the sample denominator `n-1` for `s`, while `SKEWP`
uses the population denominator `n` for `σ`. `SKEW` uses the sample estimate
and its `n/((n-1)(n-2))` bias correction. `DEVSQ` is not a sample or
population variance: it returns the unnormalized sum of squared deviations.

The formulas define the mathematical result, not a required algorithm. This
finite binary64 profile must avoid an intermediate overflow when the final
finite result is representable:

* `AVEDEV` uses an exact or compensated mean and a scaled absolute-deviation
  pass. The sum of absolute deviations may exceed binary64 even when dividing
  by `n` would produce a finite answer; that intermediate overflow must not be
  reported as `#NUM!`.
* `DEVSQ` may use the exact fixed-width dyadic moment state described below,
  including raw sums `S1..S4` and exact centered numerators `C2`, `C3`, and
  `C4`, rather than a rounded mean followed by a naive residual pass. A
  one-member sequence returns canonical `+0`. A negative residual caused only
  by roundoff is clipped to zero; a genuinely non-finite final result is
  `#NUM!`.
* `GEOMEAN` tracks the sign parity and zero state separately and computes the
  logarithmic mean of absolute magnitudes or an equivalent stable root. It
  must not multiply the input values into an overflowing product merely to
  take the root. The real odd-root profile above applies to a negative
  product. A nonzero result that underflows retains its IEEE sign; an exact
  zero product is canonical `+0`.
* `HARMEAN` uses a fixed-size extended or compensated reciprocal accumulator
  with a forward error bound on its denominator. Reciprocal terms may be much
  larger than binary64 even when the final harmonic mean is finite or
  underflows to zero. If that bound cannot classify the denominator as
  exactly zero or nonzero, or cannot establish sufficient sign and numerical
  accuracy for the final quotient, the evaluator may replay the complete
  descriptor into a fallibly growing adaptive exact-rational aggregate. The
  aggregate retains only rational limbs and checked count state; it never
  retains a per-cell value buffer. Its growth is charged to work and storage
  limits.
  It must not reject a finite nonzero input or a finite nonzero denominator
  using an arbitrary semantic epsilon or a fixed precision cap. An exact zero
  denominator is `#DIV/0!`; a nonzero denominator with a non-finite final
  result is `#NUM!`.
* `KURT`, `SKEW`, and `SKEWP` may use the same exact fixed-width dyadic
  moment state, retaining raw `S1..S4` and exact centered `C2`, `C3`, and `C4`
  numerators without retaining an input vector. They must not form an
  avoidably overflowing `xᵢ-x̄`, power, or correction factor when the
  represented final result is finite. Their exact algebraic zero is canonical
  `+0`; a negative nonzero result that rounds to underflowed `-0` retains that
  sign.

The implementation may use the fixed-width exact dyadic moment helper above,
an exact finite-binary64 sum where the shared aggregate profile provides one,
compensated moments, scaling, or another bounded method. The exact helper is
an aggregate state only: it retains no input vector, and its checked fixed
width is part of the finite-number profile rather than an accuracy epsilon.
It must reject NaN and infinity at the Number boundary and
must never publish NaN or infinity. Ordinary last-bit differences from the
chosen finite binary64 algorithm are validated with the repository's documented
finite-number tolerance; the contract does not claim decimal or cross-platform
bit identity for every moment.

## Pseudotypes, coercion, and references

`NumberSequence` and `NumberSequenceList` preserve the ODF distinction between
a scalar conversion and a reference sequence. This profile uses the existing
finite, locale-independent decimal bridge for direct scalar Text, which is a
permitted choice under §6.3.5; malformed, NaN, or infinite Text produces a
generated formula `#VALUE!`.

### NumberSequence

`DEVSQ` and `SKEWP` consume `NumberSequence`. The conversion is:

* A scalar Number contributes one finite Number. A scalar Logical contributes
  `0` or `1`. A scalar Text is parsed by the profile's decimal bridge. A
  scalar Empty follows the repository scalar bridge and contributes `0`; an
  explicit Missing value produces `#VALUE!`.
* A single logical Reference, including a three-dimensional cuboid, contributes
  only Number and formula Error cells. Referenced Empty, Text, and
  distinguished Logical cells are omitted. Numeric-looking Text in a
  Reference is still omitted; it is not parsed cell by cell.
* An explicit `ReferenceList` is not a `NumberSequence`. `DEVSQ([.A1]~[.B1])`
  and `SKEWP([.A1]~[.B1])` return `#VALUE!` before any resolver cell read.
  Separate Reference arguments remain valid and are concatenated in argument
  order. A computed expression may first have to produce a descriptor; once
  the list shape is known, the reducer refuses before scanning its cells.
* An inline rectangular Array is the repository's already-valued sequence
  extension. Elements are visited row-major: Number is included, Logical is
  `0`/`1`, Empty is `0`, finite numeric Text is parsed, malformed or
  non-finite Text is generated `#VALUE!`, Missing is `#VALUE!`, and a formula
  Error is retained. Complex values are outside this finite real profile and
  produce `#VALUE!`.

### NumberSequenceList

`AVEDEV`, `GEOMEAN`, `HARMEAN`, `KURT`, and `SKEW` consume
`NumberSequenceList`. They use the same scalar, Reference, and inline Array
conversion as `NumberSequence`, with one addition: an ordered `ReferenceList`
is admitted. Each Reference in the list is converted to a NumberSequence in
occurrence order. A cuboid is traversed sheet by sheet and then row-major
within each sheet; repeated or overlapping occurrences are retained. The
reducer never applies implicit intersection to an admitted sequence
Reference.

ODF allows row-major or column-major order within a sheet. Row-major is the
repository choice so both evaluator profiles and read-budget evidence have a
deterministic order. Formula Error precedence uses that same conceptual order.

The resolver-free scalar evaluator admits the scalar forms above. A Reference
or ReferenceList that it cannot retain, and an Array when its scalar API cannot
retain the matrix descriptor, produce the existing typed
`UnsupportedKind::Reference` or `UnsupportedKind::Array` refusal. The
resolver-backed value evaluator handles structurally admitted References,
ReferenceLists, and Arrays. Neither evaluator recalculates a formula cell or
imports a cached formula result.

A Complex value is not silently projected to a real component. A direct
Complex value, an inline Complex element, or a Complex provider value becomes
the generated `#VALUE!` profile error. Empty worksheet cells inside a
Reference are omitted by NumberSequence conversion; scalar Empty and inline
Array Empty use the existing value-bridge zero extension above.

## Function-specific behavior

### AVEDEV and DEVSQ

`AVEDEV` first obtains the average of all admitted Numbers and then averages
the absolute deviations from that mean. Its empty behavior is therefore the
same `#DIV/0!` profile as `AVERAGE`. `DEVSQ` uses the same center because Part
4 explicitly defines `a` as `AVERAGE(N)`. A one-member `DEVSQ` sequence has a
zero deviation sum and returns canonical `+0`.

`DEVSQ` and `SKEWP` are the two singular-reference functions in this batch.
A union expression passed as one argument is a shape error, while multiple
separate logical References are separate sequence arguments. `AVEDEV` and the
other `NumberSequenceList` functions flatten an explicit ReferenceList in its
written occurrence order.

### GEOMEAN

`GEOMEAN` counts admitted finite Numbers, including signed values. It does
not adopt a spreadsheet-host positive-only restriction because §6.18.34
specifies no such constraint. The product sign is the parity of negative
nonzero inputs. A zero factor makes the exact product zero and returns `+0`.
Without a zero factor, an even number of negative factors gives the positive
root; an odd negative product has a real negative root only when `n` is odd.
For even `n` and a negative product, the real-valued result is undefined and
returns `#NUM!`. This covers the singleton negative case without a special
positive-only exception.

The implementation may use `exp(Σ ln(|xᵢ|)/n)` for nonzero magnitudes, with the
sign and parity applied afterward. It must retain a finite result across
avoidable product overflow and must classify exact zero separately from a
nonzero underflow.

### HARMEAN

`HARMEAN` admits signed finite nonzero Numbers. It does not impose the
positive-only restriction used by some host applications. Every term
`1/xᵢ` must be defined; a zero input returns `#DIV/0!`. The reciprocal sum is
allowed to cancel between positive and negative terms. An exact zero sum is
`#DIV/0!`; a nonzero sum, however small, is valid and is not rejected by an
arbitrary tolerance. The final sign follows `n / denominator` and may be
negative.

### KURT

`KURT` is sample excess kurtosis, with the `n+1`, `n-1`, `n-2`, and `n-3`
factors shown in the equation above. At least four admitted Numbers and a
nonzero sample standard deviation are required. A constant four-or-more
sequence is a degenerate `#VALUE!` constraint failure, not a NaN or an
infinite result. The implementation must not replace the sample standard
deviation with the population denominator.

### SKEW and SKEWP

`SKEW` uses the sample standard deviation and the sample bias correction;
`SKEWP` uses the population standard deviation and no sample correction. Both
require at least three admitted Numbers. A constant sequence has zero sample
or population standard deviation and returns the selected `#VALUE!`
degenerate-constraint result. The two functions must remain distinct for
nonconstant data; they are not aliases differing only by spelling.

## Matrix projection and demand-cache invariants

None of these seven signatures contains a scalar Number, Integer, Criterion,
or `ForceArray` parameter. Their sequence arguments are complete sequence
arguments in every evaluator mode. A multi-cell Reference is therefore not
implicitly intersected at the current output coordinate, and an admitted
ReferenceList is not projected to one cell. `DEVSQ` and `SKEWP` retain their
pre-scan list-shape refusal in matrix mode as well as scalar mode.

The §3.3 rule for a function that returns an Array applies to that producer's
own scalar inputs. For example, a nested `MUNIT` call uses its `[0,0]` size
input when the array-returning call itself is evaluated in matrix mode. Once
that call has produced an Array, this contract's existing inline-Array
sequence bridge consumes its elements row-major as one complete sequence; the
producer's result is not confused with a scalar parameter of the descriptive
reducer. A scalar conditional `Criterion` remains position-sensitive, and
`MUNIT` stays excluded from full-argument demand propagation.

A descriptive reducer over a fixed Reference, ReferenceList, literal Array, or
invariant nested reducer is cacheable only after its complete sequence
descriptor is classified coordinate-independent. A computed reducer under a
projected lazy `IF` remains conservative unless its complete sequence is
proven invariant. Statistical nested-criterion propagation may carry complete
sequence arguments through the projected branch, but it must not promote the
position-sensitive `MUNIT` expression into a full-argument cache entry.

Cache entries contain only a scalar Number or formula-error payload and the
source/context identity. A typed evaluator failure is never converted to a
formula Error or cached as one. No descriptive reducer returns a matrix merely
because it was evaluated in matrix mode; an enclosing array-producing
expression applies its own output-shape rules to the already-computed scalar.

## Formula-error precedence and typed failures

Arguments are evaluated eagerly in source order. The conceptual sequence order
is argument order, ReferenceList occurrence order, sheet order for a cuboid,
and row-major cell order within each sheet. The first formula Error in that
order is retained. The primary admitted scan always completes after retaining
a formula Error, so a later typed resolver, source, resource, cancellation,
or allocation failure can supersede the retained formula result. A replay is
performed only when it is still required to produce or classify a result,
such as a centered statistic's scheduled second pass or the uncertain-
denominator/quotient HARMEAN fallback; once the formula Error is already
final and no such typed operation remains, an otherwise meaningless replay
is skipped. If a replay has been scheduled, it continues to observe the same
typed-failure precedence, budgets, cancellation checks, and source fences.

If no source Formula Error exists, the first generated formula error in
conceptual order is returned. Generated conversion, empty, count, degenerate,
zero-denominator, negative-even-root, and non-finite-result errors use the
profile above. A formula Error is never converted into a Number merely to
satisfy a count or moment constraint.

Typed `Unsupported`, cancellation, source-version changes, resource limits,
and allocation failures remain `EvaluationFailure` values. They are never
converted into formula Errors and are never caught by `IFERROR` or `IFNA`.
The value evaluator must preserve the source fence before and after the full
calculation and check cancellation before publication.

## Resource, precision, and memory boundaries

The value evaluator streams each admitted Reference and ReferenceList cell by
cell. It retains fixed-size descriptive state, checked counts, and bounded
reference metadata. It must:

* validate area/list geometry and checked row/column/cell products before a
  scan;
* charge AST work, argument conversion, every physically inspected cell,
  retained Text bytes, every reducer operation, and every replayed pass to the
  caller's hierarchical budget;
* call `charge_cell_work` before `read_reference_cell`, enforce the cumulative
  `max_reference_cells` ceiling across arguments, list occurrences, and replay
  passes, and check cancellation before and after each resolver read;
* use `read_to_element` with borrowed provider Text, charging Text bytes
  without cloning one owned string per observation;
* use fallible capacity for AST arrays, bounded reference metadata, and the
  explicitly permitted adaptive exact-rational harmonic aggregate limbs; the
  scalar reducer must not materialize a range or retain a per-cell value
  vector. The adaptive state is bounded by the caller's hierarchical
  storage/work limits, not by an implementation-chosen precision ceiling;
* use function-specific pass counts. `DEVSQ`, `KURT`, `SKEW`, and `SKEWP` may
  complete in one admitted scan with the exact fixed-width dyadic moment state
  (`S1..S4`, `C2`, `C3`, `C4`) and no input vector. `AVEDEV` requires an exact
  mean pass followed by a charged descriptor replay for its absolute
  deviations unless an equivalent bounded state is proven. Any replay is a
  real resolver scan with its own work/read/cancellation charges, and budgets
  are not reset between passes. `HARMEAN` may similarly replay for the
  adaptive exact-rational fallback described above;
* retain the complete ordered descriptor for the two-pass operation so a
  projected matrix branch does not silently re-evaluate a range at one
  output coordinate. If an expression is position-dependent, the demand-cache
  classifier decides whether the required re-evaluation can be reused;
* check source identity/version before and after all passes and publish no
  result after any typed failure or fence violation; and
* release all temporary reservations on every formula or typed failure path.

`GEOMEAN` can use one fixed-state pass. `DEVSQ`, `KURT`, `SKEW`, and `SKEWP`
may use one scan through the exact dyadic moment helper; `AVEDEV` uses its
exact-mean pass and charged replay. `HARMEAN` normally uses one pass and
only enters the adaptive exact-rational fallback when its forward error bound
cannot classify the reciprocal denominator or cannot guarantee sufficient sign
and quotient accuracy; that fallback is a charged descriptor replay and may
fail with a typed allocation/resource result. The other reducers may replay
references, but they must not reset
`max_reference_cells` or any hierarchical work/storage budget for a second
pass. A `ReferenceList` passed to `DEVSQ` or `SKEWP` is a shape preflight and
performs zero resolver reads when its descriptor is already available. Formula
cells are not recalculated and cached results are not refreshed.

## Native observations and required evidence

The native receipt under
`docs/report/spec-gap-validation-evidence/ods-formula-descriptive-statistics/native/`
contains bounded LibreOffice FODS observations from core commit
`d804d6aff49054bad1719ec3c2d136b545bbc7e7`. Its 39 finite cached values are
useful corroboration for ordinary positive data and literal-array/reference
closures. They are not normative input and do not override this contract.
In particular, the native fixtures observe positive-only or host-specific
behavior such as `GEOMEAN` returning `#NUM!` for negative data, a zero-input
`HARMEAN` generic invalid error, and `#DIV/0!` or generic invalid errors for
constant or undersized `KURT`/`SKEW`/`SKEWP`. ODF §6.18.34 and §6.18.36 state
no positive-input constraint, so this profile admits signed values and gives
the explicit real-root/reciprocal-denominator results above. Native error
spellings are retained as host observations only.

Validation must cover every function through scalar values, inline Arrays,
empty and mixed References, a supported 3-D Reference, and an explicit
ReferenceList where the signature admits it. It must include:

* direct Number, numeric and malformed Text, Logical, Empty, Missing,
  Complex, non-finite, and formula Error values;
* the NumberSequence versus NumberSequenceList distinction, including
  numeric-looking Text and distinguished Logical values inside References;
* zero, negative, and mixed-sign geometric/harmonic inputs, odd/even negative
  products, zero harmonic members, exact reciprocal cancellation, and valid
  arbitrarily small nonzero denominators;
* AVEDEV/DEVSQ one-member and empty sequences; KURT's n=4 boundary and
  sample-standard-deviation degeneracy; SKEW/SKEWP's n=3 boundary and
  sample/population distinction;
* exact zero versus negative underflowed `-0`, finite extreme values,
  cancellation, avoidable product/reciprocal/moment overflow, and final
  non-finite `#NUM!` results;
* first formula-error order, typed resolver/resource/cancellation precedence,
  source-version fences, borrowed Text, cumulative read/work limits across
  replay passes, and zero-read singular ReferenceList refusal; and
* function-specific pass-count receipts: one reference scan for `DEVSQ`,
  `KURT`, `SKEW`, and `SKEWP` when the exact dyadic moment helper is used;
  an exact-mean scan plus one charged replay for `AVEDEV`; one ordinary scan
  for `GEOMEAN` and `HARMEAN`; and at most one additional charged HARMEAN
  fallback replay when denominator or quotient accuracy is uncertain. The
  receipts must also show that a retained formula Error completes the primary
  scan without causing a gratuitous replay; and
* projected lazy branches that preserve complete sequence arguments, reuse an
  invariant descriptive reducer, and leave nested `MUNIT` scalar criteria
  position-sensitive.

The independent oracle must evaluate the equations using a wider or exact
intermediate where needed. Native caches can corroborate ordinary finite
results but cannot redefine signed-domain, empty-set, error-subtype, precision,
or resource behavior selected by this contract.
