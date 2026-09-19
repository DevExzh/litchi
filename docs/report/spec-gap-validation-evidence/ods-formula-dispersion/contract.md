# ODF 1.4 dispersion evaluator contract

This contract defines the bounded implementation profile for the eight
dispersion functions in OpenFormula 1.4 Part 4: `VAR`, `VARA`, `VARP`,
`VARPA`, `STDEV`, `STDEVA`, `STDEVP`, and `STDEVPA`. It is the semantic
boundary for the resolver-free scalar evaluator, the resolver-backed value
evaluator, and their independent validation. It does not authorize formula
cell recalculation, cached-result use, source refresh, or publication of a
changed cell.

The normative source is the repository-local ODF distribution:

| Source | SHA-256 |
| --- | --- |
| archive `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| member `part4-formula/OpenDocument-v1.4-os-part4-formula.html` | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |

The directly relevant sections are §§3.2.3, 3.3, 3.6–3.7, 4.6–4.11.13,
6.1–6.3, 6.13.6–6.13.7, 6.18.72–6.18.75, and 6.18.82–6.18.85. The
implementation also reuses the existing bounded database variance kernel and
the resource rules in ADR 0004, ADR 0005, ADR 0006, ADR 0008, and ADR 0024.

## Signatures and count constraints

The signatures retain the ODF pseudotypes. `+` means one or more supplied
arguments in the common function template; a semicolon separates formula
arguments. `#VALUE!` is `ScalarError::Value`, `#NUM!` is
`ScalarError::Number`, and `#DIV/0!` is `ScalarError::DivisionByZero`.

| Function | Part 4 signature | Mathematical result | Minimum admitted members |
| --- | --- | --- | --- |
| `VAR` | `VAR({NumberSequence N}+)` | Sample variance `s²` | 2 Numbers |
| `VARA` | `VARA({Any Sample}+)` | Sample variance after the A conversions | 2 converted members |
| `VARP` | `VARP({NumberSequence N}+)` | Population variance `σ²` | 1 Number; one returns `+0` |
| `VARPA` | `VARPA({Any Sample}+)` | Population variance after the A conversions | 1 converted member; one returns `+0` |
| `STDEV` | `STDEV({NumberSequenceList N}+)` | Sample standard deviation `s` | 2 Numbers |
| `STDEVA` | `STDEVA({Any Sample}+)` | Sample standard deviation after the A conversions | 2 converted members |
| `STDEVP` | `STDEVP({NumberSequence N}+)` | Population standard deviation `σ` | 1 Number; one returns `+0` |
| `STDEVPA` | `STDEVPA({Any Sample}+)` | Population standard deviation after the A conversions | 1 converted member; one returns `+0` |

Part 4 describes a failed count constraint as an `Error`, without requiring a
particular error subtype. This repository chooses formula `#VALUE!` for an
empty population, a sample with fewer than two admitted members, a missing
required variadic argument, and other pseudotype/shape failures. That choice
matches the existing database `DVAR`/`DVARP`/`DSTDEV`/`DSTDEVP` profile and the
shared `VarianceAccumulator` denominator failure. The local LibreOffice
fixtures use `#DIV/0!` for several zero/insufficient calls; those are retained
as host-compatibility observations and do not change this normative repository
choice.

The empty-set rules are therefore:

* `VAR`, `VARA`, `STDEV`, and `STDEVA` return `#VALUE!` unless at least two
  members survive conversion and error handling.
* `VARP`, `VARPA`, `STDEVP`, and `STDEVPA` return `#VALUE!` when no member
  survives conversion. A single admitted member returns canonical `+0` for
  both its variance and standard deviation.
* All eight functions with zero supplied arguments return `#VALUE!`. A call
  with explicit missing slots, such as `VARP(;)`, has supplied arguments;
  each missing slot is normalized to a formula `#VALUE!` before the reducer
  observes it. It is not an omitted zero-argument call.
* A formula Error supplied by an argument or read from an admitted reference
  is retained and returned before the empty/count result. Thus an Error does
  not become a numeric member merely to satisfy a minimum count.

Each successful reducer publishes one scalar finite Number. A direct reducer
does not broadcast or materialize a result matrix; an enclosing projected
operator may broadcast that already-computed scalar under its own rules.

## Dispersion equations and numeric profile

For the admitted converted sequence `x₁, …, xₙ`, the mathematical definitions
are:

```
x̄       = (1 / n) * Σᵢ xᵢ
s²       = (1 / (n - 1)) * Σᵢ (xᵢ - x̄)²
σ²       = (1 / n)       * Σᵢ (xᵢ - x̄)²
s        = sqrt(s²)
σ        = sqrt(σ²)
```

`VAR` and `VARA` use `s²`; `VARP` and `VARPA` use `σ²`; the corresponding
`STDEV*` function uses the square root of the same sample or population
quantity. Variance is non-negative. A mathematically zero variance or
standard deviation is published as canonical `+0`.

The eight functions reuse `NumericAggregate` and its existing
`VarianceAccumulator` operations:

* `SampleVariance` and `PopulationVariance` implement `VAR`/`VARP` and their
  A variants.
* `SampleStandardDeviation` and `PopulationStandardDeviation` implement the
  four `STDEV*` functions.
* The state is fixed-size and allocation-free. It retains a checked count, a
  first-value offset, a scale, compensated first and second centered moments,
  and no per-member vector.
* Values are normalized by the largest absolute value seen so far. The first
  value remains the origin; the state is rescaled before a larger magnitude is
  admitted. The compensated centered moments preserve cancellation and
  adjacent representable differences better than a naive sum-of-squares
  implementation.
* At publication, variance uses the checked denominator `n - 1` or `n` and
  scales the normalized second centered moment. Standard deviation computes
  the scaled square root directly. This direct standard-deviation path can
  remain finite when the corresponding sample variance would overflow.
* Non-finite input Numbers and non-finite final results are formula `#NUM!`.
  An intermediate scaled product is not itself a failure when the final
  standard deviation remains finite. A negative centered residual caused by
  roundoff is clipped to zero; a genuinely non-finite state remains `#NUM!`.

This is a compensated finite-binary64 profile, not an exact rational or
one-final-rounding guarantee. Stable ordinary results may be asserted by exact
bits where the operation is invariant; values affected by the compensated
moment path are compared using the repository's documented finite-number
tolerance. The kernel must retain cancellation, adjacent-large-value deltas,
and finite standard deviations across a variance overflow, as covered by the
existing database vectors (for example, `1e154` and `-1e154`). It must not
silently emit NaN or infinity.

## Pseudotypes, scalar arguments, arrays, and references

`NumberSequence`, `NumberSequenceList`, and `Any` preserve the ODF distinction
between a scalar conversion and a reference sequence. The repository profile
uses a finite, locale-independent decimal bridge for direct scalar Text; this
is an implementation choice under §6.3.5, which otherwise leaves Text
conversion host-defined.

### NumberSequence and NumberSequenceList

`VAR`, `VARP`, and `STDEVP` consume `NumberSequence`. `STDEV` consumes
`NumberSequenceList`.

The prose summary of §6.18.74 describes `STDEVP` as including Text and
Logical values, while its signature is `NumberSequence` and §6.3.7 says that a
referenced sequence contains only Numbers and Errors (with distinguished
Logical cells omitted). This profile follows the pseudotype conversion and
therefore omits referenced Text, Empty, and distinguished Logical cells for
`STDEVP`, as it does for `VAR` and `VARP`; direct scalar and Array conversion
still follows the scalar rules below.

* A scalar Number contributes one Number. A scalar Logical contributes `0` or
  `1`. A scalar Text is parsed as a finite decimal Number; malformed, NaN, or
  infinite Text produces a generated formula `#VALUE!`. A scalar Empty is
  converted to numeric `0`; an explicit Missing value is a formula `#VALUE!`.
* A single logical Reference, including a three-dimensional cuboid, contributes
  only Number and formula Error cells. Referenced Empty, Text, and
  distinguished Logical cells are omitted. A referenced Text that looks
  numeric is still omitted; it is not parsed cell by cell.
* `STDEV` additionally accepts an ordered ReferenceList. Each occurrence is
  converted to a NumberSequence in list order. `VAR`, `VARP`, and `STDEVP`
  reject an explicit ReferenceList as a pseudotype/shape `#VALUE!` before any
  resolver cell read. They do not flatten the list or apply implicit
  intersection.
* An inline rectangular Array is an already-valued sequence in this profile.
  Its elements are visited row-major and use scalar conversion: Number is
  included, Logical is `0`/`1`, Empty is `0`, finite numeric Text is parsed,
  Missing is `#VALUE!`, formula Error is retained, and Complex is `#VALUE!`.
  This Array extension is separate from the omission rules for referenced
  cells.

### Any and the A variants

`VARA`, `VARPA`, `STDEVA`, and `STDEVPA` consume `Any`.

* Number contributes its finite value. Text contributes numeric `0`, including
  empty Text. Logical contributes `0` or `1`. Empty cells and scalar Empty
  values are omitted. Missing is a formula `#VALUE!`.
* A Reference contributes Number, Text-as-zero, and Logical-as-zero-or-one;
  Empty is omitted. Formula Errors remain formula Errors. An ordered
  ReferenceList is admitted for all four A variants under this repository
  profile; its areas are traversed in occurrence order. The explicit
  ReferenceList wording in §§6.18.83 and 6.18.85 is applied consistently to
  `STDEVA` as the same `Any` family profile.
* Inline Array elements use the same A conversion in row-major order. Complex
  values are outside this finite real profile and produce formula `#VALUE!`;
  they are not projected to a real component.
* The sample A constraints are measured after the conversion above. Text and
  Logical members count because they become zero or one; Empty does not count.
  A formula Error wins as a formula result before a count decision. This is
  equivalent to the normative `COUNTA(Sample)>1` condition for the admitted
  Any values while keeping unsupported Complex values as explicit conversion
  errors.

The resolver-free scalar evaluator admits the scalar forms above. A Reference
or ReferenceList, and an Array when the scalar evaluator cannot retain its
matrix descriptor, produce the existing typed `UnsupportedKind::Reference` or
`UnsupportedKind::Array` refusal. The value evaluator is the resolver-backed
path for references and arrays; neither evaluator recalculates a formula cell.

## Ordered traversal and formula-error precedence

Arguments are evaluated eagerly in source order. The conceptual sequence order
is argument order, ReferenceList occurrence order, three-dimensional sheet
order, and row-major order within each area. ODF permits either row-major or
column-major processing inside a sheet; row-major is the repository choice so
both evaluator profiles and read-budget evidence have one deterministic order.

All eight functions propagate formula Errors. The evaluator retains the first
formula Error in conceptual order while continuing a bounded admitted scan.
Continuing is required so a later typed resolver/source/resource failure can
supersede the retained formula value. If no formula Error exists, the first
generated formula Error in conceptual order is returned. Shape refusal,
malformed scalar Text, Complex conversion, Missing, and insufficient count all
use the selected `#VALUE!` profile; non-finite numeric state uses `#NUM!`.

Typed `Unsupported`, cancellation, source-version changes, resource limits,
and allocation failures remain `EvaluationFailure` values. They are never
converted into formula Errors and are never caught by `IFERROR` or `IFNA`.
The first formula Error therefore does not permit the scan to stop before
`charge_cell_work`, resolver reads, cancellation checks, and source fences have
completed their required work.

## Matrix projection and demand-cache invariants

Dispersion reducers consume complete sequence/reference descriptors inside a
projected matrix branch. A multi-cell Reference must not be implicitly
intersected at the current output position.

* `STDEV` and the A variants preserve complete ordered ReferenceList/Reference
  descriptors because their pseudotypes admit them. `VAR`, `VARP`, and
  `STDEVP` reject a `RuntimeAreaSet` marked as a list before scanning, yielding
  `#VALUE!` with zero resolver reads.
* A scalar result is cacheable only when every direct scalar descendant is
  coordinate-independent under the enclosing projection. Fixed references,
  literal arrays, and invariant nested dispersion reducers may use the demand
  cache. A computed reducer containing a projected multi-cell reference stays
  position-dependent unless its complete descriptor is proven invariant.
* Statistical reducer propagation may carry full arguments through a projected
  lazy `IF` and through a nested statistical criterion. `MUNIT` remains
  position-sensitive and is excluded from full-argument propagation; a nested
  MUNIT scalar parameter must be evaluated at the current output coordinate.
* Cache entries contain only the scalar Number or formula-error payload and the
  required source/context identity. They must not retain a cell vector or hide
  a typed failure.

Cache classification is bounded and iterative. It must not recurse through an
unbounded formula AST or turn a position-dependent nested expression into an
invariant value merely because its first output coordinate happened to match.

## Resource, memory, and resolver boundaries

The value evaluator streams every admitted Reference and ReferenceList cell by
cell. It retains only fixed variance state, checked counters, and bounded
reference metadata. It must:

* validate area/list geometry and checked row/column/cell products before a
  scan;
* charge AST work, argument conversion, every physically inspected cell,
  retained Text bytes, and each reducer operation to the caller's hierarchical
  budget;
* call `charge_cell_work` before `read_reference_cell`, enforce the cumulative
  `max_reference_cells` ceiling, and check cancellation before and after each
  resolver read;
* use `read_to_element` with borrowed provider Text, charging Text bytes
  without cloning one owned string per observation;
* use fallible capacity for retained area/reference metadata and never build a
  full cell vector for a scalar result;
* check source identity/version before and after evaluation and cancellation
  before publishing the result; and
* publish no result after any typed failure or source-version fence violation.

The NumberSequence list refusal for `VAR`, `VARP`, and `STDEVP` is a shape
preflight and therefore performs zero resolver reads when the descriptor is
already available. A computed expression may first need to evaluate into a
reference descriptor under its own bounded rules; once the list shape is
known, the dispersion reducer must refuse before scanning its cells.

Formula cells are not recalculated and cached results are not refreshed. The
resolver is immutable and read-only for the duration of evaluation. No
variance or standard-deviation reducer may allocate a per-cell value vector,
reset a budget per argument, or publish a partial result.

## Required evidence and regression vectors

The implementation review must cover every function through scalar values,
inline arrays, a single rectangular Reference, a 3-D Reference where the
resolver supports it, and an explicit ReferenceList where the signature
admits it. It must include:

* ordinary sample/population vectors such as `{2; 6; 4}` and one-member
  population cases, with exact formulas checked against the selected numeric
  tolerance;
* all-Empty references, scalar Empty, empty Text, malformed numeric Text,
  Logical values, Complex values, Missing slots, and formula Errors;
* the distinction between a referenced numeric-looking Text (omitted by
  NumberSequence) and direct/array numeric-looking Text (parsed), plus the A
  variants' Text-as-zero rule;
* explicit ReferenceList acceptance for STDEV and all A variants, and zero-read
  `#VALUE!` refusal for ReferenceList passed to VAR, VARP, or STDEVP;
* first formula-error order, typed resolver/resource/cancellation precedence,
  source-version fences, borrowed Text, and cumulative read/work limits;
* projected `IF` cases that retain complete references, cache invariant nested
  reducers, and recompute a position-sensitive nested `MUNIT` criterion;
* cancellation and each relevant reference/metadata limit during a long scan;
  and
* extreme finite values, cancellation, adjacent large values, a finite
  standard deviation whose variance overflows, and canonical zero results.

Native LibreOffice FODS files under
`3rdparty/libreoffice-core/sc/qa/unit/data/functions/statistical/fods/` are
useful corroboration for ordinary values and host-specific behavior. They do
not override the local ODF pseudotype rules, the chosen `#VALUE!` constraint
subtype, or the repository's no-recalculation/resource contract.
