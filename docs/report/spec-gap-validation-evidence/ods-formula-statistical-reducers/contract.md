# ODF 1.4 statistical-reducer evaluator contract

This contract defines the bounded implementation profile for the nine core
statistical reducers in OpenFormula 1.4 Part 4: `COUNT`, `COUNTA`,
`COUNTBLANK`, `AVERAGE`, `AVERAGEA`, `MIN`, `MAX`, `MINA`, and `MAXA`. It is
the semantic boundary for the resolver-free scalar evaluator, the
resolver-backed value evaluator, and their independent validation. It does
not authorize formula-cell recalculation, cached-result use, source refresh,
or publication of a changed cell.

The normative source is the repository-local ODF distribution:

| Source | SHA-256 |
| --- | --- |
| archive `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| member `part4-formula/OpenDocument-v1.4-os-part4-formula.html` | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |

The directly relevant sections are §§3.2.3, 3.3, 3.6–3.7, 4.6–4.11.13,
6.1–6.3, §6.13.6–§6.13.8, and §6.18.1–§6.18.4, §6.18.45–§6.18.46, and
§6.18.48–§6.18.49. The repository profile also follows the existing numeric
aggregate contract for finite binary64 values, exact sum state, cancellation,
and resource accounting.

## Signatures and selected empty behavior

The signatures retain the ODF pseudotypes. `+` means one or more supplied
arguments in the common function template; a semicolon separates formula
arguments. `#VALUE!` is `ScalarError::Value`, `#DIV/0!` is
`ScalarError::DivisionByZero`, and `#NUM!` is `ScalarError::Number`.

| Function | Part 4 signature | Included values | Empty or zero-argument result in this profile |
| --- | --- | --- | --- |
| `COUNT` | `COUNT({NumberSequenceList N}+)` | Numbers only | `COUNT()` is `0`; an admitted sequence with no numbers is `0` |
| `COUNTA` | `COUNTA({Any AnyValue}+)` | Every non-empty value, including errors and empty text | `COUNTA()` is `0`; an admitted sequence with no non-empty values is `0` |
| `COUNTBLANK` | `COUNTBLANK(ReferenceList R)` | Empty cells and selected empty text values | Missing/constant/array range is `#VALUE!`; one or more references are required |
| `AVERAGE` | `AVERAGE({NumberSequence N}+)` | Numbers in the sequence | `AVERAGE()` and an admitted sequence with no numbers are `#DIV/0!` |
| `AVERAGEA` | `AVERAGEA({Any N}+)` | Numbers, Text as `0`, Logical as `0`/`1`; empty cells omitted | `AVERAGEA()` and no included values are `#DIV/0!` |
| `MIN` | `MIN({NumberSequenceList N}+)` | Numbers in the sequence | `MIN()` and no numbers return `+0` |
| `MAX` | `MAX({NumberSequenceList N}+)` | Numbers in the sequence | `MAX()` and no numbers return `+0` |
| `MINA` | `MINA({Any N}+)` | Numbers, Text as `0`, Logical as `0`/`1`; empty cells omitted | `MINA()` and no included values return `+0` |
| `MAXA` | `MAXA({Any N}+)` | Numbers, Text as `0`, Logical as `0`/`1`; empty cells omitted | `MAXA()` and no included values return `+0` |

The zero-argument choices for `COUNT`, `COUNTA`, `MIN`, `MAX`, and `MINA`
are within the implementation-defined latitude in Part 4 and keep the
reducers composable. `MAXA` has no explicit zero-argument sentence but uses
the same additive empty-set profile as `MINA`. `AVERAGE` and `AVERAGEA`
follow their explicit no-value error constraints. Each supplied missing slot,
such as either slot in `COUNT(;)` or `AVERAGE(;)`, is converted at the
argument boundary to the same formula `#VALUE!` value in both evaluator
profiles; it is not an omitted zero-argument call. `COUNT` suppresses each
generated formula Error and contributes zero, while `COUNTA` counts each as
one non-blank Error value. The other reducers retain the first such Error and
return `#VALUE!`. `COUNTBLANK` has no scalar missing-slot form: a missing or
non-reference range argument is a `#VALUE!` shape error. Thus `COUNT()` and
`COUNTA()` remain zero-argument calls. The parser represents `COUNT(;)` as
two missing arguments, so it returns `0`, while `COUNTA(;)` counts both and
returns `2`. The scalar evaluator may represent each boundary value directly
as `WorkingValue::Error(Value)`; the value evaluator may carry
`RuntimeValue::Missing` while evaluating a slot, but it must normalize that
marker to the same formula Error before the reducer observes it so the two
profiles agree.

Each reducer publishes one scalar Number (or one formula/typed failure). The
reducer itself never materializes or broadcasts that scalar to an enclosing
matrix shape. If an enclosing array-producing operator, such as a projected
`IF`, uses the reducer as a branch, that enclosing operator may broadcast the
already-computed scalar under its own matrix rules. A direct formula Error
retains its source error according to the error rules below, except that
`COUNT` and `COUNTA` have the specific handling stated in their sections.

## Pseudotypes, scalar arguments, arrays, and references

`NumberSequence` and `NumberSequenceList` use the §6.3.7 and §6.3.8
conversions, with the repository's finite, locale-independent Number bridge:

* A scalar Number is one sequence member. A scalar Logical is converted to
  `0` or `1`. A scalar Text is parsed as a finite decimal Number; malformed,
  NaN, or infinite text is a generated `#VALUE!` conversion error. A scalar
  Empty value follows the existing value-bridge scalar policy and contributes
  numeric `0`; an explicit Missing value is `#VALUE!`.
* A Reference contributes only Number and formula Error cells. Empty, Text,
  and distinguished Logical cells are omitted. This is the ODF distinction
  between a sequence conversion of a reference and conversion of a scalar
  Text or Logical. The one exception is `COUNT`, which ignores Error cells as
  required by its “does not propagate Errors” rule.
* A `NumberSequenceList` accepts an ordered ReferenceList. References are
  visited in list occurrence order, a cuboid's planes in increasing sheet
  order, and each plane row-major. Repeated or overlapping references count
  once per occurrence. `COUNT`, `MIN`, and `MAX` therefore accept ReferenceList
  arguments.
* A `NumberSequence` accepts one logical Reference, including a 3-D cuboid,
  but does not accept a ReferenceList. `AVERAGE([.A1]~[.B1])` is `#VALUE!`;
  it is not flattened or implicitly intersected. A three-dimensional logical
  Reference remains one sequence and is valid for `AVERAGE`.

The value evaluator uses the established bounded repository extension for an
inline rectangular Array: elements are visited row-major and use the scalar
conversion profile for NumberSequence reducers. `Any` reducers likewise
visit inline array elements row-major. A ReferenceList cannot be converted to
an Array; passing one to a non-List pseudotype is `#VALUE!`. `COUNTBLANK`
accepts only ReferenceList/Reference values and rejects constants and Arrays.

`Any` performs no implicit Number conversion before the reducer. For
`COUNTA`, every value other than Empty is one non-blank value; Text (including
empty Text), Logical, Number, Complex, and Error each count once. For
`AVERAGEA`, `MINA`, and `MAXA`, Number values are included, Text values are
the numeric value `0`, Logical values are `0`/`1`, and Empty cells are
omitted. A Complex value is outside this finite real reducer profile and is a
generated `#VALUE!`; it is not silently projected to a real component.
Formula Errors remain formula Errors. A direct scalar Text is therefore
included as zero by the `A` functions even when its contents are not numeric;
the locale-independent Text-to-Number parser is used only by NumberSequence
functions such as `COUNT` and `AVERAGE`.

The resolver-free scalar API admits the scalar forms above. A Reference or
Array encountered there produces its existing typed `UnsupportedKind::Reference`
or `UnsupportedKind::Array` refusal. A scalar call to `COUNTBLANK` with a
constant range argument produces a formula `#VALUE!`; a resolver-backed call
with a structurally admitted reference is evaluated by the value VM.

## COUNT and COUNTA error behavior

`COUNT` is the deliberate exception to the common error-propagation rule.
Errors in a referenced sequence are ignored and do not increment the count.
An Error value produced by a direct scalar expression is also ignored after
eager evaluation; it is not converted to a Number and does not become the
result. A direct numeric Text that parses successfully contributes one. A
malformed direct Text reaches the profile's Number conversion as a formula
Error, but `COUNT` suppresses that conversion Error under its explicit
“does not propagate Errors” rule and contributes zero. This keeps the
function's direct NumberSequenceList conversion while preserving its special
error behavior; typed evaluator failures remain hard failures.

`COUNTA` counts an Error as one non-blank value, both when it is a direct
formula result and when it is read from a reference. Empty Text is content and
therefore counts. Empty cells alone do not count. This function never turns an
Error into the result merely because it counted it.

## COUNTBLANK profile

`COUNTBLANK` traverses the supplied ordered ReferenceList without materializing
its cells. A cell is blank when it is physically Empty. This profile also
counts a cell returning the empty Text value `""` as blank, selecting the
implementation-defined choice permitted by §6.13.8 and matching the common
spreadsheet host behavior. `COUNTBLANK` and `ISBLANK` therefore remain
distinct: an empty-text formula can be counted by `COUNTBLANK` while it is not
an Empty cell for `ISBLANK`.

Numeric zero, non-empty Text, Logical values, formula Errors, and Complex
values are non-blank. Unsupported cells and provider failures remain typed
evaluator failures. Duplicate ReferenceList entries are counted per
occurrence. A 3-D logical Reference is traversed sheet-by-sheet in the same
order as other ReferenceList functions.

## AVERAGE, AVERAGEA, MIN, MAX, MINA, and MAXA

`AVERAGE` computes the exact numeric sum of admitted Number values divided by
the checked count of those values. It ignores referenced Empty, Text, and
distinguished Logical cells; direct scalar Text and Logical values use the
NumberSequence conversion described above. An empty selected sequence returns
`#DIV/0!`. `AVERAGEA` performs the same exact reduction after including Text as
zero and Logical as `0`/`1`, while omitting Empty cells. It returns `#DIV/0!`
when no value is included. Formula Errors in either reducer are retained and
returned after the conceptual scan, before any generated empty-set or numeric
error.

`MIN` and `MAX` reduce admitted Number values using exact binary64 ordering.
They ignore non-number cells omitted by NumberSequence conversion. `MINA` and
`MAXA` include the `A`-family Text/Logical conversions and omit Empty cells.
For each extrema reducer, no included value produces numeric `+0`, and a
selected zero result is canonicalized to `+0`: `-0` and `+0` compare equal,
and the result does not depend on ReferenceList order. Formula Errors are
propagated for all four extrema functions because no function text grants the
`COUNT` exception. Text conversion failures can only arise from direct
NumberSequence arguments; `A`-family Text is always zero.

The extrema comparison is over finite binary64 values. Non-finite provider
Numbers are converted by the existing value bridge to `#NUM!` formula Errors.
No epsilon, locale, or host ordering policy is introduced.

## Numeric and signed-zero profile

The exact sum state from the numeric aggregate family is reused for
`AVERAGE` and `AVERAGEA`. The implementation retains represented binary64
operands as an exact sum, divides that exact sum by the checked integer count,
and performs one final nearest-even binary64 conversion of the exact quotient
under the finite Number profile. It must preserve cancellation,
minimum-subnormal averages, and the sign of an underflowed negative average
where the final IEEE-754 result is `-0`. A final non-finite result is
`#NUM!`; intermediate exact-sum growth is not by itself a failure.

`COUNT`, `COUNTA`, and `COUNTBLANK` use checked integer counts and publish a
finite Number. Extrema use comparisons without arithmetic accumulation and
canonicalize zero only at result publication. The result state is fixed-size;
no reducer may collect every selected value merely to compute an average or
extreme.

## Formula-error precedence and typed failures

Function arguments are evaluated eagerly in source order. The conceptual
traversal for a reference reducer is argument order, ReferenceList occurrence
order, sheet-plane order, and row-major cell order. The implementation may
read in another order only if it preserves the same first formula Error.

* For `AVERAGE`, `AVERAGEA`, `MIN`, `MAX`, `MINA`, and `MAXA`, the first
  formula Error in that conceptual order wins over a later conversion,
  empty-set, extrema, or numeric error. The scan may continue after recording
  the first formula Error so that typed source/budget fences are still honored.
* `COUNT` ignores all formula Errors encountered in its sequence. `COUNTA`
  counts them as non-blank values. `COUNTBLANK` counts them as non-blank
  cells. These are explicit function rules, not a general relaxation of
  formula error propagation.
* If no input Formula Error exists, the first generated formula error in
  conceptual order is returned. AVERAGE/AVERAGEA empty-set `#DIV/0!`, a
  malformed scalar NumberSequence Text `#VALUE!`, and a rejected shape or
  pseudotype `#VALUE!` follow this rule.
* `Unsupported`, cancellation, source-version changes, resource limits, and
  allocation failures remain typed `EvaluationFailure` values. They are not
  converted into formula Errors and are never caught by `IFERROR` or `IFNA`.

## Matrix projection and cache invariants

These reducers consume complete sequence/reference arguments even when nested
inside a projected matrix branch. A range argument must enter matrix/reference
context so a multi-cell Reference is not implicitly intersected at the current
output position. The scalar result is cacheable only when every direct scalar
descendant is coordinate-independent under the enclosing projection. A
computed scalar argument that contains a projected multi-cell reference must
remain position-dependent and must not reuse the first output cell's result.
Sequence reducers over fixed references, literals, and invariant nested
reducers may use the existing demand cache. Cache classification is bounded,
iterative, and budgeted; it must not recurse through an unbounded formula AST.

Three-dimensional References are not broadcast or collapsed: all admitted
planes are reduced in sheet order. `AVERAGE` accepts one logical 3-D Reference;
`COUNT`, `MIN`, `MAX`, `COUNTA`, `AVERAGEA`, `MINA`, `MAXA`, and `COUNTBLANK`
also preserve ordered ReferenceList occurrences as their signatures allow.

## Resource, memory, and resolver boundaries

The value evaluator must stream admitted references and lists cell-by-cell,
retaining only fixed reducer state, checked counts, and bounded reference
metadata. It must:

* validate reference area/cell counts and checked row/column arithmetic before
  scanning;
* charge AST work, argument conversion, every physically inspected cell,
  retained Text bytes, and reducer operations to the caller's hierarchical
  budget;
* check cancellation during long scans and source identity before and after
  evaluation;
* use fallible reservations for any retained metadata and never build a
  full cell vector for a scalar result; and
* publish no result when a typed failure or source-version fence occurs.

Text comparisons for `COUNTA`, `COUNTBLANK`, and `A` reducers borrow resolver
text; they must not clone a cell string per observation. Empty-text matching
is a length check. Formula cells are not recalculated and cached cell values
are not refreshed. The resolver is immutable and read-only for the duration
of evaluation.

## Required semantic and resource evidence

The focused tests and independent oracle must cover every reducer through:

* scalar Number, numeric and malformed Text, Logical, Empty, Missing, and
  formula Error values;
* inline arrays, single-cell and multi-cell References, 3-D References,
  ordered ReferenceLists, duplicate/overlapping occurrences, and the
  NumberSequence versus Any distinction;
* mixed Number/Text/Logical/Empty/empty-Text/Error cells, COUNT's ignored
  errors, COUNTA's counted errors, COUNTBLANK's selected empty-text policy,
  average empty-set errors, and extrema empty-set/signed-zero behavior;
* exact cancellation, minimum-subnormal and negative-underflow averages,
  finite-limit values, and one-final-rounding comparisons;
* projected nested reducers that verify cache invariance and no accidental
  broadcast or implicit intersection; and
* work, storage, cancellation, unsupported-cell, and source-version refusal
  paths with no partial result publication.

Native application caches are interoperability evidence only. They cannot
override this contract's explicit choices for empty text, zero arguments,
COUNT error handling, exact numeric reduction, or signed-zero extrema.
