# ODF 1.4 conditional aggregate evaluator contract

This contract scopes the six conditional aggregate functions in OpenFormula
1.4 Part 4: `SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTIFS`, `AVERAGEIF`, and
`AVERAGEIFS`. It is the semantic and resource boundary for the value
evaluator implementation and for its independent validation. These functions
are reference functions: a resolver-free scalar evaluator cannot execute them
because their range parameters are required to be references. The value
evaluator may execute them against its immutable, read-only resolver.

The normative source is the local ODF distribution:

| Source | SHA-256 |
| --- | --- |
| archive `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| member `part4-formula/OpenDocument-v1.4-os-part4-formula.html` | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |

The directly relevant sections are §§4.7–4.9 and 4.11.8 (Empty Cell,
Reference, ReferenceList, and Criterion), §3.4 (host-defined matching
properties), §§3.6–3.7 (numerical model and limits), §§6.1–6.3 (common
function, conversion, and error rules), §6.13.9–§6.13.10, §6.16.62–§6.16.63,
and §6.18.5–§6.18.6. Section §4.11.8 is especially important: a criterion
reference to an empty cell is numeric zero, whereas an explicitly empty text
criterion has the separate empty-cell matching rules.

## Signatures and result operations

The signatures below are normative. The optional pairs in the `IFS` forms are
zero or more additional `(Reference; Criterion)` pairs, but at least one pair
is required after the initial range.

| Function | Part 4 syntax | Operation | Required result rule |
| --- | --- | --- | --- |
| `SUMIF` | `SUMIF(ReferenceList\|Reference R; Criterion C [; Reference S])` | Select positions in `R` matching `C`, then sum Numbers from `R` or the generated `S` range. | Empty selection is numeric `0`; non-Number selected values are omitted from the sum. |
| `SUMIFS` | `SUMIFS(Reference R; Reference R1; Criterion C1 [; Reference R2; Criterion C2]...)` | Select positions satisfying every criterion and sum the corresponding values in `R`. | Empty selection is numeric `0`; non-Number selected values are omitted. |
| `COUNTIF` | `COUNTIF(ReferenceList R; Criterion C)` | Count positions in `R` matching `C`. | Result is a Number count. |
| `COUNTIFS` | `COUNTIFS(Reference R1; Criterion C1 [; Reference R2; Criterion C2]...)` | Count positions satisfying every criterion. | Result is a Number count. |
| `AVERAGEIF` | `AVERAGEIF(Reference R; Criterion C [; Reference A])` | Select positions in `R`; average Numbers in `R` or generated `A`. | No selected Numbers is `#DIV/0!` (`ScalarError::DivisionByZero`) under this repository profile. |
| `AVERAGEIFS` | `AVERAGEIFS(Reference A; Reference R1; Criterion C1 [; Reference R2; Criterion C2]...)` | Select positions satisfying every criterion and average corresponding Numbers in `A`. | No selected Numbers is `#DIV/0!` (`ScalarError::DivisionByZero`) under this repository profile. |

The `AVERAGEIF` rule has two distinct empty cases. A range can contain no
positions matching the criterion, or can contain matching positions but no
Number values to average. Both produce the existing value-evaluator average
error, `#DIV/0!`; an Empty, Text, Logical, Complex, or omitted value does not
become a zero contribution. `SUMIF` and `SUMIFS` use the existing exact
numeric sum reducer, so the final Number is rounded once from the represented
binary64 operands. `AVERAGEIF` and `AVERAGEIFS` use the same exact sum state
and divide by the selected Number count before the one final binary64
conversion. A non-finite final conversion is `#NUM!` (`ScalarError::Number`).
The exact reducer's signed-zero and minimum-subnormal behavior is retained;
the implementation must not insert a platform-dependent `f64` running sum or
an underflowing intermediate average that changes those results.

Part 4 specifically says that `SUMIF` may omit generated destination
positions whose row or column is beyond sheet bounds. `AVERAGEIF` has the
same top-left-plus-geometry construction, but §6.18.5 does not grant that
clipping permission. This profile therefore returns formula `#REF!`
(`ScalarError::Reference`) when any generated `AVERAGEIF` destination row,
column, or sheet plane is outside the provider's bounds. The `IFS` forms have
no generated destination geometry: all supplied ranges must have the same
finite shape.

## Reference admission and geometry

The range parameters are pseudotypes, not ordinary scalar or array
arguments. A constant Number, Text, Logical, Empty value, or inline Array in
a range slot is an invalid constant range and produces formula `#VALUE!`
after supplied arguments have been evaluated. It must not be treated as a
one-cell range or broadcast. A range expression that is a structurally
recognized but unsupported reference representation (for example an
external/subtable reference that this resolver cannot represent) remains a
typed `Unsupported(Reference)` capability refusal. A provider refusal while
reading a permitted range remains its typed evaluator failure.

The resolver-backed profile admits the following reference forms:

* A single logical `Reference` is admitted when its retained area planes are
  finite and within the provider's declared extents. A three-dimensional
  reference is one logical reference, not a `ReferenceList`; its planes are
  traversed in increasing sheet order, then row-major within each plane.
  Pairwise forms require the same number of planes and the same rows and
  columns per plane. This preserves the reference's sheet dimension instead
  of silently intersecting it to the caller's sheet.
* `SUMIF` and `COUNTIF` admit an ordered `ReferenceList` for `R`. Every
  logical reference in the list is traversed in list occurrence order, with
  its planes in sheet order and cells row-major. Overlapping or duplicate
  references are counted once per occurrence, as required by a sequence-like
  ReferenceList; they are not deduplicated.
* If `SUMIF` has an optional `S`, `R` may be a single `Reference` or a
  `ReferenceList` containing exactly one logical reference. A list containing
  more than one logical reference produces formula `#VALUE!`, exactly as
  §6.16.62 requires. `S` itself is one logical `Reference`, never a
  `ReferenceList`. Its actual row and column extent is ignored after its
  top-left anchor is obtained.
* `SUMIFS`, `COUNTIFS`, and `AVERAGEIFS` require one logical `Reference` for
  every range parameter. A ReferenceList is a pseudotype mismatch and
  produces formula `#VALUE!`; it is not flattened, implicitly intersected, or
  accepted merely because it contains one area. A 3-D logical Reference is
  accepted only when every paired range has the same plane count and shape.
* `AVERAGEIF` accepts one logical `Reference` for `R`, and its optional `A`
  is one logical `Reference`. A ReferenceList in either slot is formula
  `#VALUE!`. The same single-reference restriction applies to the criterion
  range of `SUMIF` when a destination range is supplied, as described above.

The implementation may use a private bounded range view rather than
materializing all cells. It must retain enough sheet, row, column, and plane
geometry to read corresponding positions and to construct destination
coordinates without using the destination's actual shape.

### Anchor destinations

For `SUMIF` with `S`, and `AVERAGEIF` with `A`, the destination is constructed
from the top-left cell of the supplied reference and the complete geometry of
the criterion range `R`:

* The destination's actual number of rows and columns are ignored. A one-cell
  anchor such as `C1` is sufficient and is expanded to the rows and columns
  of `R`.
* For a single-sheet `R`, destination position `(row, column)` is the
  top-left destination coordinate plus the corresponding zero-based offset.
  The destination's source sheet is the sheet containing its top-left cell.
* For 3-D `R`, corresponding destination planes start at the top-left plane
  of the single logical `S`/`A` reference and follow the sheet order of `R`.
  The actual plane span and row/column extent of `S`/`A` are ignored just as
  their actual two-dimensional extent is ignored; a one-plane reference is a
  valid anchor for a multi-plane `R`. This avoids imposing an invented
  destination-shape equality that §6.16.62 does not require. A generated
  plane outside the provider's sheet order is a reference-boundary failure
  (`#REF!`) for both functions. The explicit SUMIF clipping permission is
  limited to generated rows and columns.
* A generated `SUMIF` destination coordinate with a row or column outside
  the provider's sheet extent is silently omitted, following §6.16.62. Its
  in-bounds positions before and after an omitted edge remain in natural
  reference order. A generated plane outside the provider's sheet order is
  `#REF!`, because §6.16.62 names only row and column clipping. A malformed
  or unresolvable anchor reference is a formula reference/value error before
  scanning. `AVERAGEIF` returns `#REF!` for any generated out-of-bounds row,
  column, or plane and therefore must not publish a partial average.

The destination anchor is resolved and its top-left cell is validated before
the scan, but destination cells themselves are read only after their
corresponding criterion position matches. An Error, unsupported cell, or
provider failure in an unselected generated destination position is therefore
not observed.

## Criterion values and matching

Each `Criterion` argument is one scalar `Number`, `Logical`, or `Text`, or a
single-cell `Reference` that is read and converted to that scalar value. An
Empty reference cell is converted to numeric `0`, as §4.11.8 requires. A
multi-cell reference, a 3-D reference containing more than one cell, an
inline Array, or a ReferenceList supplied as a criterion is a pseudotype
mismatch and produces formula `#VALUE!`; no implicit intersection of a
multicell criterion range is performed. A syntactically empty required
criterion (`; ;`) is `#VALUE!`, distinct from a reference to an Empty cell.
Formula Errors in a criterion reference propagate as that formula Error.
Complex criteria are `#VALUE!` under this finite real profile.

The criterion expression is evaluated in scalar operator context so a scalar
expression such as `">" & [.T2]` can project its one-cell reference and produce
one Text criterion. This scalar expression coercion does not authorize implicit
intersection of a direct multicell criterion reference or array; those values
remain pseudotype mismatches as stated above. The resulting runtime value must
still be a scalar Number, Logical, or Text (or a single-cell reference read as
one of those values). A `RuntimeValue::Array`, including a one-cell computed
array, remains a formula `#VALUE!` criterion mismatch; §6.2's
implementation-defined pseudotype latitude is not used to widen this profile.

Number and Logical criteria without a text operator use typed equality:

* A Number criterion matches a Number cell with equal binary64 value. It does
  not coerce Empty, Logical, or Text cells to a Number for equality.
* A Logical criterion matches a Logical cell with equal Boolean value. It
  does not coerce Number `0`/`1` to Logical for equality.

Text criteria are parsed for one leading operator, in this order: `>=`, `<=`,
`<>`, `>`, `<`, and `=`. With no leading operator, equality is used. The
remaining text is kept byte-for-byte, including whitespace. Only an
operator-prefixed, non-empty remaining string that parses as a finite Number
uses numeric comparator semantics. A bare Text criterion remains a Text
criterion even when it looks like a number; therefore bare `"3"` matches a
Text cell containing `"3"`, not a Number cell containing `3`. This follows
§4.11.8's separate “value beginning with a comparator or operator” and “Other
Text value” cases. In particular, `"=0"` does not match an Empty cell. A
malformed or non-finite numeric-looking string remains a text criterion and is
not a conversion failure.

An empty remaining text value has the special §4.11.8 behavior:

* `"="` and an empty text criterion match only Empty cells;
* `"<>"` matches non-Empty cells;
* `<`, `<=`, `>`, and `>=` with an empty right side match no candidate.

For a non-empty text criterion, this implementation profile uses literal,
case-sensitive, whole-cell text comparison. Ordered text comparisons use
Unicode scalar lexicographic ordering. A Text candidate is compared as Text;
Number, Logical, Empty, and Complex candidates do not equal the text, and
match a `<>` text criterion as unequal. A candidate Formula Error is handled
by the error policy below. This profile records the host choices as
`HOST-CASE-SENSITIVE=true`,
`HOST-SEARCH-CRITERIA-MUST-APPLY-TO-WHOLE-CELL=true`,
`HOST-USE-REGULAR-EXPRESSIONS=false`, and `HOST-USE-WILDCARDS=false`.
The matcher must not accidentally use the database module's empty-criterion
conversion for an explicitly empty Text criterion: only a reference to an
Empty cell converts to Number zero.

The three host text properties are normative variability points in ODF, so a
future host-configured matcher may implement regex, wildcard, case-folded, or
substring behavior. This batch does not add a host-property object or a
regex dependency to the public API; tests and native comparisons must be
labelled as evidence for the fixed profile above. Wildcard/regex syntax must
not be silently enabled by reusing a different host profile.

## Multi-criterion selection and lazy reads

`SUMIFS`, `COUNTIFS`, and `AVERAGEIFS` use the same zero-based position in all
criteria ranges. All range shapes, including plane count for 3-D references,
are checked before any corresponding cell scan. A logical AND is applied in
criterion-pair order. The implementation may stop evaluating a position at
the first false criterion, so later criteria cells and the selected value are
not read for that position. A true result causes exactly one destination or
count observation for that position.

`SUMIF`, `COUNTIF`, and `AVERAGEIF` first evaluate the criterion range `R`.
For each matching position, `SUMIF`/`AVERAGEIF` read the generated `S`/`A`
destination if present, or use the matching `R` cell otherwise. A nonmatching
position never reads its optional destination. For `SUMIF` without `S`, the
matching Number is the Number in `R`; for `AVERAGEIF` without `A`, the
matching Number is likewise the Number in `R`.

The required destination value is evaluated lazily after criteria selection,
including for `AVERAGEIFS` where §6.18.6 explicitly says the `A` cell is
evaluated only when every criterion matches. This is observable with provider
Unsupported values, formula Errors, and read counters and is part of the
resource contract.

## Formula errors and typed failures

Function arguments are eagerly evaluated in source order. A direct formula
Error already present in an argument is retained with the leftmost-error
policy before generated arity, constant-range, or shape errors. An invalid
constant range is a generated formula `#VALUE!`, while an invalid average
result is generated `#DIV/0!`; these values remain catchable by `IFERROR` and
`IFNA` according to their existing rules.

Criterion argument errors have the following explicit behavior:

1. A formula Error in a direct criterion scalar or in its single-cell
   criterion reference is returned as that Error before range scanning.
2. A formula Error read from a criterion range while deciding a position is
   returned as that Error. It is not converted to Empty, Number zero, or a
   nonmatching value. If an earlier criterion in an `IFS` position is false,
   a later criterion cell is not read and cannot produce an error.
3. An Error in a selected sum/average destination cell is returned only when
   that destination position is actually selected. Errors in nonmatching
   destinations remain unobserved. `COUNTIF` and `COUNTIFS` have no separate
   destination range, so an Error in their admitted criteria range is covered
   by rule 2.
4. A provider `Unsupported`, cancellation, source-version change, resource
   limit, or allocation failure remains a typed `EvaluationFailure`. It is not
   converted to a formula Error and cannot be caught by `IFERROR`/`IFNA`.

The scan continues only as far as needed by the selected function and the
profile's error-order guarantee. Error ordering is deliberately sequential:

1. Direct formula Errors already present in supplied arguments win in source
   argument order.
2. Range geometry and arity errors are generated before the cell scan. A
   single-cell criterion reference is prepared in its source argument order;
   an Error read there wins before scanning range positions.
3. Once scanning begins, positions are visited in ReferenceList occurrence,
   sheet-plane, row, and column order. Within one position, criterion pairs
   are inspected left to right, with the selected destination read last.
   Therefore a selected destination Error at an earlier position wins over a
   criterion-range Error at a later position, and a false earlier criterion
   prevents later criterion cells at that same position from being read.

The implementation may short-circuit a false criterion position, but it must
not let a later physically read error hide an earlier one. It must not
pre-read all optional destination cells merely to search for errors, because
that would violate the lazy selected-value rule. A formula Error observed in a
cell is retained while the admitted scan continues, matching the existing
database evaluator. If a subsequent provider operation returns a typed
`EvaluationFailure` (including Unsupported, cancellation, source-version,
resource, or allocation failure), that typed failure propagates immediately
and supersedes the retained formula Error: the observation is incomplete and
must not be published as a catchable formula value. The first-error ordering
above applies among formula Errors successfully observed before any typed
failure.

The same precedence applies when the numeric reducer records a generated
error. The scan-observed Formula Error wins over a later final conversion
error such as `#NUM!` from a SUM or AVERAGE result; the implementation must not
replace an already observed cell error while publishing the reducer result.
Under the current finite range and fixed accumulator bounds, `push_number`
cannot fail before a later admitted cell is observed, so this rule concerns
final conversion (and remains the required policy if those bounds change).

`COUNTIF` and `COUNTIFS` count positions, including positions whose matching
criterion is a text/number inequality, rather than counting only Number
cells. A candidate Error cannot be classified under this profile and follows
rule 2. `SUM*` and `AVERAGE*` only contribute finite Number cells; Empty,
Text, Logical, and Complex destination cells are omitted. A selected Formula
Error follows rule 3 before an empty-selection result is published.

## Resolver-free scalar boundary

The public resolver-free `evaluate_scalar` path has no reference capability
and does not execute this family. It still preserves the established
argument/error boundary: a call such as `SUMIF(1;1)` or `AVERAGEIFS(1;1;1)`
with only constant range arguments is scheduled and finishes as formula
`#VALUE!`, while a reference argument reaches the typed
`Unsupported(Reference)` refusal before the conditional function can run.
The value evaluator handles constant range arguments because it can
distinguish their scalar/array runtime types and returns formula `#VALUE!`
for the invalid pseudotype. This distinction is intentional: a typed
capability refusal means the API cannot retain or read the reference, whereas
formula `#VALUE!` means a resolver-backed call received a non-reference
value.

When a conditional aggregate appears in a lazy branch, the value evaluator
must preserve existing branch laziness. An unselected conditional aggregate
does not inspect any of its references. A selected aggregate is coordinate
independent and may be cached by the existing projection cache, but the
cache key must include the complete expression and demand shape; no result
may be shared across a changed resolver source version.

## Work, storage, and reference limits

The operation is scalar-result, read-only calculation. It must:

* charge function arguments, range geometry, criterion parsing, every
  physically inspected criterion cell, and every selected destination cell to
  the existing work budget;
* enforce the cumulative `max_reference_cells` limit over all range reads and
  never reset that count for each criterion or pair;
* use checked arithmetic for plane counts, rows, columns, offsets, generated
  destination coordinates, and list sizes;
* retain only bounded range metadata, compiled criterion matchers, exact sum
  state, and at most one row's scratch values; it must not materialize an
  unbounded copy of every range; and
* reserve all variable storage through the existing fallible storage budget
  before allocation, releasing retained text, matcher, range, and scratch
  reservations in the correct drop order.

Shape validation may inspect geometry and provider extents without reading
cell values. Once validation succeeds, the read order described above is
observable and must remain stable for error precedence and provider-read
tests. A typed failure must publish no partial numeric result.

The immutable resolver is source-version fenced before and after evaluation.
Formula cells are not recalculated and cached formula values are not treated
as fresh values. External names, package database ranges, host data refresh,
macros, and volatile services remain outside this batch. No public resolver,
AST, package identifier, host handle, or dependency is added.

## Required validation

Focused tests and independent evidence must cover:

* all six functions in scalar-demand and matrix-demand value modes;
* one-cell, row, column, rectangular, 3-D, and ordered ReferenceList ranges;
* `SUMIF`/`COUNTIF` list order, duplicate/overlap occurrence behavior, and
  the `SUMIF` optional-destination rejection for a list with more than one
  reference;
* equal and unequal `IFS` shapes, plane counts, and constant/inline-array
  range rejection;
* single-cell criterion references, Empty criterion references, empty Text,
  Number, Logical, numeric comparator Text, malformed Text, every comparator,
  bare numeric-looking Text versus Text and Number candidates, explicit `=0`
  versus Empty, and whitespace preservation;
* scalar and multicell criterion arrays/references, direct formula Errors,
  criterion-range Errors, selected and unselected destination Errors, and the
  leftmost conceptual error order, including a typed provider failure after a
  stored formula Error;
* anchor destinations whose actual range is one cell but whose generated
  geometry spans the full criterion range, including generated positions
  beyond sheet bounds and lazy reads of those positions;
* no matches, matches with no Number destination values, exact cancellation,
  subnormal values, finite extrema, Number overflow, and average division by
  zero;
* resolver read counts, duplicate reference occurrence, cancellation,
  source-version changes, work/storage/reference-cell refusals, unsupported
  provider cells, lazy `IF` branches, projection-cache reuse, and release of
  all retained reservations; and
* resolver-free scalar calls returning typed `Unsupported(Function)` while
  resolver-backed calls with constant range arguments return formula
  `#VALUE!`.

Independent numerical goldens must be generated from exact `Fraction` values
of the represented binary64 inputs and independent integer criterion masks.
Native cached spreadsheet observations may corroborate ordinary vectors but
cannot define host wildcard, regex, locale, or empty-criterion behavior.

This contract follows the accepted boundaries in
`docs/adr/0001-priorities-and-api-layers.md`,
`docs/adr/0004-semantic-api-design.md`, and
`docs/adr/0005-io-memory-and-performance.md`: correctness-first typed API
layers, no hidden recalculation, finite and fallible resource accounting, and
measured evidence that identifies the exact source and host profile.
