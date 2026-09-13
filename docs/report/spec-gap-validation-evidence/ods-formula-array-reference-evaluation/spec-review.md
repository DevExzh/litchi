# OpenFormula array and reference evaluation review

Status: normative/design review for the next evaluation substrate. No
implementation or gate result is accepted by this report yet.

## Primary source

The source is the checked-in ODF 1.4 specification archive
`3rdparty/specs/OpenDocument-v1.4-os.zip` (SHA-256
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`). The
normative formula entry is
`part4-formula/OpenDocument-v1.4-os-part4-formula.html` (SHA-256
`ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`). The
requirements below are from the anchors `a_3_2_2_Expression_Calculation`,
`a_3_2_3_Operator_and_Function_Evaluation`,
`a_3_3_Non-Scalar_Evaluation__aka_'Array_expressions'_`,
`a_4_7_Empty_Cell`, `a_4_8_Reference`, `a_4_9_ReferenceList`,
`a_4_10_Array`, `a_5_8_References`, `a_5_9_Reference_List`,
`a_5_10_Quoted_Label`, `a_5_11_Named_Expressions`,
`a_5_12_Constant_Errors`, `a_5_13_Inline_Arrays`,
`a_6_3_Implicit_Conversion_Operators`,
`Infix_Operator_Reference_Range`,
`Infix_Operator_Reference_Intersection`, and
`Infix_Operator_Reference_Concatenation`. The logical aggregate and iteration
rules are in §§6.15.2 (`AND`), 6.15.8 (`OR`), and 6.15.10 (`XOR`), while the
pseudotype boundary is in §§4.11.2, 4.11.12, 6.3.3, 6.3.4, 6.3.7, and 6.3.8.

The current audit explicitly leaves formula evaluation, array expressions,
external references and recalculation open in `docs/report/spec-gap-audit.md`
§10. This review narrows a possible local, inert evaluation capability; it
does not imply workbook recalculation or external-source support.

## Evaluation contexts and order

Part 4 §3.2.2 evaluates constants, operations, function calls, named
expressions, quoted labels, automatic intersections and arrays according to
their type. Operator precedence and parentheses govern operations. After
evaluation, a Reference used where one non-reference value is needed undergoes
the §3.3/§6.3 scalar conversion; an Array used for display follows the
non-scalar display rules.

Section 3.2.3 makes eager source-order argument evaluation the default. A
function-specific exception may be lazy. Formula Errors may short-circuit
unless the function suppresses them, and an incorrect type invokes its
specified implicit conversion. Therefore a reference or array resolver must
run inside the function's argument policy: eager functions visit every present
argument in source order, while `IF`/error-handler style lazy functions must
avoid resolving or iterating an unselected branch. A resolver failure remains
an evaluator failure; it must not be turned into a formula Error merely so an
error handler catches it.

There are two distinct modes:

* In ordinary scalar context, an inline array contributes element `(0,0)`. A
  single-cell reference is dereferenced. A multi-cell reference uses implied
  intersection with the formula cell's current row and column, and succeeds
  only when exactly one cell remains; otherwise it produces an Error.
* In matrix context, identified by a matrix cell area or a function's
  `ForceArray` behavior, a scalar-expected function or unary/binary operator is
  evaluated once per output position. Range and reference concatenation are
  reference operations and do not use scalar array lifting. Inline arrays and
  individual references are interchangeable inputs for this iteration;
  `ReferenceList` is the explicit exception described in §4.9.

Matrix results are rectangular, with height and width equal to the maximum
height and width of all non-scalar arguments. Broadcasting is exact:

| Argument shape | Value used at output position `(column,row)` |
| --- | --- |
| scalar or singleton | Repeat everywhere |
| one-column vector | Its value at `row`, repeated across columns |
| one-row vector | Its value at `column`, repeated across rows |
| one-column plus one-row vectors | Their row/column values form the cross product |
| two-dimensional matrix | Corresponding value, or `#N/A` when outside its shape |

Functions that return arrays are not implicitly iterated in matrix mode; their
`(0,0)` element is used. A display area then uses an existing result position,
repeats a one-column result across later columns, repeats a one-row result down
later rows, or displays `#N/A` when none of those rules applies.

The specification's examples are useful conformance vectors:

```text
ABS({-3;-4})                         -> ABS(-3)       (scalar row vector)
ABS({-3|-4})                         -> ABS(-3)       (scalar column vector)
{1;2;3|4;5;6}                        -> 1             (simple display)
1 + {1;2;3|4;5;6}                   -> {2;3;4|5;6;7}
{1|2} + {10;20|30;40}               -> {11;21|32;42}
{1;2} + {10|20}                     -> {11;12|21;22}
{1;2} + {3;4;5}                     -> {4;6;#N/A}
{1} + {1;2}                         -> {2;3}
MID("abcd";{1;2};{1;2;3})          -> {"a";"bc";#N/A}
```

For a resolver-backed formula in cell B2, `[.A1:.C1]` implicitly intersects
to B1, while the same row reference in D4 has no intersection and returns
`#N/A` or a more specific Error. `[.A1:.A3]` in B2 intersects to A2 and in D4
has no intersection. The general §6.3.3 rule computes the union of the current
row and current column, intersects it with the reference, and requires exactly
one remaining cell.

There is a real edge overlap between that general wording and the more
specific vector wording in §3.3. For a row vector, §3.3 selects the cell at
the evaluation position's column and the vector's row; for a column vector it
selects the cell at the evaluation position's row and the vector's column.
Thus, for a formula in `B1`, `[.A1:.C1]` selects `B1`, even though a literal
§6.3.3 union of the whole current row with the current column would contain
all three cells and fail the exact-one test. The dual case is a formula in
`A1` with `[.A1:.A3]`, which selects `A1`. The shape-specific §3.3 rule should
therefore be dispatched before the generic §6.3.3 rule for one-row or
one-column references; this is a specific-over-general profile decision made
necessary by the overlapping text. In either case the evaluation coordinate
must lie inside the vector's span: `D1` with `[.A1:.C1]` and `A4` with
`[.A1:.A3]` still produce `#N/A`. A 1x1 reference is read directly, while a
two-dimensional reference continues to use §6.3.3's union and exact-one
test. Matrix and `ForceArray` contexts bypass both implicit-intersection
paths.

The vector examples in §3.3 also establish the orientation used here:
`{1;2}` is a one-row (horizontal) vector and `{1|2}` is a one-column
(vertical) vector. Do not infer the opposite orientation from the `Nx1`
notation in the prose.

### Cross-sheet intersection and final projection

An explicit sheet locator supplies the sheet of the selected cell; the
evaluation position supplies the row or column coordinate used by the
shape-specific vector rule. Therefore a formula in `Main.B2` with
`[Data.A1:.A3]` has a column vector on `Data` and selects `Data.A2` under
§3.3. The dual example, `Main.B2` with `[Data.A1:.C1]`, selects `Data.B1`.
The standard gives no same-sheet restriction in either vector clause. This is
the cross-sheet reading of the explicit vector rule and should not be replaced
by a literal set intersection that first discards the other sheet.

A multi-sheet cuboid has no first-sheet fallback in scalar conversion. For
`[Sheet1.A1:Sheet3.A3]`, an ordinary scalar conversion applies the current
formula row and column union from §6.3.3 and intersects that set with the
cuboid. Under the sheet-local cell model, only cells on the formula's current
sheet can be candidates: exactly one candidate returns its value; no candidate
returns `#N/A` or a more specific Error; and multiple candidates return an
Error. For example, `Sheet2.A1` with `[Sheet1.A1:Sheet3.A1]` can select
`Sheet2.A1`, while a formula on a sheet outside the cuboid cannot select from
its first sheet. A formula at `Sheet2.B2` with the 3-by-3 cuboid has multiple
current-row/current-column candidates and therefore fails the exact-one test.

`ForceArray` and matrix evaluation do not define a generic 3-D-reference to
2-D-Array projection. Section 6.3.4 permits iterating a multi-cell reference,
but §4.11.12's sheet/row/column order is a sequence rule, not an Array shape
rule. A bounded implementation should preserve a cuboid as a Reference and
return a typed unsupported result (or a documented, function-specific
per-sheet projection) rather than silently selecting the first sheet or
flattening sheets into rows.

Finally, §3.2.2 returns a bare Reference as a Reference. Implicit intersection
is applied when an operator/function needs one non-reference value, or when an
explicit display/scalar projection is requested; it is not an automatic
mutation of the expression's natural result. Keep `evaluate -> Reference`
separate from `project_scalar(position) -> cell value`, and apply the display
area rules to an Array only at the display boundary. A reference-producing
`:`/`!`/`~` expression must remain reference-valued until such a consumer asks
for a scalar.

### Logical sequence versus matrix semantics

The logical family has an explicit non-scalar exception. `AND` (§6.15.2) and
`OR` (§6.15.8) both have the signature
`Logical|NumberSequenceList` and explicitly say that, in array context, all
arguments are aggregated rather than evaluated as a matrix and no array is
returned. A range or `ReferenceList` passed to either function is therefore
flattened through `NumberSequenceList` in the §6.3.7/§6.3.8 order. It is not
broadcast element-by-element. The same explicit exception applies to an
inline Array in array context: it contributes to the aggregate rather than
producing a matrix result. The specification gives no separate general
Array-to-`NumberSequenceList` conversion, so this array handling must remain
local to the AND/OR exception. For a scalar Text argument, the
`Number` conversion permitted by `NumberSequenceList` still applies; Text
cells encountered inside a referenced range are omitted, rather than
converted. Number cells are included, distinguished Logical cells are
omitted, Empty cells are omitted, and Error cells remain in the sequence and
propagate under the ordinary eager error rules. For references and
ReferenceLists, §4.11.12 requires reference-list and sheet order; choose one
of its permitted row-at-a-time or column-at-a-time orders consistently within
each sheet.

The words “in array context” are material: an inline Array in ordinary scalar
context still contributes element `(0,0)` through §3.3. Only the matrix/array
context invokes the AND/OR aggregate exception, so a profile should test both
contexts instead of applying aggregation to every call that happens to contain
an Array.

The element typing of that Array aggregate is deliberately unresolved by the
standard. Section 6.3.7/6.3.8 enumerates scalar Number/Text/Logical and
Reference/ReferenceList conversions, but does not define an
Array-to-`NumberSequenceList` conversion. Section 6.15.2/.8 only requires the
array-context aggregate and the absence of an array result. Consequently the
specification does not establish whether `AND({TRUE();FALSE()})` includes both
distinguished Logical elements, filters them as if they came from a reference,
or reaches the zero-element result after filtering. The first of those choices
is a reasonable format-owned profile: flatten Array elements in the profile's
documented order, apply the Logical conversion to each element (including
distinguished Logical values), and propagate Errors. It remains a profile
choice; its result must not be presented as an ODF-mandated or Excel-derived
array rule. Empty/Text Array elements also require an explicit profile because
their direct sequence conversion is absent.

The standard does not state what `AND` or `OR` returns when a present
`NumberSequenceList` argument becomes a zero-element sequence after that
filtering. This is distinct from `AND()`/`OR()` with zero syntactic
parameters, for which §6.15 permits either a Logical value or an Error. A
bounded implementation should document one profile choice: the logical
identities `AND(empty sequence) = TRUE` and `OR(empty sequence) = FALSE` are
the natural all/any identities, but they are an inference rather than a
mandated ODF result. Returning a formula Error is another explicit profile
choice. In no profile should omitted Empty/Text cells be silently replaced by
zero or FALSE sequence elements.

`XOR` (§6.15.10) has the different signature `Logical` and has no aggregate
exception. Its one-or-more parameters are combined by parity, and a
non-scalar Reference or Array passed in matrix context follows the normal
§3.3 implicit iteration and broadcasting rules. In scalar context it uses the
`(0,0)`/implied-intersection result. A `ReferenceList` is not implicitly
flattened for `XOR`; §4.9 makes passing one where a scalar is expected an
Error. `NOT` has the same no-aggregate behavior. This distinction must remain
visible in the API so a shared logical dispatcher does not accidentally apply
AND/OR aggregation to XOR or NOT.

## Values, conversion, and absence

The value model must keep these states distinct:

| State | Required distinction or conversion |
| --- | --- |
| Number, Logical, Text | Scalar values. Logical may be a distinct type or Number 0/1. Text length zero is the empty string. |
| Empty cell | Neither zero nor empty string and distinct from every Error, especially `#N/A`. Conversion is contextual: Number uses 0, Logical uses FALSE, Text uses an empty string; Number/Logical sequences omit empty cells. |
| Missing parameter | A syntactically empty function parameter, distinct from an Empty cell, zero, and empty Text. An empty parameter list is zero parameters; an empty parameter is legal syntax but a function need not accept it. |
| Error | A formula value that normally propagates. `#N/A` is required; `#REF!`, `#VALUE!`, `#NUM!`, `#DIV/0!` and others are conventional supported names. Serialized `office:string-value` does not make an Error into Text. |
| Reference | A cell, rectangle, cuboid or whole-axis descriptor. It is not materialized as a scalar until conversion or iteration. |
| ReferenceList | One or more ordered references. It cannot be converted to an Array; a function not accepting it must return an Error. Duplicates remain and count as separate areas. |
| Array | A rectangular value grid in supported evaluation, with at most one value per position; a position may contain no value. Array elements can themselves be Empty or Error when derived from cells. |

Section 6.3 requires the following conversions:

* `Scalar` returns Number, Logical or Text unchanged; a single-cell reference
  is read; a multi-cell reference uses implied intersection. An Array is
  reduced to `(0,0)` by the §3.3 scalar rule rather than silently flattened.
* `Number` preserves Number, maps Logical to 0/1, and leaves Text conversion
  implementation-defined and locale-sensitive. A reference is intersected,
  then an empty cell is 0. A format-owned profile must state its text/locale
  policy rather than claim universal ODF behavior.
* `Integer` first uses Number, then applies the function's rounding rule; if
  none is specified, rounding is implementation-defined.
* `NumberSequence` creates a one-element sequence from Number, Text or Logical
  via Number. A Reference contributes only Number and Error cells; Empty and
  Text cells are omitted. `NumberSequenceList` adds ReferenceList and processes
  each reference in occurrence order.
* `Logical` maps Number zero/nonzero to FALSE/TRUE, preserves Logical, and has
  implementation-defined Text conversion. A reference is scalar-converted and
  an empty cell is FALSE. `LogicalSequence` omits Empty cells and includes only
  Logical/Error cells (or Numbers too when Logical is not distinct).
* `Text` renders Number without whitespace, preserves Text, renders Logical as
  `TRUE` or `FALSE`, and converts an empty reference to empty Text.

Do not collapse `Empty`, `Missing`, empty Text, `#N/A`, and an unsupported
resolver result into one sentinel. For example, an empty referenced cell can
become 0 for Number or empty Text for Text, while `{\"\"}` is already a Text
array element and an omitted argument remains Missing. Division by an empty
cell is listed as `#DIV/0!` in §5.12's conventional error table.

## Inline arrays

The syntax is:

```text
Array     ::= '{' MatrixRow ( RowSeparator MatrixRow )* '}'
MatrixRow ::= Expression ( ';' Expression )*
RowSeparator ::= '|'
```

The grammar accepts one or more nonempty rows and can therefore parse a ragged
shape such as `{1;2|3}`. Section 4.10 defines Array as rows with equal column
counts, and §5.13 says an evaluator supporting inline arrays shall accept
nonempty rectangular matrices with constant values. This is a syntax-versus-
evaluation distinction: retain the parsed ragged structure for diagnostics,
then return a formula Error for a use that cannot support its nonrectangular or
nonconstant content. Never silently pad a ragged array or reinterpret it as a
vector. `{}` and a row containing an empty expression are syntax errors.

At minimum, a constant-array profile should accept constant Number and Text
elements and preserve constant Error elements when its Error model supports
them. Expressions involving references, names, labels or nonconstant functions
need an explicit profile decision; the §5.13 interoperability note warns that
anything beyond constant Number or String can impair interchange. A rejected
array capability must remain typed and must not be confused with a formula
element Error.

### Array/reference conversion boundary

The pseudotypes make the boundary one-way only where the expected operation
defines it. Section 4.11.2 says an Array with more than one element is not a
Scalar, just as a multi-cell reference is not a Scalar. Section 6.3.4's
`ForceArray` rule permits a multi-cell *reference* supplied to a scalar-valued
operation to be iterated into an array result; this is an evaluation-mode
operation, not a general conversion that gives an Array reference identity.
An inline Array, including a one-element Array, has no cell coordinates and
must not be fabricated into a `Reference` or used as an operand of `:`, `!`,
or `~`. Those operators require a `Reference` or `ReferenceList`; a mismatched
Array therefore produces the operation's formula Error.

Section 4.9 is explicit that a `ReferenceList` cannot be converted to an
Array, and that passing a ReferenceList where a function expects a scalar in
array iteration is an Error. A single `Reference` may be exposed as a lazy
matrix view when §3.3 or `ForceArray` requests iteration, without eagerly
materializing every cell. This does not extend to ReferenceList flattening:
only pseudotypes such as `NumberSequenceList` explicitly accept and stream a
ReferenceList. Conversely, a ReferenceList passed to `XOR`/`NOT` or another
scalar-only function must remain an Error at the type boundary. The reference
operators are the deliberate exception: `:`, `!`, and `~` explicitly accept
ReferenceLists and must preserve their list semantics. These rules avoid
treating an inline value grid as a range and preserve the distinction between
cell resolution and array evaluation.

## References and local resolution

Section 5.8 defines six range-address families:

```text
Reference ::= '[' (Source? RangeAddress) | ReferenceError ']'
RangeAddress ::= SheetLocatorOrEmpty '.' Column Row (':' '.' Column Row)?
              | SheetLocatorOrEmpty '.' Column ':' '.' Column
              | SheetLocatorOrEmpty '.' Row ':' '.' Row
              | SheetLocator '.' Column Row ':' SheetLocator '.' Column Row
              | SheetLocator '.' Column ':' SheetLocator '.' Column
              | SheetLocator '.' Row ':' SheetLocator '.' Row
```

`SheetName` is quoted or an unquoted component; a `SheetLocator` may carry an
ordered `.SubtableCell` chain. Columns are one or more uppercase A–Z letters;
rows are one-based decimal integers. `$` markers on sheet, column and row
components are lexical absolute/relative metadata and must survive any
translation or source-preserving operation. A sheet locator may contain a
quoted subtable name or a cell subtable selector.

No sheet locator means the current sheet at the formula's evaluation position.
In the first range form, a right endpoint beginning with `.` inherits the
first endpoint's explicit locator. Two explicit sheet locators form a cuboid
including sheets positioned between them. A whole-column or whole-row form
extends to the evaluator's supported sheet bounds; an explicit coordinate
beyond those capabilities is an Error, not a reference. Unquoted sheet names
may contain a colon under the grammar's character class, so a parser cannot
split a bracket body at the first colon without respecting quoted names,
subtables and endpoint forms.

`Source` is an RFC3987 IRI-reference prefix and its resolution is host-defined.
The core evaluator should accept an explicit caller-owned resolver for local
sheet identities and cell values. Source-qualified or otherwise external
references remain inert and return a typed unsupported-capability result unless
the caller supplies an explicit provider. There is no ambient network or
workbook access. A resolver should expose a lazy reference shape and bounded
cell access rather than allocate every cell of a whole row, column or cuboid.

The three reference operators have separate semantics:

| Operator | Result and behavior |
| --- | --- |
| `Left : Right` | Reference to the smallest inclusive 3-D cuboid containing both operands, including sheets between explicit sheet endpoints. Reference lists extend the range for every element. Each operand must be Reference/ReferenceList. |
| `Left ! Right` | Reference intersection. Non-reference operands are an Error. Reference lists use every left/right combination; empty intersections are omitted, and if all are empty an Error is returned. |
| `Left ~ Right` | Ordered ReferenceList concatenation, not set union. Left areas precede right areas, duplicates are retained, and `AREAS` is additive. Non-reference operands are an Error. |

Their precedence is `:` above `!` above `~`, with the Table 1 left
associativity. `&` is the scalar string concatenation operator and must not be
confused with `~`. Embedded bracket ranges such as `[.A1:.B2]` should be
preferred when serializing, while a general `:` operator must still preserve
the leftmost/reference-list rules.

## Labels, names, and errors at the boundary

Quoted labels are text found in table cells, but in formula position they are
looked up as row/column labels. Defined label ranges on the current sheet take
precedence; column-oriented ranges are searched before row-oriented ranges.
Automatic label lookup is controlled by the host property
`HOST-AUTOMATIC-FIND-LABELS`, has a specified distance/direction tie-break, and
is explicitly a portability hazard. A single label used where a scalar is
expected gets implicit intersection at the formula cell's row or column. As a
non-scalar argument, an automatic label becomes a contiguous range with the
specified one-empty-cell skip rule. `QuotedLabel !! QuotedLabel` must resolve
to exactly one cell or return an Error. A local reference substrate should
keep labels separate from ordinary Text and can leave automatic labels as a
typed unsupported host capability.

Global simple named expressions are required when names are supported. Names
match case-insensitively and cannot differ only by case. Sheet-local names are
optional; the most-specific name wins, and lookup walks the current or named
sheet's container chain. External named expressions use the same inert Source
boundary as references. A name resolver is therefore separate from a cell
reference resolver, and a missing/ambiguous name must not become an empty cell.

## API and resource requirements

An ergonomic format-owned evaluator can retain the existing immutable
`Expression` and add an explicit context containing:

* an evaluation position (sheet identity, zero-based row/column API
  coordinates, and optional matrix display area; reference row lexemes remain
  one-based at the syntax boundary);
* a caller-owned local `ReferenceResolver` capability that returns lazy cell,
  rectangle, cuboid or ordered-list views; and
* an explicit evaluation mode (`Scalar`/implicit-intersection or `Matrix`) and
  host policy for locale, text conversion, labels and names.

The public result should distinguish scalar values, bounded rectangular arrays,
formula Errors, and typed evaluator failures. Reference and array views should
carry dimensions without exposing package IDs or archive handles. External
sources, names and labels can be refused until their explicit providers exist.

Before matrix execution, compute checked output dimensions and a checked product
of rows × columns. Apply finite maxima for array cells, sheet rows/columns,
reference-list areas, resolver cell visits, work, storage and cancellation.
Whole-axis references must be intersected or streamed under those bounds rather
than expanded into an unbounded vector. Array elements, matrix results and
reference-list staging need fallible reservations; failed checks publish no
partial result. An iterative flat evaluator is preferable to recursive
per-element evaluation so a long array or broadcast chain cannot overflow the
call stack. Lazy branches must incur no resolver work for skipped elements.

This follows ADR 0001's typed unsupported/error boundary, ADR 0003's immutable
snapshot model, ADR 0004's data-bearing semantic values, ADR 0005's
caller-supplied budgets/cancellation/fallible storage, and ADR 0006's inert
external-content rule. Evaluation must not mutate cell caches, recalculate a
workbook, fetch a Source IRI, or publish package bytes.

## Minimum regression vectors

The implementation review should include at least these cases, with a resolver
position and explicit expected outcome:

1. `ABS({-3;-4})` and `ABS({-3|-4})` both select element `(0,0)` in scalar mode;
   matrix mode produces per-position results.
2. `1+{1;2;3|4;5;6}`, the row/column broadcasts above, and
   `{1;2}+{3;4;5}` prove rectangular max-shape and out-of-range `#N/A`.
3. `{1;2|3}` parses as a ragged array but is rejected by a rectangular array
   evaluator; `{}` and `{1;}` are syntax errors.
4. In B2, `[.A1:.C1]` selects B1 and `[.A1:.A3]` selects A2; outside those
   vectors the result is `#N/A`.
5. `[.A1:.B2]`, `[Sheet.A1:.B2]`, `[Sheet1.A1:Sheet2.B2]`, `[.A:.C]`, and
   `[.1:.3]` exercise local, inherited, cuboid, whole-column and whole-row
   forms. A large row/column must hit a finite bound before materialization.
6. `[.A1:.B2]~[.B2]` preserves two areas and the duplicate B2; intersecting it
   with `[.B1:.C3]` omits empty combinations but returns the surviving area.
7. A resolver cell containing Empty converts to 0, FALSE, or empty Text only
   in the corresponding conversion; `{\"\"}`, Missing, `#N/A`, and an
   unsupported external Source remain distinct.
8. In `IF(TRUE();1;[.A1:.B2])`, the unselected reference is never requested.
   An eager function requests all present references in source order and
   propagates the first formula Error without catching resolver failures.
9. For `AND`/`OR`, a referenced range containing Number, Empty and Text cells
   proves that only Number (and non-distinguished Logical) sequence elements
   participate; a range containing no eligible values exercises the explicitly
   documented empty-sequence identity or Error profile. A referenced Error
   must propagate. Separately, `AND({TRUE();FALSE()})` and an array containing
   Empty/Text elements exercise the chosen Array-element profile; the standard
   does not mandate that profile. `AND`/`OR` return one Logical aggregate in
   matrix context.
10. `XOR({TRUE();FALSE()})` in matrix context iterates per output position,
    while the analogous `AND` returns one aggregate Logical; this contrast does
    not by itself specify how distinguished Logical Array elements are typed.
    `XOR` and `NOT` given a ReferenceList refuse it at the scalar type
    boundary.
11. In formula `B1`, `[.A1:.C1]` selects `B1` under the §3.3 vector-specific
    branch; in `A1`, `[.A1:.A3]` selects `A1`; `D1` and `A4` are outside the
    respective spans and return `#N/A`. A two-dimensional range in a position
    where its current-row/current-column union has multiple cells returns an
    Error under §6.3.3.
12. `{1}:[.A1]` and `[.A1]~{1}` refuse their Array operands as
    reference-operator type errors; `[.A1]~[.B2]` remains a ReferenceList and
    cannot be converted to an Array.
13. With formula position `Main.B2`, `[Data.A1:.A3]` selects `Data.A2` and
    `[Data.A1:.C1]` selects `Data.B1`; changing only the formula sheet does not
    change the coordinate match.
14. `[Sheet1.A1:Sheet3.A1]` selects `Sheet2.A1` only when the formula is on
    `Sheet2` at `A1`; a formula outside the cuboid has no candidate, and a
    3-by-3 cuboid at an interior formula position fails the exact-one test.
    No case selects the first sheet merely because it is listed first.
15. Evaluating a bare `[.A1:.A3]` through a value API preserves a Reference;
    an explicit scalar/display projection then applies the current position's
    intersection. Reference operators likewise retain Reference/ReferenceList
    results until a consumer requests a value.

Implementation review is pending against the frozen evaluator source. This
document deliberately makes no claim that the current scalar evaluator already
implements these rules.

## Empty operand profile for the scalar bridge

Section 4.7 distinguishes Empty from Number zero, empty Text, FALSE and Error.
Section 6.3.2 dereferences a single-cell Reference without defining a generic
Empty-to-Scalar conversion. Sections 6.4.7–6.4.9 specify comparisons for
Number, Text and Logical but do not define direct comparisons involving Empty.
The database-criterion rule in §4.11.8 does not supply general operator semantics.

This evaluator's explicit profile is `Empty = Empty` → TRUE, equality between
Empty and a typed scalar → FALSE, inequality as the negation of equality, and
ordered comparison involving Empty → `#VALUE!`. Formula Errors retain the
existing left-to-right precedence. These Empty cases are implementation
choices, not claims that the specification mandates these comparison results.

Prefix `+` accepts Any and preserves Empty (§6.4.15); it must not silently
produce Number zero. Numeric, Logical and Text operands instead convert Empty
to zero, FALSE or empty Text at the corresponding parameter boundary. Thus a
blank passed to `ARABIC(Text)` becomes empty Text and yields zero (§6.19.2).
A blank passed to `DECIMAL(Text; Integer)` becomes empty Text in its first
parameter; its result follows the established empty-text `#VALUE!` profile,
because §6.19.10 does not specify that case. For first parameters accepting
TextOrNumber in the BIN/OCT/HEX conversions, the chosen blank profile is also
empty Text, preserving the existing empty-input behavior instead of changing
it to the numeric-zero path. Section 6.19.4 explicitly permits Error or zero
for an empty binary string.
