# ODF 1.4 reference and worksheet metadata contract

Status: normative implementation contract for the next ODS formula batch. This
document records the selected evaluator profile and the evidence required for
implementation; it does not claim that production support or validation is
complete.

The batch covers the eight OpenFormula 1.4 §6.13 functions that inspect
reference geometry or worksheet metadata:

`AREAS`, `COLUMN`, `COLUMNS`, `ISREF`, `ROW`, `ROWS`, `SHEET`, and `SHEETS`.

The sixteen scalar value-inspection functions are covered by the preceding
value-inspection contract. This batch does not add a public workbook metadata
API, resolve external workbooks, dereference cells merely to answer a geometry
question, or turn a provider failure into a formula error.

## Normative source

The source is the repository-local OpenDocument distribution. A web copy or a
spreadsheet application's compatibility behavior is not a normative substitute.

| Source | SHA-256 |
| --- | --- |
| archive `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| member `part4-formula/OpenDocument-v1.4-os-part4-formula.html` | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |

The primary entries are §§6.13.2, 6.13.4, 6.13.5, 6.13.24, and 6.13.29–6.13.32.
The type and evaluation rules used here are §§3.2.3, 3.3, 4.8, 4.9, 4.11.2,
4.11.13, 5.6, 5.8, 5.9, and 6.3. The local archive also states that a
ReferenceList contains one or more references, retains reference order, cannot
be converted to an Array, and is an error when passed to a function that does
not accept it.

This contract follows accepted ADRs 0001, 0002, 0003, 0004, 0005, 0006, 0008,
0010, 0011, 0023, and 0024. In particular, the evaluator remains under the
ODS family owner, keeps typed provider failures distinct from formula values,
uses explicit caller context, and applies bounded work, storage, cancellation,
and source-version rules.

## Runtime profile and terminology

The value evaluator receives an explicit `Context` containing a
`Position { sheet, row, column }`, an `ExecutionContext`, and a `Mode`. Its
`Resolver` already supplies:

* `sheet_extent(sheet)`, for checked whole-row and whole-column geometry;
* `sheet_index(sheet)`, returning a stable zero-based workbook order;
* `sheet_name_at(index)`, returning a borrowed name in that order; and
* `sheet_count()`, returning the complete workbook count.

The resolver's ordered sheet set is the workbook profile. The host must include
hidden sheets in that set; the metadata functions never filter them. The
current worksheet resolver indexes every supplied sheet in source order and
does not need a visibility flag to count or number hidden sheets.

`RuntimeValue::Areas` retains a `RuntimeAreaSet`. A direct reference has one
reference record and may have several physical sheet planes. A reference list
retains one record per reference in source/operator order. `is_list` identifies
the list kind. The physical `areas` vector is therefore not the number returned
by `AREAS`: a three-dimensional cuboid has one reference record even when it
has several sheet planes. Public `Area` bounds may merge contiguous planes;
its sheet extent is still part of the one cuboid.

The terms below have these meanings:

* **cell read** means a resolver `read_cell` operation. Metadata calls such as
  `sheet_index` and `sheet_extent` are not cell reads.
* **formula Error** means a `ScalarError` carried as a formula value, such as
  `#VALUE!` or `#REF!`.
* **typed failure** means `EvaluationFailure::Unsupported`, cancellation,
  resource, allocation, execution, or source-version failure. A typed failure
  is never caught by these functions or converted into a formula Error.
* **one-entry list** means a ReferenceList with exactly one retained reference
  record. A three-dimensional reference remains one entry.

All one-based results below are checked conversions from the resolver's
zero-based coordinates or sheet indices. A checked overflow is a typed
invalid-expression/resource failure, never a wrapped number.

## Arity and pseudotypes

| Function | ODF syntax | Accepted count | Result |
| --- | --- | ---: | --- |
| `AREAS` | `AREAS( ReferenceList R )` | exactly 1 | Number |
| `COLUMN` | `COLUMN( [ Reference R ] )` | 0 or 1 | Number or a row array of Numbers |
| `COLUMNS` | `COLUMNS( Reference\|Array R )` | exactly 1 | Number |
| `ISREF` | `ISREF( Any X )` | exactly 1 | Logical |
| `ROW` | `ROW( [ Reference R ] )` | 0 or 1 | Number or a column array of Numbers |
| `ROWS` | `ROWS( Reference\|Array R )` | exactly 1 | Number |
| `SHEET` | `SHEET( [ Text\|Reference R ] )` | 0 or 1 | Number |
| `SHEETS` | `SHEETS( [ Reference R ] )` | 0 or 1 | Number |

The empty parameter list is a zero-argument call. It is not one argument that
happens to be empty. `COLUMN()`, `ROW()`, `SHEET()`, and `SHEETS()` use their
specified defaults. A required slot that is explicitly present but empty, a
missing required argument, an extra argument, and an unaccepted pseudotype
produce formula `#VALUE!` after the evaluator's ordinary AST/arity checks.
The optional functions do not treat an explicit empty slot as an omitted
argument. This keeps omission distinct from an Empty cell and from empty Text.

The ODF parameter syntax permits empty parameter slots lexically, but does not
require every function to accept them. This profile accepts only the omission
forms in the table and records explicit empty slots as invalid arguments.

## Reference and ReferenceList admission

`AREAS` accepts a direct Reference (counted as one retained record) and accepts
an arbitrary nonempty ReferenceList. It returns the number of retained reference
records, in list order, including duplicates. A direct 3-D Reference returns
`1`; a two-reference concatenation returns `2`. It never counts one record once
per sheet plane. A formula Error propagates as a formula value; scalar values,
Arrays, Text, and Logical values that are not reference descriptors return
`#VALUE!` without a cell read.

The other six Reference parameters follow the explicit admission rule in §§4.9
and 5.9. Those sections do not define a general list-to-Reference conversion:
they admit a ReferenceList only for a function where the computation for one
Reference is identical to the computation for an arbitrary sequence of single
references occupying the identical cell range. The normative `COLUMNS` example
is the counterexample: a rectangular Reference reports its two-dimensional
column extent, while decomposing that same range into single-cell References
would report the sequence length. A one-entry list is still a ReferenceList
for this rule; it is not a degenerate conversion to a Reference, and its one
retained record can itself cover a multi-cell range.

The relevant sentence says that a ReferenceList “can be passed as an argument
to functions where passing one reference results in an identical computation as
an arbitrary sequence of single references occupying the identical cell range.”
That is a function-level admission condition over the arbitrary decomposition;
it does not say that a list with one retained record is automatically a
Reference. Sections 4.9 and 5.9 also make a list an Error when the receiving
function does not meet that condition.

Accordingly, this profile accepts a direct Reference for `COLUMN`, `COLUMNS`,
`ROW`, `ROWS`, `SHEET`, and `SHEETS`, but rejects every ReferenceList for
those six functions, including a list containing exactly one retained
Reference. The refusal is a pseudotype decision before cell reads. A list is
never flattened, materialized as an Array, or implicitly intersected merely
because the caller requested scalar mode.

`COLUMN` and `ROW` additionally enforce `AREAS(R)=1` for a direct Reference. A
3-D cuboid remains one Reference and satisfies this constraint. A ReferenceList
does not become a direct Reference by having one record, so it is refused before
that area check. `COLUMNS`, `ROWS`, `SHEET`, and `SHEETS` likewise require the
direct Reference pseudotype named by their syntax.

`ISREF` is the exception: every Reference and every ReferenceList, including a
multi-entry list and a source-qualified reference descriptor, produces `TRUE`.
It observes the runtime kind and does not dereference or resolve cell values.
An invalidated `ReferenceError` is already a formula Error value and produces
`FALSE` under the explicit `ISREF` rule.

ReferenceList refusal is a shape/type decision. If a computed expression has
to be evaluated to discover that it produced a list, that upstream expression
may do its own work; once the descriptor is known to be inadmissible, the
metadata function itself performs zero cell reads.

## Function semantics

### `AREAS`

`AREAS(R)` returns the checked `RuntimeReference` record count. The count is
metadata-only. A direct reference, including a 3-D cuboid, is one area. The
ordered records produced by `~` are counted separately, with duplicate and
overlapping references retained. A formula Error input propagates as that
Error; a non-reference value is `#VALUE!`.

### `COLUMN`

`COLUMN()` returns `Context::position.column + 1`. The position is the formula's
explicit current cell and is fixed for the evaluation; a projected matrix
output coordinate does not silently turn the no-argument call into a different
formula position.

`COLUMN(R)` uses the one admitted cuboid's column bounds. A one-column
reference returns its first one-based column as a Number. A reference covering
multiple columns returns a `1 × N` row array containing every column number in
ascending logical order. A 3-D cuboid uses its common column bounds once; it
does not repeat the array for every sheet plane and it does not read any cell.
A ReferenceList, including one containing one retained Reference, is `#VALUE!`
before this geometry operation.

The function consumes a Reference descriptor, not the value at the current
intersection. For example, with the formula position at `C5`,
`COLUMN([.B2:.D2])` computes the row array `{2,3,4}`. If a caller explicitly
requests the existing scalar publication mode, the evaluator may apply its
ordinary final projection to that generated array (the first element for an
origin-less result); the function operation must first retain the complete
array. It must never project the input reference to `C2` and return `3`.

### `COLUMNS`

`COLUMNS(R)` returns the number of columns in the one admitted Reference. A
3-D cuboid returns its two-dimensional column extent once. `COLUMNS(A)`
returns the checked column dimension of a complete inline/computed Array and
does not inspect its cell contents. Formula Error elements inside an already
evaluated Array do not alter its shape. A scalar, Text, Logical, Empty, or any
ReferenceList, including a one-entry list, is `#VALUE!`.

The Array or Reference argument is scheduled as a complete argument even in
scalar mode. Implicit intersection would lose the very shape the function is
defined to report.

### `ISREF`

`ISREF(X)` returns `TRUE` exactly when `X` has runtime type Reference or
ReferenceList. It returns `FALSE` for Number, Logical, Text, Empty, Array,
and every formula Error, including `#N/A` and `#REF!`. It does not dereference
an admitted reference and does not invoke `read_cell`. A `ScalarCell` demand
token is still a Reference for this operation and must not be mistaken for the
cell's eventual scalar value.

`ISREF` consumes a complete `Any` argument. An inline Array is classified as an
Array and returns one scalar `FALSE`; it is not elementwise-lifted merely
because the enclosing evaluation is in matrix mode.

### `ROW`

`ROW()` returns `Context::position.row + 1`, using the fixed explicit formula
position.

`ROW(R)` uses the one admitted cuboid's row bounds. A one-row reference returns
its first one-based row as a Number. A reference covering multiple rows returns
an `N × 1` column array containing every row number in ascending logical order.
A 3-D cuboid uses its common row bounds once and never repeats values per plane.
It consumes complete reference geometry, so at formula position `C5`,
`ROW([.B2:.D4])` computes `{2;3;4}` (subject to the ordinary final scalar
publication rule described for `COLUMN`). It does not implicitly intersect the
input to one cell. A ReferenceList, including one containing one retained
Reference, is `#VALUE!` before this geometry operation.

### `ROWS`

`ROWS(R)` returns the checked row extent of the one admitted Reference.
`ROWS(A)` returns the checked row dimension of a complete Array and does not
inspect Array elements. A 3-D cuboid reports its two-dimensional row extent
once. Wrong pseudotypes and every ReferenceList, including a one-entry list,
return `#VALUE!` before cell reads.

### `SHEET`

`SHEET()` returns the one-based stable resolver index of the sheet named by the
explicit current `Position`. If the current sheet is missing, the resolver's
ordinary missing-name result becomes formula `#REF!`; a provider failure
remains typed.

For `SHEET(R)`, the direct Reference is not dereferenced. It returns the
one-based stable index of the first sheet in the resolved cuboid. For a 3-D
cuboid, “first” is the lower sheet index after the resolver normalizes the
cuboid; the result is one number, not one per plane. A ReferenceList,
including a one-entry list, is `#VALUE!` at the pseudotype gate.
A reference without an explicit sheet resolves against the current Position
while building its descriptor and then follows the same metadata path.

For `SHEET(T)`, the Text overload looks up the exact sheet name through
`Resolver::sheet_index` and returns its one-based index. The standard Text
conversion is applied before lookup: Number has its ordinary locale-independent
scalar spelling and Logical becomes `TRUE` or `FALSE`. Empty, an unsupported
scalar, or a formula Error follows the conversion/error rules; a valid Text
name that is not present returns formula `#REF!`. A literal string that happens
to contain `#N/A` is Text and is looked up as a name; a formula `#N/A` value is
an Error and propagates.

In matrix mode, a non-scalar Array supplied to the Text overload is iterated
elementwise under §3.3, with scalar Number/Logical/Text conversion at each
position and a Number result per element. The output shape is the broadcast
shape of the Text array. In scalar mode, the existing explicit scalar demand
selects the array's first element before this scalar overload is applied.
A Reference argument is always the metadata overload and is never turned into
the text stored in its cell.

The profile rejects any source-qualified Reference (`Reference::Source`) with
formula `#VALUE!` before resolver access. This is distinct from a local missing
sheet (`#REF!`) and from a typed provider failure.

### `SHEETS`

`SHEETS()` returns `Resolver::sheet_count()` as a Number. The resolver profile
counts hidden sheets and preserves the complete ordered workbook set.

`SHEETS(R)` returns the number of sheets covered by the one admitted Reference.
A single-sheet Reference returns `1`; a 3-D cuboid returns its inclusive sheet
extent. The implementation must sum checked `Area.extent()[0]` values or use
the equivalent retained physical-plane count. It must not count a merged
public area as one sheet and must not count list records as sheets. A
ReferenceList, including a one-entry list, is `#VALUE!` at the pseudotype gate.
No cell is dereferenced.

The profile rejects a source-qualified Reference with formula `#VALUE!` before
resolver access. A local missing/invalid reference returns its formula
reference Error; typed resolver failures remain typed.

## Scalar and matrix integration

The distinction between input demand and result publication is required for
these functions:

* `AREAS`, `COLUMN`, `COLUMNS`, `ROW`, `ROWS`, `SHEET`, and `SHEETS` must retain
  a complete admitted Reference or Array descriptor until their function-local
  operation runs. `VisitArgument` scalar projection is not a valid shortcut.
* `ISREF` must retain complete `Any` runtime kind, including a ReferenceList or
  `ScalarCell` token. It must classify before `value_to_slot` or cell selection.
* `COLUMN` and `ROW` produce their documented arrays before any enclosing
  scalar publication. In matrix mode the array is retained. In scalar mode the
  current value evaluator's final scalar-demand boundary may project an
  origin-less generated array, but no function-local input intersection is
  permitted.
* `COLUMNS` and `ROWS` consume complete Array shape. They do not iterate the
  elements of an Array or read a Reference to discover cell values.
* `SHEET` distinguishes its Reference overload from its Text overload before
  conversion. Text arrays may iterate in matrix mode; a Reference is always a
  descriptor. `SHEET()` and `COLUMN()`/`ROW()` no-argument calls use the fixed
  explicit Position and can be broadcast as scalar results.

The function catalog and the value scheduler must agree on exact arity,
complete-argument scheduling, output shape, and reference/list admission.
Known shape/type refusals happen before `select_area_element`,
`materialize_for_array`, or `read_reference_cell`. A lazy `IF` branch that is
not selected remains unevaluated; a selected computed descriptor may perform
the work needed to produce that descriptor.

The scalar evaluator's `Evaluator` does not carry `Position`, `Resolver`, or
workbook sheet metadata. Its shared catalog may classify literal scalar values,
but it cannot faithfully execute a reference or current-sheet operation. The
scalar entry point therefore returns typed
`EvaluationFailure::Unsupported(UnsupportedKind::Reference)` for a
context-dependent `COLUMN()`, `ROW()`, `SHEET()`, `SHEETS()`, any `SHEET`
argument that would need a workbook lookup (including Text,
Number-to-Text, or Logical-to-Text conversion), or a reference descriptor that
reaches it. The value evaluator owns the complete contextual implementation.
Adding implicit global workbook state or process locale is outside this
contract.

## Errors and source constraints

Formula Errors are values. `ISREF` inspects them and returns `FALSE`. For the
other seven functions, a formula Error supplied where the function's ordinary
conversion rules accept an Error is returned as that Error. A non-error value
with an unaccepted pseudotype produces `#VALUE!`. A missing local sheet or an
invalidated local reference produces the ordinary formula `#REF!` where the
function needs a sheet/reference lookup.

The `SHEET` and `SHEETS` entries explicitly forbid Source Locations. In this
profile, a source-qualified descriptor is admitted to those function-local
checks and returns formula `#VALUE!` before any resolver lookup or cell read.
`ISREF` admits a source descriptor and returns `TRUE`, because it only inspects
the reference kind. These outcomes also apply when an `IF` expression selects
the source descriptor: selection preserves the descriptor, and the metadata
function classifies or rejects it without asking an external provider.

The limited resolver still has no external-workbook metadata provider for the
other five functions. A source leaf passed to `AREAS`, `COLUMN`, `COLUMNS`,
`ROW`, or `ROWS` therefore returns typed `Unsupported(Reference)` in this
profile. This is a deliberate capability boundary for those operations, not a
formula `#VALUE!` pseudotype conversion. A source operand used by reference
arithmetic (`:`, `!`, or `~`) also returns typed `Unsupported(Reference)` before
any metadata function can inspect the result. This keeps generic source
arithmetic refusal distinct from provider-free descriptor selection. None of
these source paths performs a resolver cell read or external fetch.

The following expressions are executable contract cases. `book.ods` is any
valid source IRI and `Data` is a local sheet in the resolver profile:

| Expression | Required result | Cell reads/provider fetches |
| --- | --- | --- |
| `ISREF(['book.ods'#.A1])` | `TRUE` | 0 / 0 |
| `ISREF(IF(TRUE();['book.ods'#.A1];[.A1]))` | `TRUE` | 0 / 0 |
| `ISREF(IF(FALSE();['book.ods'#.A1];[.A1]))` | `TRUE` for the selected local Reference | 0 / 0 |
| `SHEET(['book.ods'#.A1])` | formula `#VALUE!` | 0 / 0 |
| `SHEETS(['book.ods'#.A1])` | formula `#VALUE!` | 0 / 0 |
| `SHEET(IF(TRUE();['book.ods'#.A1];[Data.A1]))` | formula `#VALUE!` | 0 / 0 |
| `SHEETS(IF(TRUE();['book.ods'#.A1];[Data.A1]))` | formula `#VALUE!` | 0 / 0 |
| `SHEET(IF(FALSE();['book.ods'#.A1];[Data.A1]))` | the local one-based `Data` sheet index | 0 / 0 |
| `SHEETS(IF(FALSE();['book.ods'#.A1];[Data.A1]))` | `1` | 0 / 0 |
| `ISREF(['book.ods'#.A1]~[.A1])` | typed `Unsupported(Reference)` | 0 / 0 |
| `SHEET(['book.ods'#.A1]~[.A1])` | typed `Unsupported(Reference)` | 0 / 0 |

The source IRI is lexical data in these examples; it is never opened. The
selected `IF` source cases must not be replaced by a generic provider refusal,
and the unselected source cases must remain lazy.

`IFERROR` and `IFNA` preserve the same distinction. A direct source reference
is a successful Reference value, not a formula Error, so the handler returns
the source descriptor and the final metadata consumer observes its kind or
source constraint. A source arithmetic expression is a typed provider
capability failure; typed failures are not formula values and cannot trigger
the handler fallback. These cases are executable as follows:

| Expression | Required result | Cell reads/provider fetches |
| --- | --- | --- |
| `ISREF(IFERROR(['book.ods'#.A1];"Main"))` | `TRUE` | 0 / 0 |
| `ISREF(IFNA(['book.ods'#.A1];"Main"))` | `TRUE` | 0 / 0 |
| `SHEET(IFERROR(['book.ods'#.A1];"Main"))` | formula `#VALUE!` | 0 / 0 |
| `SHEET(IFNA(['book.ods'#.A1];"Main"))` | formula `#VALUE!` | 0 / 0 |
| `SHEETS(IFERROR(['book.ods'#.A1];"Main"))` | formula `#VALUE!` | 0 / 0 |
| `SHEETS(IFNA(['book.ods'#.A1];"Main"))` | formula `#VALUE!` | 0 / 0 |
| `ISREF(IFERROR(['book.ods'#.A1]+1;"Main"))` | typed `Unsupported(Reference)` | 0 / 0 |
| `ISREF(IFNA(['book.ods'#.A1]+1;"Main"))` | typed `Unsupported(Reference)` | 0 / 0 |

The `SHEET`/`SHEETS` source constraint is applied after a successful
`IFERROR`/`IFNA` source result; the handler must not substitute `"Main"` merely
because the final metadata function rejects that source descriptor.

The `SHEET` clause also says that evaluation outside a table cell is an Error.
The value evaluator's explicit `Position` is the selected worksheet-cell
context for this profile. A future host API that evaluates formulas outside a
table cell must add an explicit context bit and map that case to a formula
Error; no ambient “current cell” is permitted.

## Resource, cancellation, and cache contract

Metadata-only does not mean unbounded. The implementation must:

* charge metadata and output work through the existing scalar/evaluation
  budget and check cancellation at established loop checkpoints;
* enforce `max_reference_cells` while admitting logical reference geometry,
  even when the operation performs zero cell reads;
* enforce `max_reference_areas` and checked descriptor capacity;
* enforce `max_array_cells` before allocating `COLUMN`/`ROW` or matrix `SHEET`
  output, using fallible reservation and correct drop-before-refund order;
* keep source-version and final cancellation fences around the full evaluation;
  and
* leave provider metadata failures, cancellation, resource limits, and source
  changes as typed failures.

The zero-read assertion applies to `Resolver::read_cell` and the evaluator's
reference-cell read counter. `sheet_index`, `sheet_extent`, `sheet_name_at`,
and `sheet_count` may be called for descriptor admission and are subject to
their own work/cancellation/provider contracts.

Demand-cache classification must follow runtime position sensitivity rather
than function purity alone:

* `COLUMN()` and `ROW()` without arguments, and `SHEET()` without arguments,
  are position-sensitive and cannot share a projected scalar result across
  output coordinates.
* Explicit geometry/reference descriptors can be cached only after complete
  argument propagation proves that the descriptor and result are invariant.
  Range references must not be implicitly intersected before classification.
* `SHEET(Text)` and `SHEETS()` may use a cache when their complete arguments and
  stable resolver metadata are invariant; text-array matrix results remain
  coordinate-sensitive unless the complete array is retained.
* `ISREF` may classify an invariant complete runtime kind, including a
  ReferenceList, without reading it. A formula Error payload remains a formula
  value, not a typed cache failure.
* Nested scalar descendants under position-sensitive functions remain
  position-sensitive. Existing MUNIT scalar arguments remain excluded from
  complete-reference propagation.

No cache entry may contain a typed provider/resource/cancellation/source
failure, and no cache hit may bypass a required source or cancellation fence.

## Required contract tests

The focused validation batch should add semantic and resource tests for:

1. exact arity, omitted optional arguments, explicit empty slots, and wrong
   pseudotypes for every function;
2. direct single-cell, rectangular, whole-row, whole-column, reverse-endpoint,
   and 3-D References;
3. direct references, one-entry and multi-entry list refusals for
   `COLUMN`/`COLUMNS`/`ROW`/`ROWS`/`SHEET`/`SHEETS`, `AREAS`/`ISREF` list
   admission, duplicates, and nested reference concatenation, including
   `AREAS` record-vs-plane counts;
4. `COLUMN`/`ROW` scalar and matrix results, proving that a scalar formula at
   `C5` with `COLUMN([.B2:.D2])` retains `{2,3,4}` before the documented final
   projection;
5. `COLUMNS`/`ROWS` on complete inline Arrays containing values and formula
   Errors, with no content-dependent result;
6. `ISREF` on every scalar kind, formula Error, direct Reference, ReferenceList,
   direct and selected source Reference, source arithmetic, Array, and
   scalar-cell demand token;
7. `SHEET` no-argument current position, exact Text lookup, Number/Logical Text
   conversion, unknown names, explicit and 3-D References, text arrays in
   matrix mode, direct source refusal, and selected-source `IF` refusal;
8. `SHEETS` no-argument count, hidden-sheet inclusion, single-sheet and 3-D
   References, one-entry and multi-entry list refusals, direct and
   selected-source refusal, and source-arithmetic typed refusal;
9. formula Error precedence versus typed resolver/provider failures;
10. zero `read_cell` calls for every metadata-only success and known refusal,
    including list, source, shape, and `max_reference_cells` admission cases;
11. output-cell, metadata-work, cancellation, allocation, and source-version
    limits; and
12. lazy `IF`, projected matrix branches, direct and selected source leaves,
    source arithmetic refusal, nested geometry calls, and demand cache cases
    with different current positions.

An independent oracle should compare exact scalar values, array shapes, list
classification, sheet order, cell-read counts, and typed failures. Native
spreadsheet observations may document compatibility differences, but they do
not replace the local ODF contract.
