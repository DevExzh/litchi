# ODS lookup and reference-function contract

Status: implementation contract v2; independent specification review pending.
No source freeze, passing tests, native compatibility or performance acceptance
is claimed by this contract. Baseline is `635fd2e1348b621426b50909cbd5765c91837306`.

## Authority and scope

Normative authority is OpenDocument 1.4 Part 4 §§6.14.2, 6.14.3, 6.14.5–6.14.9,
6.14.11 and 6.14.12, with §§3.2.3, 3.3, 4.7–4.11, 5.6, 5.8–5.9 and 6.3.
The local archive `3rdparty/specs/OpenDocument-v1.4-os.zip` has SHA-256
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`;
member `part4-formula/OpenDocument-v1.4-os-part4-formula.html` has SHA-256
`ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.
The [published specification](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/part4-formula/OpenDocument-v1.4-os-part4-formula.html)
is the external reference; retained hashes identify the exact review input.

This batch implements ADDRESS, CHOOSE, HLOOKUP, INDEX, INDIRECT, LOOKUP,
MATCH, OFFSET and VLOOKUP in the existing explicit read-only evaluator.
GETPIVOTDATA needs pivot calculation and MULTIPLE.OPERATIONS needs dependency
recalculation/substitution; both remain separate tracked gaps. Named-expression
resolution, external-workbook access, recalculation and cache publication are
not supplied by the current Resolver and remain typed capability boundaries.
No package mutation or ambient fetch is introduced.

Normative obligations are distinguished below from choices in this evaluator's
profile. A native application's different result does not silently replace this
contract. Generic errors whose exact spelling is unspecified by the standard
use the explicit profile below.

## Common conversion and evaluation profile

- Integer parameters apply the existing Number conversion and truncate toward
  zero (§6.3.6 permits this choice where the function gives no rounding rule).
  Logical converts to 1/0. Text-to-Number uses the existing finite fast-float
  parsing profile; malformed or non-finite Text produces `#VALUE!`. Logical
  parameters refuse Text, including numeric Text. Non-finite numbers and
  unrepresentable integer domains are errors;
  saturating casts must not silently admit an out-of-range value.
- Logical parameters use the existing scalar Logical profile: Number zero is
  false, nonzero true; invalid Text does not silently select a search mode.
- Text parameters admit Number and Logical via the existing deterministic Text
  conversion. They preserve Text bytes. Actual Empty converts to zero, false
  or empty Text according to the expected type.
- An omitted optional argument or explicit missing AST slot takes its declared
  default. A referenced Empty cell is an actual value, not an omitted argument.
  Required missing parameters produce formula `#VALUE!`.
- Formula errors retain their original kind through conversion. Ordinary eager
  child evaluation is retained; a typed provider/resource/cancellation/source
  failure never becomes a formula error and is never caught by IFERROR/IFNA.
- The scalar entry point has no Resolver or Position. It implements ADDRESS,
  lazy scalar CHOOSE and known scalar type/arity refusals. Valid reference
  operations requiring context return typed `Unsupported(Reference)`. It must
  not invent a formula origin when checking relative R1C1 syntax.
- Reference lists retain their identity, order and duplicates. They are not
  arrays. INDEX explicitly admits lists and selects one logical record;
  CHOOSE may return the selected list unchanged. Search data and OFFSET refuse
  all lists, including one-record lists, before reading cells.
- Existing scalar/matrix publication rules continue to apply. Reference
  constructors return descriptors, not eagerly read values. A later consumer
  determines whether to intersect, iterate, reduce or inspect the descriptor.

## ADDRESS — §6.14.2

Arity 2–5: Row, Column, optional Abs=1, A1Style=TRUE, optional Sheet.
Row and Column must be positive; Abs must be 1–4. Invalid conversion/domain
produces `#VALUE!` (or the original formula error). The supported coordinate
integer domain is the checked machine-size coordinate domain; formatting does
not consult a resolver or clip to an arbitrary workbook's used extent.

For A1, modes 1–4 produce `$A$1`, `A$1`, `$A1`, `A1`, respectively. Columns
use uppercase Latin letters. R1C1 is a supported extension explicitly permitted
by the clause: modes 1–4 produce `R1C1`, `R1C[1]`, `R[1]C1`, `R[1]C[1]`.
The supplied numbers are emitted directly; ADDRESS does not subtract the
formula position for relative R1C1 modes.

A nonempty Sheet is escaped/quoted as needed for the canonical reference
parser, with doubled apostrophes. The separator is `.` for A1 and `!` for
R1C1. This profile treats omitted or empty Sheet as no sheet prefix. Output
has no enclosing reference brackets. ADDRESS scalar parameters participate in
ordinary matrix iteration. Output Text and formatting work are bounded.

## CHOOSE — §6.14.3

Arity at least 2: Integer Index followed by one or more Any branches. Index is
one-based after truncation; an invalid or out-of-range index is `#VALUE!`.
Evaluate Index, then only the selected expression. Unselected branches are
neither evaluated nor probed for metadata, shape, errors or provider work.
A selected scalar-index Reference, ReferenceList, Array or formula Error keeps
its runtime identity in both scalar and matrix evaluation modes, including
through IF/IFERROR/IFNA and metadata consumers.

This profile supports matrix Index iteration: select a branch separately for
each index coordinate, with ordinary singleton/row/column broadcasting and
selected-branch shape planning. Only branches selected by at least one valid
coordinate may contribute shape or evaluation work. Invalid index coordinates
produce formula errors in their output cells. An array-valued Index constructs
an output Array of selected values; it does not construct an array of reference
descriptors. This differs from a scalar Index evaluated in matrix mode, which
preserves the selected descriptor. An external reference selected for a cell of
the constructed Array requires the unsupported external provider. Scalar Index
selection preserves the existing metadata
source-reference policy just as scalar IF selection does.

## INDEX — §6.14.6

Arity 1–4 is accepted: DataSource, optional Row=0, Column=0, AreaNumber=1.
The one-argument spelling follows the clause's both-selectors-omitted behavior.
DataSource is ReferenceList or Array; an ordinary Reference supplies one record.
Its expression is evaluated completely in the caller's mode, not forcibly in
matrix mode. Scalars cannot masquerade as a one-cell data Array.

Row/Column zero, omitted or missing select the entire corresponding dimension:
zero Row selects a column; zero Column selects a row; both zero select the
whole selected area. Nonnegative selectors are one-based otherwise. Negative
selectors/conversion refusal produce `#VALUE!`; selectors beyond the selected
shape and invalid AreaNumber produce `#REF!`. Array DataSource requires
AreaNumber=1. Formula errors in parameters propagate without being converted
into bounds errors.

AreaNumber counts logical records in reference-list order, including duplicates,
not the physical sheet planes of a 3-D cuboid. Selecting a record returns a
single derived Reference (list identity cleared), preserving its sheet span and
sliced row/column geometry. Construction performs zero cell reads. Array input
returns the selected scalar/Array contents without inventing reference identity;
all unselected input cells remain uninspected after expression evaluation.

INDEX can return a non-scalar object and follows §3.3.2.2.1: scalar selectors
consume element (0,0) of Array arguments in matrix mode, rather than producing
an array of references. Reference selector conversion uses the existing scalar
parameter/intersection rules. The returned complete descriptor or slice remains
available to the outer consumer.

## OFFSET — §6.14.11

Arity 3–5: Reference R, RowOffset, ColumnOffset, optional NewHeight/NewWidth.
Offsets are signed integers after truncation. Omitted/missing dimensions use the
original dimensions; explicit dimensions must be positive (`#VALUE!` otherwise).
Actual Empty dimensions convert to zero and are rejected, not treated as omitted.

Return a derived Reference with shifted top-left and selected dimensions.
Preserve every plane of a 3-D cuboid and validate resulting bounds against every
participating sheet extent. Reject ReferenceList and other wrong input kinds as
`#VALUE!`. Negative resulting coordinates, overflow or out-of-sheet bounds are
formula `#REF!`, not internal or provider failures. No cells are read. Scalar
selector Array arguments follow the same first-element rule as INDEX.

## INDIRECT — §6.14.7

Arity 1–2: Text Ref, optional Logical A1=TRUE. Parse local A1 cell, range,
whole-row/column and canonical 3-D references, preserving quoted sheet names.
Both `.` and `!` sheet separators are supported outside quoted names/source IRIs.
An unqualified address uses the explicit current sheet. Sheet lookup is exact.

R1C1 is supported: absolute coordinates, bracketed signed relative offsets,
omitted current row/column offsets, ranges and whole axes. Relative offsets
use the explicit formula position with checked arithmetic. ADDRESS's R1C1
output must round-trip. Canonical source/subtable syntax remains inert and
obeys existing typed capability boundaries; no textual formula expression is
executed as a substitute for a reference. Malformed reference text or invalid
resolved coordinates produces `#REF!`; valid unsupported capabilities remain
typed refusals. The implementation must distinguish these paths.

The result is a derived descriptor whose retained sheet names borrow the stable
Resolver catalog, not temporary parser-owned strings. Parsing storage and text
work are admitted before allocation. Source-qualified text may retain the
existing source marker for ISREF/SHEET/SHEETS descriptor inspection; consumers
requiring external geometry or values return typed Unsupported(Reference).
Scalar argument Arrays use element (0,0) per §3.3.2.2.1, including the normative
SUM(INDIRECT({"A1";"A2"})) first-reference behavior in matrix mode.

## Search data and comparison profile

HLOOKUP/VLOOKUP/MATCH DataSource and LOOKUP Searched/Results are ForceArray
Reference-or-Array parameters: preserve the complete expression and descriptor.
Data must be one 2-D rectangle/Array; 3-D inputs, ReferenceLists and scalar data
are formula `#VALUE!` before cell reads. MATCH and LOOKUP Results must be vectors.
A scalar formula Error supplied as data propagates its original error.

Search uses borrowed scalar values. Exact matching is type-preserving (Number
is not Text or Logical). The supported Logical extension orders Number before
Text before Logical, with FALSE before TRUE. Text uses pinned Unicode 17 full
C+F case folding and lexicographic folded-scalar order, without normalization.
This comparison owner does not alter ordinary scalar/conditional-criterion
comparison profiles. Wildcards/regular expressions are disabled and comparison
applies to the complete cell: `*`, `?` and comparison-looking Text are literal.

This profile normalizes Empty search keys/cells to numeric zero for both
equality and ordering: an Empty candidate exactly matches a Number zero key.
Complex values
are not comparable and produce `#VALUE!`. Lookup-key formula errors propagate.
Formula errors in visited search cells propagate. Exact search stops at its
first match when no earlier formula error was observed. After a formula error,
retain the first error and scan the remaining search vector so a later typed
provider, resource, cancellation or source failure supersedes it; a later match
does not erase the retained error, and no result cell is read for that failed
search. Sorted approximate search may use binary search and need not read
unvisited cells. Sorted input is a caller obligation under this comparator;
unsorted approximate results are implementation-dependent and not normative
oracle successes. No full-range scan is needed to validate the sort assumption.

For HLOOKUP/VLOOKUP/LOOKUP's Any key, this profile resolves reference keys to
values and supports matrix iteration of Array/reference keys like the scalar
MATCH key. Scalar mode uses ordinary implicit intersection/first-element rules.
ReferenceList keys are refused. Scalar mode/index parameters of searches also
follow ordinary scalar/matrix iteration; ForceArray data does not participate
as an output-shape parameter. This explicit Any-key profile is not a general
conversion of every Any parameter (INDEX/CHOOSE retain their own semantics).

## MATCH — §6.14.9

Arity 2–3: Search, SearchRegion, optional MatchType=1. MatchType is truncated
then must be -1, 0 or 1; invalid values are `#VALUE!`. The region is a single
row or column. Return a one-based Number index.

Type 0 examines source order and returns the first exact match. Type 1 chooses
the last position whose value is <= Search in ascending sorted order. Type -1
chooses the last position whose value is >= Search in descending sorted order.
No qualifying value is `#N/A`. Ascending Text queries cannot fall back to Number;
descending Number queries cannot fall back to Text (`#N/A` in either case).

## HLOOKUP / VLOOKUP — §§6.14.5, 6.14.12

Arity 3–4: Lookup, DataSource, result Row/Column, optional RangeLookup=TRUE.
The result selector is one-based: nonpositive is `#VALUE!`, beyond the data
shape is `#REF!`; check it before searching cells. False/zero RangeLookup selects
first exact match in the first row/column, scanning left-to-right/top-to-bottom.
True/nonzero selects the last qualifying <= key in ascending sorted order,
including the last exact duplicate. No match is `#N/A`; a Text query cannot
fall back to a Number candidate. Read only the selected result cell after a
match. Return the selected value, preserving Empty, Text, Logical and formula
Error identity. Unselected result cells are not read.

## LOOKUP — §6.14.8

Arity 2–3: Find, Searched, optional Results. Search is ascending approximate,
last duplicate <= Find, with the same no-match and Text-to-Number barrier.
For square/tall Searched, inspect the first column; for wider-than-tall data,
inspect the first row. Without Results, return the corresponding value from the
last column/row, respectively.

With Results, require a row/column vector and use the match's same index,
independent of result orientation. The detailed extension semantics take
precedence over the introductory equal-length constraint: lengths may differ.
If an Array result lacks the selected index, return `#N/A` in this profile.
For a Reference result, extend only when the selected index lies beyond its
existing length, in its vector direction, to the searched vector's length.
A single-cell reference extends downward. A longer result is not shortened.
If necessary extension exceeds sheet bounds, return `#N/A`. An in-range match
must not be rejected merely because a hypothetical full extension would fail.
Only the selected output cell is read; extension is descriptor arithmetic.

## Resource, cache and validation obligations

All reads use charge_cell_work, read_reference_cell and read_to_element with
borrowed Text. No search materializes an input reference range. Exact traversal
uses O(1) retained search state; approximate numeric search uses logarithmically
bounded reads, retaining duplicate-last semantics. AST Arrays, output Arrays,
reference metadata and temporary parser storage have checked finite reservations
with allocation-before-token drop ordering. Geometry limits apply even to
zero-read descriptors. Long text work and all search steps check cancellation.

Source-version and cancellation fences remain before/after complete evaluation.
Typed provider errors from metadata and shape discovery propagate unchanged;
public Unsupported/InvalidExpression values are not internal shape-probe signals.
Known arity/type/shape/index refusals perform no cell reads, except evaluation of
computed children genuinely required to discover those values.

Demand caching retains only supported invariant scalar payloads. Safe direct
search data may be complete invariant arguments; computed projected data and
scalar descendants such as MUNIT remain position-sensitive. Descriptor, Array
and Text results are not inserted into a scalar-only cache. CHOOSE never inspects
unselected branches merely to establish cacheability.

Acceptance requires each requirement in coverage-requirements.json to be bound
to focused tests and independent evidence, required oracle tests without ignored
placeholders, independent semantic/resource source reviews, all seven isolated
gates, native compatibility observations, frozen performance inputs/raw samples,
retained-evidence verification and owned temporary-build cleanup. No passing
claim may be inferred from this planning contract alone.
