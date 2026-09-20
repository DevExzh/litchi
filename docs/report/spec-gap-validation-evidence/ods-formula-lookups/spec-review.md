# Semantic review of the ODS lookup and reference batch

Status: independent semantic review of implementation contract v2. The
reviewed contract is `contract.md`; its current hash is recorded below after
the final report edit. Contract semantics are accepted. Source integration
remains HOLD pending the implementation findings in this report. This report
makes no production-support, gate, native, or performance claim; the source
is not frozen.

The review covers `ADDRESS`, `CHOOSE`, `HLOOKUP`, `INDEX`, `INDIRECT`,
`LOOKUP`, `MATCH`, `OFFSET`, and `VLOOKUP` from OpenDocument 1.4 Part 4
§6.14.2, §6.14.3, §6.14.5–§6.14.9, §6.14.11, and §6.14.12. The retained
normative archive is
`3rdparty/specs/OpenDocument-v1.4-os.zip` (SHA-256
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`), whose
formula member
`part4-formula/OpenDocument-v1.4-os-part4-formula.html` has SHA-256
`ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.
The implementation baseline is `635fd2e1348b621426b50909cbd5765c91837306`.

`GETPIVOTDATA` (§6.14.4) and `MULTIPLE.OPERATIONS` (§6.14.10) remain outside
this family. They require pivot-table metadata and dependency recalculation,
respectively, and adding them would hide separate host-service dependencies.

## Normative decisions

`CHOOSE` has an `Integer` selector and variadic `Any` values. In scalar mode,
the selector uses the ordinary scalar conversion, so an inline array selects
element `[0,0]` and a multi-cell reference uses the existing implied
intersection. In matrix mode the scalar selector is evaluated per output
position. Only the selected value expression is evaluated. Since `Any` means
no conversion under §6.3.1, a selected Reference or ReferenceList remains a
descriptor and an unselected branch performs no reads or provider work. The
sentence in the INDEX clause describing CHOOSE as not accepting range
parameters does not add a constraint to CHOOSE's explicit `Any` parameters;
the contract should preserve the selected descriptor and test it directly.

`INDEX` is the one consumer that explicitly accepts `ReferenceList|Array`.
`AreaNumber` selects an ordered logical record, preserving duplicate records;
a multi-sheet cuboid is one record and its selected geometry retains all of
its physical planes. Row and column values of zero, a syntactic empty slot,
or omission have the exact §6.14.6 slice meanings: zero row selects a column,
zero column selects a row, and both select the whole area. A Reference input
returns a derived Reference descriptor for these selections and defers cell
reads until a value consumer demands them. An Array input returns a scalar or
bounded array value. This reference-preserving interpretation is required for
`ISREF`, metadata functions, reducers, and reference operators to compose with
INDEX; the contract must label it as the evaluator profile for the `Returns:
Any` clause.

`HLOOKUP`, `VLOOKUP`, `LOOKUP`, and `MATCH` accept one direct Reference or an
Array in their ForceArray data parameter. A runtime ReferenceList is rejected
before cell reads because decomposing a list can change vector orientation,
length, or table geometry. The same rejection applies to `OFFSET`, whose
parameter is a strict Reference. A direct 3-D Reference is retained for
descriptor-preserving `INDEX` and `OFFSET`; search consumers use a two-
dimensional ForceArray profile and refuse multi-plane search data before cell
reads rather than silently flattening sheets or choosing the current plane.
This is an explicit compatibility boundary for an otherwise underspecified
3-D search, not a claim that the normative Reference type lacks cuboids.

Lookup keys are `Any` for H/V/LOOKUP and Scalar for MATCH. Supported keys are
Number, Text, Logical, and Empty; Complex and other non-scalar values return a
formula value error. A scalar key intersects/selects one element. In matrix
mode, a scalar expected key is evaluated at each output position, while every
ForceArray data source remains complete. Empty cells are normalized to numeric
zero for this comparator. A syntactically missing optional flag uses its
declared default; an actual referenced Empty value converts to zero or false
and does not select the missing default. Formula Error keys propagate before
the data source is read.

The fixed comparison profile is Number < Text < Logical, with FALSE before
TRUE. Text comparisons use the evaluator's pinned Unicode 17 full case fold
without changing generic scalar or conditional-criterion comparison. Exact
searches visit candidates in source order and select the first equal value.
Approximate searches assume the specified sorted order, may use a bounded
binary/upper-bound search, and select the last duplicate satisfying the
direction. They do not validate sortedness. The §6.14 Number/Text mismatch
barriers apply after the candidate is selected: ascending Text-key/Number-
candidate and descending Number-key/Text-candidate produce #N/A. The search
constraints permit evaluators to process Logical candidates; this profile
does so using the stated ordering.

Formula Errors encountered in visited search cells or in the selected result
remain formula outcomes. Exact scans retain an encountered formula Error and
continue only as required to preserve typed-failure precedence; a match before
an unvisited error may short-circuit. Approximate search need only observe the
cells used by its bounded search, so errors in unvisited cells do not affect
the result. Resolver/provider, cancellation, resource, source-version, and
unsupported-source failures are typed `EvaluationFailure`s and supersede a
retained formula Error; they are never converted to #N/A/#VALUE or caught by
IFERROR/IFNA.

`LOOKUP` uses the normative two-argument orientation: square/tall data searches
the first column and returns the last column; wide data searches the first row
and returns the last row. The three-argument form searches the corresponding
first vector and returns the matching position in the result vector. An Array
result that is too short produces a formula error. A direct cell-range result
may extend only when the selected match lies outside the original vector; a
single-cell reference grows as the documented column vector, and extension
past sheet limits returns #N/A. A longer result is not shortened.

`ADDRESS` is pure text construction. It supports all four absolute modes and
both A1 and R1C1 output in this profile. R1C1 relative brackets contain the
supplied row and column numbers; they are not computed by subtracting the
formula position. Sheet names are quoted and apostrophes doubled when needed;
A1 uses `.` and R1C1 uses `!`. `INDIRECT` is a scalar selector because it
returns a Reference: array text input uses `[0,0]` in matrix mode. It parses
the required A1 `.` separator and the supported `!` form, quoted names,
whole-row/column references, 3-D references, and R1C1 absolute/relative text.
R1C1 relative offsets use the explicit formula position. It constructs a
descriptor without cell reads. Source-qualified text remains inert or typed
Unsupported under the explicit-provider profile.

`OFFSET` is also a descriptor-returning scalar-selector function. It preserves
the input's sheet planes, applies checked signed row/column offsets, and
validates every resulting plane against resolver extents. Omitted dimensions
use the input dimensions; the explicitly empty height slot allowed by §6.14.11
also means omitted height before a supplied width. An actual Empty value is
converted to zero and fails the strictly-positive dimension check. Reference
lists, overflow, out-of-bounds geometry, and source references are refused
before cell reads.

All Integer parameters use the shared Number conversion and the documented
profile truncates toward zero where §6.3.6 leaves rounding implementation
defined. Formula Errors propagate before conversion. Optional Missing slots
use function defaults; syntactic empty slots with a function-specific rule
(INDEX row/column and OFFSET height) use that rule.

## Contract review disposition

Contract v2 resolves the prior semantic findings:

* valid finite Text uses the existing fast-float Number conversion;
  malformed/non-finite Text and Text supplied to Logical parameters have
  explicit formula-error behavior;
* Empty search keys/cells normalize to numeric zero for both equality and
  ordering;
* CHOOSE distinguishes a scalar selector that preserves a selected descriptor
  from an array-valued selector that publishes an Array of selected values;
* exact-search formula Errors retain first-error precedence while the remaining
  search vector is visited only to expose a typed failure; and
* short Array Results in LOOKUP use the explicit profile `#N/A`, bound for the
  oracle and focused tests separately from no-match and extension failures.

The nine-function scope, ReferenceList/3-D admission, scalar-selector
contexts, defaults, and error spellings now have a coherent normative/profile
boundary. I found no remaining contract-level disagreement.

## Source review flags

The source remains HOLD until these contract obligations are implemented and
verified:

* nonintegral row/column/index/area/match parameters were rejected rather
  than truncated toward zero;
* Error values supplied as `RangeLookup` or `MatchType` were mapped to an
  exact-search mode rather than propagated before data reads;
* exact-search formula Errors need retained state so a later typed provider
  failure supersedes them; and
* LOOKUP result-reference extension must be deferred until after the selected
  match is known to lie outside the original result vector.

The value-side reference code also needs the contract-boundary checks below:

* INDEX record selection currently retains the input `is_list` marker instead
  of collapsing the selected logical record to one derived Reference;
* INDEX must reject an Array `AreaNumber` other than 1 and must refuse scalar
  data rather than treating it as a one-cell Array;
* INDEX slice/record bounds must publish the contract's `#REF!` formula value,
  not `#VALUE!` or an internal `InvalidExpression`; and
* optional selector code must preserve the distinction between syntactic
  Missing defaults and an actual Empty value, especially for OFFSET dimensions.

These are implementation findings against the current mutable tree. They
must be cleared and rechecked against frozen source before this report can
become a semantic PASS.

## Required evidence

The semantic suite must include scalar and matrix selector cases, projected
IF/IFERROR/IFNA branches, CHOOSE unselected branches, selected references and
ReferenceLists, direct 3-D INDEX/OFFSET descriptors, rejected search lists and
3-D tables, ADDRESS quoting/A1/R1C1 output, INDIRECT relative coordinates,
LOOKUP vector extension at both an in-range and an out-of-range match, mixed
Number/Text/Logical/Empty searches, duplicate ties, case-fold pairs, and
visited versus unvisited formula Errors.

The resource suite must prove zero cell reads for descriptor construction and
shape/list refusals, ordered exact reads, bounded approximate reads, selected
result-only reads, work and cancellation charging, borrowed resolver text,
typed-failure precedence, source fences, and reservation cleanup. Matrix and
cache tests must keep complete ForceArray/reference descriptors while scalar
keys/selectors retain output-position semantics; nested position-sensitive
descendants such as MUNIT must not be reused as invariant scalar cache keys.

No source-freeze hash, gate receipt, native result, or performance conclusion
is recorded here. Contract semantics are accepted; only the listed source
integration findings and their focused tests remain before a final disposition.
