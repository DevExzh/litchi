# ODF 1.4 value-inspection and conversion contract

This contract defines the semantic boundary for the sixteen pure value
inspection and conversion functions in OpenFormula 1.4 §6.13:

`ERROR.TYPE`, `ISBLANK`, `ISERR`, `ISERROR`, `ISEVEN`, `ISLOGICAL`, `ISNA`,
`ISNONTEXT`, `ISNUMBER`, `ISODD`, `ISTEXT`, `N`, `NA`, `NUMBERVALUE`, `TYPE`,
and `VALUE`.

It is an implementation contract and normative review record. It does not
claim that production support or validation is complete. The twelve remaining
information functions that inspect references or host metadata are outside this
batch. In particular, this document does not close `CELL`, `COLUMN`,
`COLUMNS`, `FORMULA`, `INFO`, `ISFORMULA`, `ISREF`, `ROW`, `ROWS`, `SHEET`,
`SHEETS`, or `AREAS`.

The functions return ordinary formula values. They do not recalculate a
referenced formula cell, refresh a cached value, publish a changed document,
or convert a typed resolver/provider failure into a formula error.

## Normative sources

The source is the repository-local ODF distribution, rather than a web copy or
an implementation's compatibility documentation.

| Source | SHA-256 |
| --- | --- |
| archive `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| member `part4-formula/OpenDocument-v1.4-os-part4-formula.html` | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |
| member `part3-schema/OpenDocument-v1.4-os-part3-schema.html` | `43fb603f9f54f030db7082518aff6f136297d182abb20124a859d271a7969a15` |

The function entries are §§6.13.11, 6.13.14–6.13.17, 6.13.19–6.13.23,
6.13.25–6.13.28, and 6.13.33–6.13.34. The value model and conversion rules
come from §§3.2.3, 3.3, 4.1–4.11.13, 5.12, 5.14, and 6.1–6.3. The date
serial model comes from §4.3 and the date profile below. The two-digit-year
setting comes from Part 3 §19.678 (`table:null-year`).

The Part 4 entry for `ISNA` contains a source typo: its syntax line says
`ISERR(Scalar X)` while the heading, summary, and semantics are for `ISNA`.
This contract uses the function name and semantics, with the corrected syntax
`ISNA(Scalar X)`.

The regular expression printed for the locale-independent numeric form of
`VALUE` in Part 4 has an extra closing parenthesis in the HTML. The contract
uses the unambiguous intended grammar shown below; this corrects the editorial
typo without narrowing the required input set.

## Signatures and result classes

The signatures retain the OpenFormula pseudotypes. A semicolon separates
formula arguments; square brackets denote optional arguments. `ScalarError`
names below use the ordinary display spelling (`#VALUE!`, `#NUM!`, and so on).

| Function | Part 4 signature | Result | Formula-error treatment |
| --- | --- | --- | --- |
| `ERROR.TYPE` | `ERROR.TYPE(Error E)` | `Number` | Inspects a formula `Error`; a non-Error is `#VALUE!` |
| `ISBLANK` | `ISBLANK(Scalar X)` | `Logical` | Inspects errors and returns `FALSE` |
| `ISERR` | `ISERR(Scalar X)` | `Logical` | Inspects errors; `#N/A` is `FALSE` |
| `ISERROR` | `ISERROR(Scalar X)` | `Logical` | Inspects all formula errors, including `#N/A` |
| `ISEVEN` | `ISEVEN(Number X)` | `Logical` | Propagates a formula error; converts a Logical in this profile |
| `ISLOGICAL` | `ISLOGICAL(Scalar X)` | `Logical` | Inspects errors and returns `FALSE` |
| `ISNA` | `ISNA(Scalar X)` | `Logical` | Inspects errors; only `#N/A` is `TRUE` |
| `ISNONTEXT` | `ISNONTEXT(Scalar X)` | `Logical` | Inspects errors and returns `TRUE` |
| `ISNUMBER` | `ISNUMBER(Scalar X)` | `Logical` | Inspects errors and returns `FALSE` |
| `ISODD` | `ISODD(Number X)` | `Logical` | Propagates a formula error; converts a Logical in this profile |
| `ISTEXT` | `ISTEXT(Scalar X)` | `Logical` | Inspects errors and returns `FALSE` |
| `N` | `N(Any X)` | `Number` | Propagates a formula error |
| `NA` | `NA()` | `Error` | Exact zero-argument function; returns `#N/A` |
| `NUMBERVALUE` | `NUMBERVALUE(Text X [; Text DecimalSeparator [; Text GroupSeparator]])` | `Number` | Propagates a formula error; malformed text is `#VALUE!` |
| `TYPE` | `TYPE(Any Value)` | `Number` | Inspects formula errors and returns their type code |
| `VALUE` | `VALUE(Text X)` | `Number` | Propagates a formula error; unconvertible text is `#VALUE!` |

`NA` shall have exactly zero parameters. Extra or missing parameters are an
arity error at the function boundary. Its result is the ordinary formula
value `#N/A`, so `ISNA(NA())`, `ISERROR(NA())`, `ERROR.TYPE(NA())`, and
`TYPE(NA())` observe it as a value. The evaluator still keeps typed failures
from evaluating an argument or provider separate from this formula value.

The Part 4 entries for the predicates and `ERROR.TYPE` explicitly suppress
formula-value errors. That suppression applies only after a cell or expression
has produced a `ScalarError` value. It does not catch `EvaluationFailure`
values such as cancellation, a resource limit, a source-version change, or a
resolver/provider failure.

## Explicit profile and context

Part 4 makes locale and some date behavior implementation-dependent while
requiring particular formats. The current pure evaluator uses the fixed value
profile below. It does not currently expose a locale, epoch, or calculation-
settings argument; the profile is fixed so that this batch does not introduce a
new public context API. It is never obtained from the process locale,
environment, current time, filesystem, or another ambient provider.

The required deterministic profile for this batch is:

| Profile member | Selected value |
| --- | --- |
| locale | `en_US` |
| decimal separator | `.` |
| group separator | `,` |
| short month names | English `Jan` through `Dec` |
| long month names | English `January` through `December` |
| calendar | Proleptic Gregorian |
| serial epoch | 1899-12-30, with no synthetic 1900-02-29 |
| supported VALUE date domain | 1899-12-30 through 9999-12-31 inclusive |
| two-digit-year start in this fixed profile | 1930 |

The date epoch is the existing deterministic profile used by the ordinary text
formatter. Part 3 §19.678 defines `table:null-year` as a document calculation
setting, and the schema also defines document null-date data. The current pure
evaluator has no calculation-settings context, so this batch fixes the epoch at
1899-12-30 and the two-digit-year start at 1930. Passing document overrides
through an explicit evaluation context is a follow-up integration choice; it is
outside this batch and must not be replaced by an ambient locale, clock, or
document lookup. This limitation is recorded here so it is not mistaken for a
claim that the normative document settings do not exist. A date parser must not
silently inherit the formatter's output-range guard as its input domain; it
validates the VALUE domain stated here.

For a two-digit year `YY`, this fixed profile chooses the least year at least
1930 whose final two digits are `YY`: `30` means 1930, `29` means 2029, and
`00` means 2000. This is the Part 3 `table:null-year` rule with the profile's
fixed default, not a system-year or current-clock lookup.

Other locale profiles and document-specific date settings require a future
explicit context design. They are not silently inferred by this batch. The
fixed profile must support all required `en_US` forms below.

## Raw value model

Inspection needs the value before ordinary scalar coercion. A blank reference
cell is an `Empty` value, not Number zero and not Text `""`; a formula that
returns `""` is Text and is not blank. The function adapter therefore consumes
an input of the form `Empty | Present(value)` before a generic scalar argument
bridge turns Empty into a conversion default.

The standard value kinds used by this contract are Number, Logical, Text,
Error, Complex, Empty, Array, Reference, and ReferenceList. Number subtypes
(Date, Time, DateTime, Percentage, and Currency) remain Number for the
inspection functions. A formula Error is a value; a typed resolver/provider
failure is an evaluation failure.

The selected Complex policy is explicit because §4.4 permits an evaluator to
represent a complex number as Number or Text. This implementation represents a
distinguished Complex value as a Number subtype for raw inspection:
`ISNUMBER(Complex)` is `TRUE`, `TYPE(Complex)` is `1`, `ISTEXT(Complex)` is
`FALSE`, `ISNONTEXT(Complex)` is `TRUE`, and `N(Complex)` returns the same
Complex value through the Number result channel. Real-number conversions such
as parity, `VALUE`, and `NUMBERVALUE` cannot take a non-real Complex value and
return `#VALUE!`. This does not change the ordinary arithmetic bridge, which
may continue to reject Complex where a real Number is required.

## Empty, error, and scalar mapping

The following table fixes the behavior that is otherwise not fully represented
by the Part 4 type table.

| Raw input | `ISBLANK` | `ISERR` | `ISERROR` | `ISNA` | `ISLOGICAL` | `ISNUMBER` | `ISTEXT` | `ISNONTEXT` |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `Empty` | `TRUE` | `FALSE` | `FALSE` | `FALSE` | `FALSE` | `FALSE` | `FALSE` | `TRUE` |
| empty Text `""` | `FALSE` | `FALSE` | `FALSE` | `FALSE` | `FALSE` | `FALSE` | `TRUE` | `FALSE` |
| Number, including date subtypes | `FALSE` | `FALSE` | `FALSE` | `FALSE` | `FALSE` | `TRUE` | `FALSE` | `TRUE` |
| Logical | `FALSE` | `FALSE` | `FALSE` | `FALSE` | `TRUE` | `FALSE` | `FALSE` | `TRUE` |
| Text | `FALSE` | `FALSE` | `FALSE` | `FALSE` | `FALSE` | `FALSE` | `TRUE` | `FALSE` |
| Complex under this profile | `FALSE` | `FALSE` | `FALSE` | `FALSE` | `FALSE` | `TRUE` | `FALSE` | `TRUE` |
| `#N/A` | `FALSE` | `FALSE` | `TRUE` | `TRUE` | `FALSE` | `FALSE` | `FALSE` | `TRUE` |
| another formula Error | `FALSE` | `TRUE` | `TRUE` | `FALSE` | `FALSE` | `FALSE` | `FALSE` | `TRUE` |

`ISBLANK`, `ISERR`, `ISERROR`, `ISLOGICAL`, `ISNA`, `ISNONTEXT`, `ISNUMBER`,
and `ISTEXT` use type identity. They do not parse Text, turn Empty into zero,
or let a Text spelling such as `"#N/A"` become an Error. `ISTEXT` is the
logical complement of `ISNONTEXT` for all scalar values in this profile.

`ERROR.TYPE` maps the seven standard formula errors as follows:

| Error | Number |
| --- | ---: |
| `#NULL!` | 1 |
| `#DIV/0!` | 2 |
| `#VALUE!` | 3 |
| `#REF!` | 4 |
| `#NAME?` | 5 |
| `#NUM!` | 6 |
| `#N/A` | 7 |

The repository's `ScalarError` variants use the corresponding standard
mapping. A non-Error, including Empty, Complex, an Array, or a ReferenceList
that reaches the scalar operation, returns formula `#VALUE!`; it is not a
typed failure.

`N` uses the following selected profile for the implementation-defined cases:

* Number is returned unchanged, including a Number subtype and the selected
  Complex-as-Number value.
* Logical maps to `1` or `0`.
* Empty maps to numeric `0`.
* Text maps to numeric `0`. This is the explicit choice permitted by Part 4;
  Text is not parsed by `N`.
* A formula Error propagates unchanged.
* An inline Array always uses its `[0,0]` element for `N`, including when the
  expression is evaluated inside a projected matrix branch. It does not use
  the current output coordinate or lift elementwise; the resulting scalar is
  broadcast by the enclosing matrix operator. A Reference is first
  dereferenced to one scalar using the explicit intersection context described
  below. A ReferenceList cannot be dereferenced to a scalar and returns
  `#VALUE!` before a resolver read.

`TYPE` uses the Part 4 table, with the following complete mapping:

| Value presented to TYPE | Number returned |
| --- | ---: |
| Number, including date/time subtypes and Complex under this profile | 1 |
| Text | 2 |
| distinct Logical | 4 |
| Error | 16 |
| Array | 64 |
| Empty | 1 in this profile |

Part 4 does not give Empty a row. Mapping a blank scalar/reference cell to the
numeric-empty identity `1` is therefore a deliberate profile choice and must
be tested separately from `ISBLANK`; it must not be inferred by dropping the
blank cell or by turning empty Text into Empty. A direct ReferenceList is not
an Array in the ODF model and has no TYPE code in the Part 4 table. This
profile rejects it with formula `#VALUE!` before reading it.

`TYPE` always returns one scalar Number. It does not elementwise-lift an Array
or a rectangular Reference into a matrix. A direct single-cell Reference is
dereferenced and classified. A rectangular or multi-sheet cuboid Reference
with more than one cell is dereferenced as an Array and returns `64`, even when
one of its cells contains a formula Error. Formulas in an admitted Reference
are nevertheless evaluated before the result is published so that a typed
resolver/provider failure can bubble. The value scan is streaming and bounded;
it does not retain a cell vector. An inline Array is already an evaluated Array
and returns `64`.

## References, arrays, and matrix shape

The evaluator follows the Part 4 scalar and matrix rules while retaining the
function-specific exceptions in this section.

* A single-cell Reference is dereferenced before applying a raw predicate or a
  conversion. A multi-cell Reference in scalar context uses the evaluator's
  explicit current intersection coordinate. For a multi-sheet cuboid, scalar
  intersection first selects the plane containing the current sheet, then
  applies the same row/column union rule. A cuboid is therefore not rejected
  merely because it spans sheets. If the current sheet is outside the cuboid,
  or the row/column union does not identify exactly one cell, the result is the
  ordinary implicit-intersection formula error. The coordinate is supplied by
  the caller's scalar or projected-matrix context; it is never read from
  ambient state.
* In matrix context, `ISBLANK`, `ISERR`, `ISERROR`, `ISEVEN`, `ISLOGICAL`,
  `ISNA`, `ISNONTEXT`, `ISNUMBER`, `ISODD`, `ISTEXT`, `NUMBERVALUE`, and
  `VALUE` use ordinary elementwise lifting. The result shape is the broadcast
  shape of their matrix arguments. A scalar argument broadcasts. A direct
  rectangular Reference is selected one cell at a time in row-major order;
  it is not first materialized as an array. When the Reference is a multi-sheet
  cuboid, this profile's matrix scheduler selects the current sheet plane
  rather than flattening sheet planes into a 2-D output. If the evaluation
  sheet is outside the cuboid, selection produces the ordinary formula `#N/A`,
  which the selected function handles according to its normal formula-error
  rule.
* In scalar context, an inline Array supplied to an ordinary scalar predicate
  or conversion uses `[0,0]` under §3.3. In matrix context, that same Array is
  lifted elementwise under the preceding rule. `N` is an explicit scalar-result
  exception: it selects `[0,0]` and broadcasts one result, even inside
  projected matrix evaluation. `TYPE` is the other scalar-result exception:
  it classifies the whole Array as `64` and broadcasts that one result. A
  one-row or one-column input broadcasts according to the ordinary matrix
  shape rules. Out-of-range coordinates produce the usual formula `#N/A`/shape
  error.
* `TYPE` and `N` are scalar-result operations even when they receive an Array or
  Reference. `TYPE` classifies the aggregate as described above. `N` follows
  its explicit “dereferenced to a scalar” rule, so a Reference uses the
  caller-supplied scalar intersection, including the current-sheet selection
  for a 3-D Reference, and does not produce one N result per cell. A projected
  caller may evaluate that scalar at each output coordinate only when it
  explicitly supplies that coordinate; the function itself never silently
  caches one coordinate as another. Both functions broadcast their one scalar
  result only through an enclosing matrix operator.
* `ERROR.TYPE` uses the `Error` pseudotype. In scalar context an inline Array or
  multi-cell Reference is first reduced by the ordinary scalar argument rule;
  in matrix context it lifts elementwise. A non-Error element is the formula
  `#VALUE!` result for that element.
* `NUMBERVALUE` accepts one Text argument and zero, one, or two Text separator
  arguments. Its source text and supplied separators are aligned by the normal
  matrix rules; an omitted separator is a scalar profile default. A
  ReferenceList is not a Text value and is rejected before selection. A
  rectangular Reference to Text is admitted under the matrix/scalar rules
  above, while a known non-Text scalar follows the explicit conversion choice
  below. A known-invalid separator pair is a formula `#VALUE!` in scalar or
  projected context. In matrix context it preserves the broadcast shape from
  argument descriptors and fills that shape with `#VALUE!` elements before any
  source-cell reads; for example, `NUMBERVALUE(A1:A3;"..";",")` is a 3-by-1
  error array with zero reads from `A1:A3`.
* `VALUE` has one Text argument and follows the same scalar/matrix selection
  rules. It accepts a direct Reference to Text and the profile-defined Empty
  behavior; a ReferenceList is rejected before resolver reads.
* `NA()` has no argument and always has scalar shape.

An Array-returning child keeps its own matrix contract when it is consumed by
an `Any` argument. In particular, `MUNIT` (§6.5.5) returns an Array and takes
an `Integer`; under §3.3.2.2.1, a direct `MUNIT` evaluated in matrix context
uses the `[0,0]` input element when its Integer argument is non-scalar. This
includes a direct `MUNIT` branch under a projected `IF`, and it remains true
when `TYPE(Any)` classifies the resulting Array. `TYPE` must not turn its `Any`
child into a scalar demand merely because its own result is scalar. A genuinely
position-sensitive scalar MUNIT parameter is expressed as
`MUNIT(N([.G1:.G2]))`: `N` explicitly dereferences its reference to one scalar
at the caller's current position, after which MUNIT receives that scalar. A
conditional criterion that explicitly requests scalar demand uses the same
current-position rule. Argument scheduling must derive this distinction from
the expected pseudotype and evaluation context, rather than adding a
TYPE-specific or MUNIT-specific projection exception.

An inline or referenced Array containing formula Error values is still an
Array for `TYPE`; the outer result is `64`. A single-cell Reference containing
an Error is first dereferenced and is `16`. Formula Error values in the
elementwise predicate/conversion functions are handled by each function's
error column above, not hidden by a generic matrix mapper.

`ReferenceList` is an ordered union of references, not an Array. It cannot be
implicitly converted to a scalar or Array for this batch. A known list shape
passed to a `Scalar`, `Number`, `Text`, or `Any` operation that needs one value
therefore produces formula `#VALUE!` with zero resolver reads. The expression
that computes a reference descriptor may have already performed its own reads;
the zero-read guarantee begins once the resulting descriptor is known to be a
refusal.

This distinction follows the normative model: §4.8 defines a Reference as a
cuboid, while §§3.3, 6.3.2, and 6.3.3 define scalar implicit intersection.
Those rules do not authorize rejecting a cuboid solely because it spans sheets.
The current-sheet plane choice for a matrix projection is an explicit profile
decision that keeps the 2-D result shape and preserves the existing scalar
intersection behavior. `TYPE` has the separate §6.13.33 rule that a Reference
is dereferenced and its formulas evaluated, so its complete multi-sheet cuboid
is scanned and then classified as Array `64`.

## Parity functions

`ISEVEN` and `ISODD` first apply `TRUNC` toward zero and then test the exact
represented integer's remainder modulo two. The selected profile accepts a
Logical by Number conversion (`TRUE` is odd and `FALSE` is even); this is the
implementation-defined choice allowed by Part 4. Number conversion follows
§6.3.5 in this profile: Number is unchanged, Logical is `1`/`0`, and Empty
(including an empty referenced cell) is `0`. Therefore
`ISEVEN(Empty)=TRUE` and `ISODD(Empty)=FALSE`. Finite numeric Text is parsed by
the fixed numeric bridge; malformed Text produces `#VALUE!`, while a
syntactically numeric non-finite result produces `#NUM!`. Complex and other
non-real values produce `#VALUE!`; formula Errors propagate.

The parity operation must not use an arbitrary `2^53` cutoff or cast through a
narrow integer. It handles every finite binary64 value that can reach the
evaluator: for large integral values whose binary64 spacing is at least two,
the represented value is necessarily even; fractional values are truncated
before this observation. Inputs such as `0`, `-0`, `1.9`, `-1.9`, `2^53`,
`2^53+2`, `f64::MAX`, and a numeric Text that overflows the finite profile are
separate contract cases. A syntactically valid numeric conversion that yields
non-finite or out-of-domain state returns `#NUM!`; malformed text returns
`#VALUE!`.

## NUMBERVALUE

`NUMBERVALUE` is a locale-independent text transform followed by XML Schema
`xsd:float` parsing. It is not a generic locale parser and it does not parse
dates or times.

When the optional separators are omitted, the selected explicit profile uses
DecimalSeparator `.` and GroupSeparator `,`. An omitted or missing optional
slot uses that profile default; an explicitly supplied empty Text is still a
supplied separator and is validated. In particular, an explicit empty
DecimalSeparator is invalid (`LEN` is zero), while an explicit empty
GroupSeparator is valid. When separators are supplied:

* DecimalSeparator must have exactly one text character (`LEN` one), and it
  must not occur anywhere in GroupSeparator.
* GroupSeparator may be empty. Separator validation happens before scanning the
  source text and a violation returns `#VALUE!` without a resolver read for a
  known source descriptor. Thus `NUMBERVALUE(text;"";",")` is invalid, while
  `NUMBERVALUE(text;".";"")` uses no grouping character.
* The source text is transformed in this order:
  1. remove every GroupSeparator occurrence before the first DecimalSeparator;
  2. replace the first DecimalSeparator occurrence with U+002E FULL STOP;
  3. remove all whitespace characters defined by Part 4 §5.14;
  4. prepend `0` when the transformed string begins with FULL STOP;
  5. remove all trailing U+0025 PERCENT SIGN characters.
* Divide the parsed result by `100` once for each percent sign removed. A
  percent sign in any other position remains invalid.
* The final string must be a valid `xsd:float` lexical form. Malformed syntax,
  invalid grouping left in the final form, a non-Text non-Empty input, and an
  Empty input return `#VALUE!` in this profile. A finite parse is returned as
  Number. A syntactically accepted non-finite result or a finite result outside
  the repository's Number domain returns `#NUM!`; NaN and infinities are never
  published as Number values.

The source Text is charged and parsed from its borrowed representation. The
implementation may use bounded parser scratch, but must not clone a complete
reference or silently route through the ordinary Text-to-Number conversion,
because that conversion has different malformed/non-finite behavior.

## VALUE numeric, fractional, time, date, and datetime forms

`VALUE` accepts the required locale-independent and `en_US` forms. It must not
be reduced to the ordinary decimal parser.

### Locale-independent numeric and fractional forms

Regardless of locale, accept the ASCII grammar

```text
[+-]?[0-9]+([eE][+-]?[0-9]+)?%?
```

and divide by `100` for one trailing percent sign. This form has no decimal
point; a decimal point is handled by the locale form. A valid finite result is
returned as Number. A valid lexical form that produces a non-finite or
out-of-domain value is `#NUM!`; malformed text is `#VALUE!`.

Also accept the fractional grammar

```text
[+-]?[0-9]+ [0-9]+/[1-9][0-9]?
```

where the displayed space is required, the denominator is nonzero and has one
or two digits, and the leading sign applies to the complete mixed number. `0`
is allowed as the integer part for values between zero and one. The fraction
must be finite; a domain overflow is `#NUM!`.

### `en_US` decimal and currency forms

The selected profile accepts the required `en_US` grammar:

```text
[+-]?\$?([0-9]+(,[0-9]{3})*)?(\.[0-9]+)?(([eE][+-]?[0-9]+)|%)?
```

The leading `$` is ignored. Commas are ignored only when they form valid
thousands groups; malformed grouping is not silently stripped. One trailing
percent sign divides by `100`. The locale decimal point is `.` and the locale
group separator is `,`. The parser must not accept an empty currency-only or
sign-only string as zero.

### Time and date forms

Accept times in at least `HH:MM` and `HH:MM:SS`, with one or two digits for
each component, `HH` in `0..=23`, and `MM`/`SS` in `0..=59`. A time returns the
fraction of a day (`HH/24 + MM/1440 + SS/86400`). Fractional seconds in
`HH:MM:SS.s...` are accepted and added to the day fraction. `24:00` and leap
seconds are outside this profile. A time parse that is syntactically shaped
but outside these domains is `#VALUE!`.

Accept ISO dates in `YYYY-MM-DD` and validate them using the proleptic
Gregorian calendar. The serial is the day offset from the explicit epoch.
Accept `MM/DD/YYYY` in the `en_US` profile. Also accept all seven required
`en_US` forms:

| Form | Example | Year rule |
| --- | --- | --- |
| `MM/DD/YYYY` | `5/21/2006` | four digits |
| `MM/DD/YY` | `5/21/06` | explicit `table:null-year` |
| `MM-DD-YYYY` | `5-21-2006` | four digits |
| `mmm DD, YYYY` | `Oct 29, 2006` | English short month |
| `DD mmm YYYY` | `29 Oct 2006` | English short month |
| `mmmmm DD, YYYY` | `October 29, 2006` | English full month |
| `DD mmmmm YYYY` | `29 October 2006` | English full month |

Month names follow the explicit locale profile and are matched
case-insensitively in the selected English profile. Numeric month/day fields
may use one or two digits where the corresponding required form permits it;
the year fields marked `YYYY` have four digits and `YY` has two. Dates are
validated for month length and Gregorian leap years. A malformed date or
impossible day/month is `#VALUE!`; a valid date outside the selected supported
serial domain is `#NUM!`.

Accept a datetime as a date followed by a time, separated by one space or the
literal `T`. The ISO forms `YYYY-MM-DD HH:MM` and
`YYYY-MM-DDTHH:MM:SS` are mandatory. The selected profile also permits the
listed `en_US` date forms before the same separator. The result is the date
serial plus the time fraction. The date and time are both retained; this is
not `DATEVALUE` or `TIMEVALUE` behavior that discards the other component.

The date parser has no timezone or current-date dependency. The supported
domain is 1899-12-30 through 9999-12-31 inclusive in the default profile,
which covers Part 4's required 1904-01-01 through 9999-12-31 range and its
recommended lower epoch. A valid date/time whose serial cannot be represented
as a finite Number is `#NUM!`. Other additional formats are optional only
when they do not conflict with these required forms.

For `VALUE` given neither Text nor a Reference to Text, Part 4 leaves the
result implementation-defined. This profile chooses numeric Empty/reference
Empty as `0`, because an explicitly dereferenced empty cell retains its Empty
identity and the evaluator's scalar numeric empty conversion is zero. Empty
Text (`""`), Logical, Complex, Array, and ReferenceList are not Text for this
purpose: Empty Text and unsupported non-Text scalar values return `#VALUE!`,
and a known ReferenceList returns `#VALUE!` before reads. A formula Error
propagates.

## Error and evaluation precedence

Function arguments follow the normal eager evaluation rule in §3.2.3. The
inspection functions then decide whether a resulting formula Error is data or
an error to propagate:

* `ERROR.TYPE`, `ISBLANK`, `ISERR`, `ISERROR`, `ISLOGICAL`, `ISNA`,
  `ISNONTEXT`, `ISNUMBER`, `ISTEXT`, and `TYPE` inspect formula Errors as
  values and return their specified Boolean/type result.
* `N`, `ISEVEN`, `ISODD`, `NUMBERVALUE`, and `VALUE` propagate a formula Error
  supplied by an argument. They may generate a new formula `#VALUE!` or
  `#NUM!` for a non-error value that cannot satisfy their conversion/domain.
* `NA` creates `#N/A` itself and takes no argument.

When a reference scan has already retained a formula Error value, the scan
continues whenever the function must inspect the complete descriptor (notably
`TYPE` on an admitted rectangular Reference). A later typed resolver failure,
resource limit, cancellation, or source-version failure supersedes any retained
formula Error and bubbles as `EvaluationFailure`. `IFERROR`, `ISERROR`, and the
inspection functions catch only formula Error values; they do not catch or
rewrite typed failures.

## Bounded reference and cache requirements

The value implementation must preserve the existing explicit execution
context, source-version fences, cancellation checks, and resource budgets:

* Known pseudotype or shape refusals, including ReferenceList input to a
  single-value operation, are decided before resolver selection and perform
  zero resolver reads. A computed expression may have read while producing its
  descriptor; the guarantee begins once the descriptor is known.
* Admitted references are visited in sheet/row/column order with work charged
  before each cell, read limits and cancellation checked before and after the
  provider call, and successful-read accounting performed only after a read
  succeeds. The mapper retains borrowed Text and does not materialize a whole
  range solely to run a predicate or parser.
* `TYPE` on a rectangular Reference is the intentional exception to
  metadata-only classification: Part 4 requires referenced formulas to be
  evaluated. This includes every sheet plane in a multi-sheet cuboid, visited
  in sheet/row/column order. The scan may discard each completed cell after
  checking for typed failure and formula evaluation; it returns `64` after the
  complete scan. A scalar predicate or `N` reads only the current-sheet
  intersection cell of such a cuboid. `TYPE` on a known ReferenceList is a
  refusal and does no scan.
* Numeric and text parser scratch, output arrays, and shape metadata use checked
  limits. Text-byte charging occurs before parsing; text and references are not
  cloned merely to preserve a scalar result.
* Source and cancellation fences surround evaluation and publication. A cached
  formula result is not a provider refresh and cannot bypass those fences.

Purity does not imply position invariance. Literal-only results and `NA()` may
use scalar payload/error cache entries. A function that receives an implicit
intersection, a projected Reference, or a computed descriptor is position
sensitive unless its complete descriptor and coordinate context are present in
the cache key. `TYPE` may use a complete, shape-preserving descriptor cache
only after the required reference evaluation and typed-failure checks; a
coordinate-only scalar cache is insufficient. `N`, predicates, `VALUE`, and
`NUMBERVALUE` must not reuse one selected cell for another projected output.

## Required semantic gates

Before production completion, the implementation and focused tests must pin at
least the following cases:

1. Every function in the signature table with zero, valid, extra, and missing
   arguments, including exact zero-argument `NA` and the formula result of
   `NA()`.
2. Empty cell, empty Text, Number subtypes, Logical, Complex, every standard
   formula Error, and Text that merely resembles an error spelling.
3. `TYPE` on each scalar kind, an Error, Empty, a single-cell Reference, a
   rectangular Reference containing Errors, an inline Array, and a known
   ReferenceList. Verify that a rectangular Reference is scanned for typed
   failures but is classified as `64`.
4. `N` for Number, Logical, Text, Empty, Complex, Error, inline Array, a
   single-cell Reference, and a multi-cell Reference under an explicit
   intersection coordinate.
5. `ISEVEN`/`ISODD` for negative and fractional values, Logical values, zero,
   `2^53` and larger exactly represented values, `f64::MAX`, malformed text,
   overflow/non-finite text, Complex, and formula Errors.
6. `NUMBERVALUE` with omitted and explicit separators, invalid separator
   constraints, grouping before/after the decimal, all §5.14 whitespace,
   leading decimal, repeated percent signs, malformed text, non-finite text,
   Empty, formula Error, and ReferenceList refusal.
7. `VALUE` for the locale-independent integer/scientific grammar, `en_US`
   decimal/grouping/currency/percent forms, mixed fractions, all required time
   forms including fractional seconds, ISO dates, all seven `en_US` date forms,
   two-digit years under the fixed 1930 profile, ISO datetimes with both
   separators, invalid dates/times, out-of-domain dates, and formula Errors.
   Document-specific `null-year` overrides remain future explicit-context
   scope and are not claimed as a current evaluator option.
8. Matrix lifting and broadcast shape for the elementwise predicates and
   conversions, scalar-result behavior for `TYPE` and `N`, projected lazy
   branches with two differing referenced cells, and cache hits only when the
   complete descriptor/coordinate context is invariant.
9. Resource limits, cancellation, source changes, typed resolver failures after
   earlier formula Errors, borrowed text, and the zero-read shape/type refusal
   cases. A typed failure must remain observable through `ISERROR`,
   `ERROR.TYPE`, `IFERROR`, and every new dispatch path.

No production/test edit or validation PASS is implied by this contract alone.
