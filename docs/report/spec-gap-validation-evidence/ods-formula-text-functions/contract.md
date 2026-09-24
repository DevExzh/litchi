# ODF 1.4 text-function evaluator contract

This contract defines the current bounded ordinary text-function surface in
OpenFormula 1.4 Part 4. It covers the 26 functions in §6.20. The seven
normative byte-position functions in §6.7 are retained below as an explicit
future gap, rather than as part of this implementation batch. This contract
is the semantic boundary for
the resolver-free scalar evaluator, the resolver-backed value evaluator, and
their independent validation. The committed source baseline predates the
candidate §6.20 text dispatch; the candidate tree under review adds scalar
dispatch and value-mode text mapping. This contract does not claim that the
candidate evaluator also recalculates formula cells, imports cached formula
results, refreshes an external provider, or publishes a changed cell.

The normative source is the repository-local ODF distribution:

| Source | SHA-256 |
| --- | --- |
| archive `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| member `part4-formula/OpenDocument-v1.4-os-part4-formula.html` | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |

The directly relevant sections are §§3.2.3–3.4, 3.6–3.7, 4.2–4.11,
5.5–5.7, 5.12, 6.1–6.3, and 6.20; §6.7 is inventoried as the future byte
gap. The evaluator also retains the
bounded conversion, resource, source-fence, cache, and typed-failure rules
selected by ADR 0004, ADR 0005, ADR 0006, ADR 0008, ADR 0024 and the
aggregate, matrix, and statistical contracts.

## Current normative scope

The §6.20 family contains `ASC`, `CHAR`, `CLEAN`,
`CODE`, `CONCATENATE`, `DOLLAR`, `EXACT`, `FIND`, `FIXED`, `JIS`, `LEFT`,
`LEN`, `LOWER`, `MID`, `PROPER`, `REPLACE`, `REPT`, `RIGHT`, `SEARCH`,
`SUBSTITUTE`, `T`, `TEXT`, `TRIM`, `UNICHAR`, `UNICODE`, and `UPPER`.

These 26 names are the complete current scope for this batch. The registry
must not silently add host or Microsoft names such as `CONCAT`,
`DBCS`, `TEXTJOIN`, `VALUETOTEXT`, `ARRAYTOTEXT`, `TEXTBEFORE`, `TEXTAFTER`,
`TEXTSPLIT`, `USDOLLAR`, or `YEN`. Non-standard extensions require the §5.7
qualified-name rule and a separate compatibility profile.

The signatures below retain the ODF pseudotypes. A semicolon separates
formula arguments, square brackets denote optional arguments, and a default
after `=` applies only when the argument is omitted. `ByteLength` and
`BytePosition` are integer pseudotypes whose unit is selected by the explicit
byte profile below.

### Ordinary text functions (§6.20)

| Function | Part 4 signature | Return | Constraint/default |
| --- | --- | --- | --- |
| `ASC` | `ASC(Text T)` | Text | None. Full-width ASCII and the specified full-width Katakana forms are converted to half-width forms. |
| `CHAR` | `CHAR(Number N)` | Text | `N ≤ 127`; this profile applies generic integer conversion, admits resulting values 1–255, and rejects values outside that range. |
| `CLEAN` | `CLEAN(Text T)` | Text | None. Remove Unicode `Cc` and `Cn` characters. |
| `CODE` | `CODE(Text T)` | Number | The first character must exist. Values for code points ≥128 are implementation-defined by Part 4; this profile uses Unicode scalar values. |
| `CONCATENATE` | `CONCATENATE({Text T}+)` | Text | At least one argument. Arguments are concatenated in source order. |
| `DOLLAR` | `DOLLAR(Number N [; Integer D])` | Text | `D` is the decimal-place count; omitted `D` is 2 in the invariant formatter. Negative `D` rounds left of the decimal point. |
| `EXACT` | `EXACT(Text T1; Text T2)` | Logical | Always case-sensitive, independent of `HOST-CASE-SENSITIVE`. |
| `FIND` | `FIND(Text Search; Text T [; Integer Start=1])` | Number | `Start ≥ 1`; case-sensitive literal search, with no wildcard or regular-expression processing. |
| `FIXED` | `FIXED(Number N [; Integer D=2 [; Logical OmitSeparators=FALSE]])` | Text | Decimal rounding and invariant formatting; negative `D` rounds left. A true omit flag removes group separators. |
| `JIS` | `JIS(Text T)` | Text | None. Convert the exact Table 34 half-width mappings to full-width forms. |
| `LEFT` | `LEFT(Text T [; Integer Length])` | Text | `Length ≥ 0`; omitted `Length` is 1. Uses `INT(Length)` when supplied. |
| `LEN` | `LEN(Text T)` | Integer | Counts characters, not bytes. A Number is converted to Text first. |
| `LOWER` | `LOWER(Text T)` | Text | Unicode default lower-case mapping; ASCII `A`–`Z` must map to `a`–`z`. |
| `MID` | `MID(Text T; Integer Start; Integer Length)` | Text | `Start ≥ 1`, `Length ≥ 0`; `INT` conversion is explicit. |
| `PROPER` | `PROPER(Text T)` | Text | Uppercase the first letter and letters after non-letters; lowercase letters following letters. At least Latin A–Z/a–z is required. |
| `REPLACE` | `REPLACE(Text T; Number Start; Number Count; Text New)` | Text | Raw `Start ≥ 1` and `Count ≥ 0`; character positions are 1-based. Both Number arguments use generic truncation toward zero. |
| `REPT` | `REPT(Text T; Integer Count)` | Text | `Count ≥ 0`; zero returns the empty string. Part 4 prints `T(Text T; Integer Count)` in the syntax line; the summary and semantics identify the function as `REPT`, which this contract follows. |
| `RIGHT` | `RIGHT(Text T [; Integer Length])` | Text | `Length ≥ 0`; omitted `Length` is 1. Uses `INT(Length)` when supplied. |
| `SEARCH` | `SEARCH(Text Search; Text T [; Integer Start=1])` | Integer | `Start ≥ 1`; case-insensitive literal search in the base profile. |
| `SUBSTITUTE` | `SUBSTITUTE(Text T; Text Old; Text New [; Integer Which])` | Text | `Which ≥ 1` when supplied. Omitted `Which` replaces all occurrences; an empty `Old` returns `T`. |
| `T` | `T(Any X)` | Text | Return `X` only when its dereferenced value is Text; otherwise return the empty Text value. It is not a conversion function. Formula Errors still propagate. |
| `TEXT` | `TEXT(Scalar X; Text FormatCode)` | Text | `FormatCode` has an implementation-defined grammar; this profile selects the bounded invariant grammar below. |
| `TRIM` | `TRIM(Text T)` | Text | Remove leading/trailing spaces and reduce every internal run to one U+0020 space. |
| `UNICHAR` | `UNICHAR(Integer N)` | Text | `0 ≤ N ≤ 0x10FFFF`; this profile rejects surrogate code points and returns any other Unicode scalar, including U+0000. |
| `UNICODE` | `UNICODE(Text T)` | Number | `LEN(T) > 0`; return the first Unicode scalar value. |
| `UPPER` | `UPPER(Text T)` | Text | Unicode default upper-case mapping; ASCII `a`–`z` must map to `A`–`Z`. |

## Future gap: byte-position functions (§6.7)

| Function | Part 4 signature | Return | Constraint/default |
| --- | --- | --- | --- |
| `FINDB` | `FINDB(Text Search; Text T [; BytePosition Start])` | BytePosition | Same search semantics as `FIND`, using the selected byte unit; omitted `Start` is 1. |
| `LEFTB` | `LEFTB(Text T [; ByteLength Length])` | Text | Same selection as `LEFT`, using byte length; omitted `Length` is 1. |
| `LENB` | `LENB(Text T)` | ByteLength | Same text conversion as `LEN`, using byte length. |
| `MIDB` | `MIDB(Text T; BytePosition Start; ByteLength Length)` | Text | `Length ≥ 0`; same extraction rules as `MID`, using byte positions. |
| `REPLACEB` | `REPLACEB(Text T; BytePosition Start; ByteLength Len; Text New)` | Text | Same replacement semantics as `REPLACE`, using byte positions. |
| `RIGHTB` | `RIGHTB(Text T [; ByteLength Length])` | Text | Same selection as `RIGHT`, using byte length; omitted `Length` is 1. |
| `SEARCHB` | `SEARCHB(Text Search; Text T [; BytePosition Start])` | BytePosition | Same search semantics as `SEARCH`, using the selected byte unit; omitted `Start` is 1. |

The `REPT` syntax typo is a specification editorial inconsistency, not a
second function or an alias. The seven §6.7 signatures above are retained for
future implementation planning only. The UTF-8 boundary text below is a
provisional proposal, not a current dispatch, validation, or cache-identity
claim; the later byte-function batch may adopt a different explicit profile.
When that gap is implemented, `REPLACEB` and `SEARCHB` inherit the ordinary
function's `Start`/`Which` domain rules even where the short §6.7 entry only
says “as REPLACE” or “as SEARCH”.

Required arguments have strict arity. A zero-argument call to a function with
a required parameter, or a `Missing` value in a required slot, is a generated
`#VALUE!`; it is not silently replaced by an empty Text. An omitted optional
slot, including the evaluator's `Missing` marker for an optional AST position,
receives that function's documented default. A supplied non-Missing value is
converted and checked even when it is false, zero, or empty Text.

## Common scalar conversion and error profile

### Text, Number, Integer, and Any

The selected deterministic conversion profile is:

* Text is retained without normalization. A finite Number is rendered with a
  shortest round-trip decimal representation, no whitespace, and a period as
  the decimal separator. Logical values render as the uppercase Text `TRUE`
  or `FALSE`. An Empty scalar or an Empty referenced cell converts to the
  empty Text value. A Reference to more than one cell uses the ordinary
  implied-intersection rule unless the enclosing matrix evaluation iterates
  it. A Formula Error propagates unchanged. Complex values are not silently
  rendered as ordinary Text; they produce `#VALUE!` unless their producer
  already supplied a Text value.
* Number conversion accepts finite real Numbers and Logical `FALSE`/`TRUE` as
  0/1. Text conversion to Number uses only a locale-independent decimal
  grammar in this profile; malformed Text produces `#VALUE!` and non-finite
  parsed Text produces `#NUM!`. Although Part 4's `Scalar` pseudotype includes
  complex Numbers, this ordinary text/number profile does not format or
  numerically coerce Complex values; they produce `#VALUE!`. A reference to
  an Empty cell converts to 0 under §6.3.5.
* Generic Integer conversion first applies the Number conversion and then
  truncates toward zero. The functions that explicitly say `INT` (`LEFT`,
  `RIGHT`, and `MID`) use floor toward negative infinity, as `INT` requires;
  they do not inherit generic truncation. This is the repository choice for
  the otherwise implementation-defined generic Number-to-Integer conversion.
* `Any` performs no conversion. `T` receives its argument after the ordinary
  §3.3 non-scalar rule: scalar-mode inline Arrays first project `[0,0]`, and
  multi-cell References use implied intersection. It returns the projected
  value only when that value has Text type. A projected Number, Logical,
  Complex, Empty, or other non-Text value yields empty Text; a Formula Error
  propagates under the common error rule. Matrix mode applies the same test
  per output position. A known ReferenceList is not converted to an Array and
  is refused with `#VALUE!` before reading cells from that descriptor; it is
  not a scalar `Any` value merely because `T` declares `Any`.

Although `T` declares `Any`, this profile treats a matrix operand according to
the §3.3 scalar-expected rule: scalar mode uses the `[0,0]` projection and
matrix mode lifts the `T` test independently at each output position. The
`Any` declaration therefore does not make a produced Array opaque to matrix
iteration.

This profile intentionally does not use the locale-dependent conversion
permission in §6.3.5 for ordinary text functions. The `HOST-LOCALE` property
may be exposed by a future explicit profile, but no ambient process locale is
consulted.

### Argument evaluation and matrix mode

Text functions are ordinary scalar-result functions; none of their signatures
contains `ForceArray` or a sequence pseudotype. In scalar mode, a non-scalar
inline Array uses its `[0,0]` element and a Reference uses implied
intersection, as §3.3 requires. A known ReferenceList passed where a scalar
Text, Number, Integer, Logical, or `Scalar` is required is a pseudotype
failure: the value profile generates `#VALUE!` before reading cells from the
refused descriptor. Eager evaluation of other admitted arguments may still
read their own references; the per-argument refusal is not a whole-call read
guarantee.

In matrix mode, non-scalar inputs to scalar parameters iterate under §3.3:
the result is rectangular, singleton/scalar arguments broadcast, a one-row or
one-column input broadcasts along the other dimension, and an out-of-range
two-dimensional input contributes `#N/A`. This applies independently to
`FIND`/`SEARCH` queries and starts, `SUBSTITUTE`'s `Which`, formatting
arguments, and every argument of `CONCATENATE`. A function that returns one
Text/Number/Logical value does not become an Array-producing function merely
because it is evaluated in matrix mode.

The §3.3.2.2.1 exception applies to an array-returning producer's own scalar
inputs. For example, `MUNIT({2;3})` selects its `[0,0]` size input and returns
one Array. If that already-produced Array is later consumed by a scalar text
function, the consumer may matrix-lift over the produced values. The producer
input rule must not collapse an Array returned for later consumption.

Functions and operators are evaluated eagerly unless their specification says
otherwise. A Formula Error is propagated, with the leftmost supplied error
preferred when several are present. The text function `T` does not suppress
Formula Errors merely because its parameter is `Any`; its “non-Text gives
empty Text” rule applies to ordinary non-error values. A typed
  `Unsupported`, `ResourceLimit`, `Allocation`, `Cancelled`, `SourceChanged`,
  or `SourceVersionAvailabilityChanged` failure remains an `EvaluationFailure`
  and is not converted to a formula Error or caught by `IFERROR`/`IFNA`.

The selected generated formula subtype is `#VALUE!` for wrong type, invalid
constraint, missing text, empty search result, invalid code point, invalid
format code, and invalid optional argument. `#NUM!` is used for a non-finite
Number or a non-finite intermediate/final numeric conversion. Resource and
formatter-capacity failures remain typed failures. The common §6.1 error rule outranks
generated constraint errors; a later typed failure outranks a retained formula
error. Errors in an already-admitted input are never converted to empty Text.

## Unicode and character semantics

ODF Text is a Character string, positions start at 1, and normalization of
formulas, inputs, and results is implementation-defined. The repository base
profile pins Unicode 17.0 data and default algorithms in its capability
identity; a table update is a new profile and a new cache identity. It makes
these choices:

* A “character” is one Unicode scalar value, not a UTF-8 byte, UTF-16 code
  unit, grapheme cluster, or user-perceived glyph. `LEN`, `LEFT`, `RIGHT`,
  `MID`, `REPLACE`, `FIND`, `SEARCH`, and `SUBSTITUTE` therefore count and
  return by Unicode scalar positions. Combining marks remain separate
  positions; no grapheme segmentation is attempted.
* No Unicode normalization is applied. Canonically equivalent sequences with
  different code-point sequences remain different for `EXACT`, `FIND`,
  `SEARCH`, and `SUBSTITUTE`, and the first code point returned by `UNICODE`
  depends on the supplied sequence. This is the selected normalization
  profile permitted by §4.2.
* `LOWER` and `UPPER` use the Unicode 17.0 default case algorithms and default
  case mappings. Context-sensitive rules, including final-sigma behavior, use
  the surrounding original scalar sequence; locale-specific Turkish-style
  mappings are not selected without an explicit host profile. Mappings that
  expand one scalar into multiple scalars are retained and charged as output
  bytes.
* `PROPER` uses Unicode General Category `L*` for “letter” boundaries in this
  profile, with at least the required Latin behavior. The first letter or a
  letter following a non-letter is uppercased; a letter following a letter is
  lowercased. Non-letters are copied unchanged.
* `CLEAN` removes exactly Unicode General Categories `Cc` (Other, Control)
  and `Cn` (Other, Not Assigned). It keeps printable characters, spaces,
  format characters (`Cf`), separators, and all other categories in their
  original order. The ordinary U+0020 space is printable.
* `TRIM` recognizes only U+0009 HORIZONTAL TABULATION, U+000A LINE FEED,
  U+000D CARRIAGE RETURN, and U+0020 SPACE as spaces. It removes runs at both
  ends and replaces every internal run of two or more recognized spaces with
  one U+0020. Non-breaking space and other Unicode whitespace are retained.

### ASC and JIS

`ASC` and `JIS` are complementary Unicode mapping functions. They must
implement Tables 33 and 34 of Part 4, including their special exceptions,
rather than approximating the mapping by subtracting or adding `0xFEE0` for
every character:

* `ASC` maps U+FF01–U+FF5E full-width ASCII to U+0021–U+007E; maps the
  Katakana ranges and voiced/semi-voiced marks exactly as Table 33 specifies;
  maps U+2015, U+2018, U+2019, U+201D, U+3001, U+3002, U+300C, U+300D,
  U+309B, U+309C, U+30FB, U+30FC, and U+FFE5 through the table's special
  mappings; and copies all other characters.
* `JIS` maps U+0021–U+007E to U+FF01–U+FF5E, applies the Table 34 quotation,
  apostrophe, grave-accent, reverse-solidus/Yen, and Katakana exceptions, and
  copies all other characters. A half-width voiced mark is consumed with the
  preceding half-width Katakana where Table 34 requires it.

The reverse relationship is a mapping property, not a promise that every
arbitrary Unicode string round-trips byte-for-byte: characters with no table
mapping remain unchanged, and multiple source spellings can map to one target.
The reverse-solidus/Yen behavior follows the table's code-page-932 legacy
exception while the rest of the profile remains Unicode-based.

### CHAR, CODE, UNICHAR, and UNICODE

`CHAR` and `CODE` have deliberately host-dependent wording in Part 4. The
selected profile is Unicode and does not use an ambient system code page:

* `CHAR` applies generic Number-to-Integer conversion (finite fractional
  inputs truncate toward zero), requires the resulting integer to be in
  1–255, and returns the corresponding Unicode scalar U+0001–U+00FF. This
  extends the required ≤127 range through the permitted 255 range and makes
  `CODE(CHAR(N)) = N` for every 1 ≤ N ≤ 255. Zero, negative, and
  greater-than-255 converted values are `#VALUE!`; a non-finite input is
  `#NUM!` under the shared numeric-error profile.
* `CODE` requires non-empty Text and returns the Unicode scalar value of its
  first scalar. It therefore returns values above 127 according to the
  selected Unicode profile, rather than a platform code-page byte. Empty Text
  is `#VALUE!`.
* `UNICHAR` accepts every Unicode scalar value from U+0000 through U+10FFFF,
  excluding the surrogate range U+D800–U+DFFF. U+0000 is a valid generated
  Text character even though a source formula literal may reject an embedded
  NUL. Out-of-range and surrogate values are `#VALUE!`.
* `UNICODE` requires non-empty Text and returns the first Unicode scalar. It
  does not return a UTF-16 code unit for an astral character.

## Search, slicing, replacement, and repetition

The ordinary character-position operations use 1-based Unicode scalar
positions. `Start` and `Length`/`Count` conversions are checked before any
slice is made; a negative or zero-forbidden parameter is `#VALUE!`.

* `LEFT(T; Length)` returns the first `Length` scalars, or all of `T` when it
  is shorter. Omitted Length is 1. Length zero returns empty Text.
* `RIGHT(T; Length)` returns the last `Length` scalars with the same defaults
  and zero behavior.
* `MID(T; Start; Length)` returns at most Length scalars beginning at Start.
  Start is 1-based and must be at least 1. Start beyond `LEN(T)` returns
  empty Text; Length zero returns empty Text.
* `REPLACE(T; Start; Count; New)` removes Count scalars at Start and inserts
  New. Count zero inserts before the selected position. Part 4 contains an
  off-by-one contradiction here: its prose says to clamp `Start > TLen` to
  `TLen` and `Count > TLen−Start` to that remainder, while its displayed
  equivalence is `LEFT(T;Start−1)&New&MID(T;Start+Count;LEN(T))` and therefore
  replaces the final character in the ordinary one-based way. The base profile
  resolves this in favor of the displayed equivalence and the stated character
  positions. First validate finite raw `Start ≥ 1` and `Count ≥ 0`; then apply
  generic toward-zero conversion, using `trunc₀` for both values. The displayed
  `LEFT`/`MID` identity is applied only after this integerization, never to raw
  fractional Number arguments. For `TLen = 0` use `S = 1`; otherwise use
  `S = min(max(trunc₀(Start), 1), TLen + 1)`, then
  `C = min(trunc₀(Count), TLen − S + 1)`, and return
  `LEFT(T;S−1)&New&MID(T;S+C;LEN(T))`. Thus `S = TLen+1` appends, while
  `S = TLen` can replace the final scalar. A non-finite argument is `#NUM!`
  and a finite raw-domain violation is `#VALUE!`. This selected boundary rule
  must be used by `REPLACEB` after its byte-to-scalar boundary mapping.
* `REPT(T; Count)` repeats the complete Text Count times. Count zero or an
  empty T returns empty Text. Checked multiplication happens before
  allocation; an output beyond the caller's text/storage budget is a typed
  resource failure, not a truncated string.
* `FIND(Search; T; Start)` searches literally and case-sensitively from the
  scalar position Start. `SEARCH` uses the same position rule and literal
  matching but applies the profile's Unicode default full case-fold mapping.
  Case-fold expansions (for example `ß` to `ss` and a ligature to its expanded
  sequence) are compared without normalizing either operand. A candidate starts
  only at an original source-scalar boundary and succeeds only when the folded
  match also ends at an original source-scalar boundary; the reported position
  is the original source position, not a folded-sequence offset. Both return
  the first 1-based position or `#VALUE!` when no match exists.
  An empty Search is considered to match at Start when Start is within
  1..LEN(T)+1; this is the selected empty-needle profile. Start outside that
  range is `#VALUE!`.
* `SUBSTITUTE(T; Old; New; Which)` searches from the left using exact,
  case-sensitive, normalization-free scalar sequences. With no Which it
  replaces every non-overlapping occurrence. With Which it replaces only the
  requested positive occurrence. An empty Old is explicitly a no-op and
  returns T. No match also returns T.

The profile's case and literal-search choices are independent of
`HOST-CASE-SENSITIVE`, `HOST-USE-REGULAR-EXPRESSIONS`, and
`HOST-USE-WILDCARDS`. Part 3.4 permits those host properties to alter search
behavior; enabling them requires a separate compatibility profile with a
specified regex dialect, wildcard escaping, invalid-pattern error, and
precedence when both properties are true. The base evaluator never consults
an ambient host setting.

## Provisional future byte-position profile (§6.7, not current scope)

ODF intentionally leaves `ByteLength` and `BytePosition` implementation-
dependent. This draft proposes UTF-8 bytes of the Unicode Text value. A
byte position is 1-based; `LENB` counts `T.as_bytes().len()`, and a successful
`FINDB`/`SEARCHB` returns the one-based UTF-8 offset of the first matching
scalar boundary. This is a profile choice, not a claim that all ODF hosts use
UTF-8 or that byte functions are portable between hosts. It does not bind a
future byte-function cache identity.

Byte functions use complete Unicode scalars and never publish malformed UTF-8:

* `LEFTB` and `RIGHTB` return the longest prefix or suffix whose complete
  scalars fit within the requested UTF-8 byte length.
* `MIDB` starts at the scalar containing the requested byte position (a
  position inside a multi-byte encoding is snapped to that scalar's first
  byte) and takes complete scalars until the byte length is exhausted. A
  start after the final byte returns empty Text; Length zero returns empty
  Text.
* `REPLACEB` uses the same scalar-boundary snapping and complete-scalar
  removal as `MIDB`, then inserts New. Its out-of-range Start/Count handling
  follows the selected `REPLACE` clamp profile.
* `FINDB` and `SEARCHB` search only at Unicode scalar boundaries, return the
  UTF-8 byte position of the match, and otherwise follow `FIND` and `SEARCH`
  case/literal rules. Start may equal total byte length plus one for an empty
  needle; other out-of-range starts are `#VALUE!`.

The old host-style “one byte for Latin-1 and two bytes for every other
character” DBCS approximation is not this proposal. A future code-page byte
profile must be explicit and must define its own cache identity, fixtures, and
interoperability declaration.

## Formatting functions and invariant profile

Part 4 leaves the format-code grammar implementation-defined and describes
`DOLLAR`/`FIXED` in locale terms. This repository selects a built-in invariant
formatter so these functions do not depend on a process locale or other
external state.

* `DOLLAR` uses `$` as a prefix, `,` as the grouping separator, `.` as the
  decimal separator, and two decimal places when `D` is omitted. A negative
  result is rendered in parentheses around the currency value, for example
  `($1.00)`; zero uses no negative marker. `D` is a generic Integer and
  negative `D` rounds left of the decimal point.
* `FIXED` uses `,` grouping and `.` decimals, defaults `D` to 2, and defaults
  `OmitSeparators` to FALSE. A true omit flag removes grouping separators only.
* Both functions use decimal half-away-from-zero rounding, including ties and
  negative `D`; finite results that cannot be represented within the text and
  work budgets fail with a typed resource error. The formatter canonicalizes
  negative zero to the non-negative zero spelling.
* `TEXT` accepts the bounded invariant format grammar consisting of up to four
  sections separated by semicolons; `0`, `#`, and `?` digit placeholders;
  decimal and grouping marks; percent; `E`/`e` scientific notation; simple
  slash fractions; Gregorian date/time fields (`y`, `m`, `d`, `h`, `s`, and
  `AM/PM`); quoted, backslash-escaped, or underscore-padded literals; and `@`
  for Text. A section with literals only is valid for a Number, including a
  selected zero-value section, and is emitted without an implicit sign. For a
  Number, section selection uses only the first three branches: the first for
  non-negative values, the second for negative values when present, and the
  third for zero when three or four sections are present. The fourth section
  is never selected for a Number. If the selected numeric branch contains
  `@` without numeric or date tokens, `@` receives the invariant general
  Number text in that selected branch; it does not redirect to the fourth
  Text branch. Text and Logical values use the fourth branch only when four
  sections are present, otherwise the first branch, and `@` receives their
  Text spelling. A single numeric section supplies the ordinary minus sign
  automatically; a multi-section negative branch supplies its own sign or
  deliberately omits it.

  In numeric sections, `0` is a required digit, `#` is an omitted
  insignificant digit, and `?` is an insignificant digit rendered as one
  U+0020 space when absent so integer, fractional, and exponent positions
  remain aligned. Placeholder slots are positional from the decimal point
  outward; the renderer must not move all required zeroes ahead of question
  slots. Therefore `TEXT(1;"0?0")` is `"0 1"`, and
  `TEXT(1;"0.?0")` is `"1. 0"`. A decimal point remains when the pattern has fractional
  required or question-mark positions; an all-`#` optional fractional tail
  may omit the point with its absent digits. A simple fraction has one slash,
  no decimal or exponent token, and one through six denominator placeholder
  positions. If `d` denominator positions are present, the selected rational
  has denominator at most `10^d−1`; the denominator is never `10^d`. A
  missing denominator or more than six positions is a grammar error. Without
  an explicit space before the slash-side numerator, the pre-slash slots are
  an improper fraction: `TEXT(1.25;"#/#")` is `"5/4"`. A literal space
  separating whole-number slots from numerator slots selects a mixed fraction:
  `TEXT(1.25;"# ?/?")` is `"1 1/4"`; the separator is retained when the
  optional whole part is absent. `0`, `#`, and `?` retain their positional
  rational meaning as well: required zeroes pad, `#` omits insignificant
  positions, and `?` emits spaces for absent numerator or denominator digits.
  A finite numeric value is rounded to the nearest represented-binary64
  rational within the denominator bound; exact-distance ties select the lower
  denominator and then the lower numerator. The selected mixed whole part and
  the selected improper numerator are checked unsigned 64-bit integers: the
  whole part must be at most `2^64−1`, and an improper numerator formed as
  `whole × denominator + numerator` must also be at most `2^64−1`, using
  checked arithmetic. A finite value that needs `2^64` or a larger integer at
  either boundary is numerically unrepresentable and returns `#NUM!`; for
  example `TEXT(2^64;"#/#")` is `#NUM!`, while the represented-binary64
  predecessor `TEXT(18446744073709549568;"#/#")` is
  `"18446744073709549568/1"`. This is a numeric representability failure,
  distinct from a malformed fraction grammar. One percent marker is permitted
  with a numeric form, scales the value by 100, and is emitted in the result;
  a repeated marker is a grammar error.

  Scientific sections normalize every non-zero mantissa to `1 ≤ |M| < 10`
  after rounding, regardless of the number of integer placeholders before the
  decimal. Extra required integer positions pad that normalized mantissa with
  zeroes; `#` omits and `?` spaces an absent position. They do not request
  engineering scaling. Thus `TEXT(123;"00.0E+00")` is `"01.2E+02"`, while
  rounding a mantissa such as `9.9995` carries to `1.00` and increments the
  exponent. Exponent slots use the same `0`/`#`/`?` padding rules and the
  declared exponent sign.
  The exponent token is contiguous: `E` or `e`, an optional sign, then
  one or more `0`/`#`/`?` slots. A percent marker cannot split that token;
  `0.0E%+00` and `0E0%0` are malformed.

  An admitted percent marker retains its position among numeric slots after
  scaling. Suppressing optional zero digits or a zero rational field never
  suppresses the marker. Existing zero-rational suppression of the slash and
  mixed separator remains: `TEXT(0;"#%/#")` and `TEXT(0;"# ?%/?")`
  return `%`, while `TEXT(0;"0 ?%/?")` returns `0%`. Percent characters
  in quoted/escaped literals or `@` text sections are literal, not scaling
  markers; duplicate-marker refusal applies to numeric sections.

  Date sections use the proleptic Gregorian calendar with serial zero at
  1899-12-30. Signed serials are converted by their actual Gregorian day and
  time: negative values are accepted when the resulting rounded date lies in
  1899-01-01 through 9999-12-31, and values outside that interval are refused
  with `#NUM!`. The 1900 calendar has no fabricated leap day. Unsupported
  tokens, malformed sections, unquoted unsupported punctuation inside a
  numeric core, and format output beyond the caller's budget are respectively
  `#VALUE!` or typed resource failures; the grammar never executes formulas
  or accesses external state.

  The unquoted numeric core admits only placeholder characters, one decimal
  point, integer-side grouping commas, a percent marker, one scientific
  exponent with an optional sign, or one fraction slash in the forms above.
  Repeated decimal points, a misplaced grouping mark, `:` or another
  unrecognized punctuation character in that core is a grammar `#VALUE!`, not
  a character silently discarded by the parser. Numeric prefixes and suffixes
  outside that core are preserved as literal punctuation (so a currency symbol
  may be placed before the numeric span); alphabetic literals in numeric
  sections must be quoted or backslash-escaped. A text section containing `@`
  may also contain unquoted alphabetic literal text. Date sections additionally
  admit `-`, `/`, `:`, and spaces as date/time separators; other punctuation
  and alphabetic text must be quoted or escaped. An unrecognized alphabetic
  token is also `#VALUE!`.

  Thus `TEXT(1.2;"???.??")` is `"  1.2 "`, `TEXT(1;"???")` is
  `"  1"`, `TEXT(1;"0/")` and `TEXT(1;"0/???????")` are `#VALUE!`,
  `TEXT(12;"@;@;@;\"text:\"@")` uses the first numeric branch and returns
  `"12"`, and `TEXT(-1;"yyyy")` formats the actual Gregorian date
  `1899-12-29` rather than failing merely because the serial is negative.

`TEXT` passes its scalar X and FormatCode to this formatter without converting
a Text X through Number first. Ordinary §3.3 projection happens before the
`Scalar` type check: scalar-mode inline Arrays use `[0,0]`, and a multi-cell
Reference uses implied intersection. `Scalar` then means Number, Logical, or
Text (with Complex rejected by this ordinary formatter profile even though
Part 4 includes complex Numbers in `Scalar`). Number X uses the numeric/date
sections. Logical X is formatted as
`TRUE`/`FALSE` through a Text/literal section; a numeric or date section is
`#VALUE!`. Text X is returned through `@` (a numeric/date-only section is
`#VALUE!`). An Empty scalar after projection, a Missing argument, or a
ReferenceList that cannot be scalar-projected is not silently changed into
Number zero or empty Text for this parameter; it is a `#VALUE!` type failure.
A Formula Error propagates before formatting. The invariant formatter identity
and grammar version are part of any future cache context.

## Function-specific Unicode and locale notes

`EXACT` converts both arguments to Text and compares the resulting scalar
sequences exactly, including case and normalization form. It ignores
`HOST-CASE-SENSITIVE`, as Part 4 explicitly requires.

`LEN` converts a Number to Text before counting, including its fractional part
and the selected decimal separator. The base profile's ordinary Text bridge
uses a period; a future host profile may supply another separator only when
that profile is explicit and included in the cache identity.

`DOLLAR` and `FIXED` convert their Number and optional parameters with the
common finite Number/Integer bridge. Logical flags use FALSE/TRUE; Text flags
use the strict profile and invalid Text is `#VALUE!`. The `DOLLAR` result is
currency-formatted by the invariant formatter, not a raw arithmetic number
followed by string concatenation.

`T` does not stringify a Number. `T(5)`, `T(TRUE())`, and `T(Complex)` return
empty Text; `T("5")` returns `"5"`; `T(EMPTY_REFERENCE)` returns empty Text;
and a Formula Error remains that Formula Error. Complex is still included in
the ODF `Scalar` pseudotype; empty `T(Complex)` is the selected `Any` behavior,
whereas the ordinary formatter and Number conversion above reject Complex
with `#VALUE!`.

`TRIM` does not call Unicode `is_whitespace`; its four specified characters
are the entire recognized set. A run of one recognized space is preserved in
the interior; only runs of two or more collapse to one U+0020. `CLEAN` is likewise category-based, not a
simple ASCII control filter. These distinctions are required for non-ASCII
and control-character fixtures.

## Resource, resolver, and source boundaries

Text functions are scalar or matrix consumers, not range reducers. The value
evaluator must:

* reject a known ReferenceList or other structural scalar/type refusal before
  reading cells from that refused descriptor; check matrix shape and
  output-size arithmetic before the associated output allocation where the
  descriptor is available. Generic eager argument scheduling may already have
  evaluated sibling arguments;
* perform zero resolver cell reads for the refused known ReferenceList or other
  structural scalar/type argument. A computed expression may need to evaluate
  enough upstream work to reveal its descriptor, and eager sibling arguments
  may read their own references, but the text function does not scan a
  rejected list;
* apply scalar implied intersection in scalar mode and read only the selected
  cell. In matrix mode, read each required cell once per evaluation position
  according to §3.3 broadcasting, without evaluating an unselected lazy
  branch;
* charge AST visits, argument conversion, every inspected reference cell,
  every input/output UTF-8 byte, every Unicode scalar/mapping/search step,
  formatting work, and all matrix output positions to the caller's
  hierarchical work budget;
* check cancellation before and after resolver reads and bounded Unicode or
  formatting chunks, and enforce the cumulative reference-cell and text-byte
  limits without resetting them for matrix positions or retries;
* retain borrowed input Text for direct pass-through operations (`T`,
  `EXACT` comparison operands, no-op `SUBSTITUTE`, and unchanged slices where
  the lifetime permits), and allocate only when a transformation or owned
  result requires it;
* use fallible output reservations and checked multiplication for
  `CONCATENATE`, `REPT`, case expansion, ASC/JIS mappings, replacements, and
  format output. Exceeding the caller's text/storage ceiling is a typed
  resource failure, never silent truncation or a guessed formula error; and
* release every temporary reservation on formula, typed-failure, cancellation,
  formatter, and source-fence paths.

The default evaluator text-result ceiling is the existing finite
`max_text_bytes` limit. The ODF basic limit of 32,767 ASCII characters is a
minimum capability claim; this repository's byte ceiling is explicit and may
be lower under a caller-supplied profile. UTF-8 bytes, not Unicode scalar
count, govern storage admission.

The source version and cancellation fence is checked before text evaluation
and after all selected matrix positions and formatting work, before publishing
the result. No text function recalculates a formula-bearing cell or imports a
cached formula value. A typed source/resource failure supersedes a retained
Formula Error.

## Demand cache and publication

If demand caching is enabled for this family, §6.20 text functions are pure
for a fixed source snapshot, conversion profile, Unicode-table version, and
invariant-formatter grammar identity. Scalar
Number/Text/Logical/error payloads may then be cached only with all of those
context components and the complete argument dependency identity. This is a
future integration rule, not a claim that the current evaluator already
caches text results.

In a projected lazy branch, a text function may propagate complete arguments
only when every argument is proven coordinate-independent. A matrix argument
or a reference that is needed in full must retain its shape and source
descriptor. Position-dependent scalar criteria, including a nested MUNIT size
criterion, remain position-sensitive and are excluded from full-argument cache
propagation. An Array already produced by MUNIT can be consumed by a scalar
text function using the ordinary matrix-lifting rules.

Formatter parse failures are formula `#VALUE!`; resolver failures,
cancellation, allocation, resource limits, and source-version changes are
typed and are never cached as formula errors. Formula-error payloads may be
cached only after the complete required argument evaluation and source fence
succeed.

## Current implementation and profile gaps

At committed baseline HEAD `8f09231e3698`, the ODS evaluator predates the
candidate text-family files and does not establish §6.20 dispatch. The
candidate tree currently contains scalar §6.20 dispatch in
`crates/litchi-ods/src/codec/formula/evaluation/text/{core,format,search,unicode,width}.rs`,
value-mode text mapping, and the corresponding formatter and Unicode helper
paths. No candidate source hash is frozen in this contract yet. The native
receipt tooling and pinned observations are retained under
`docs/report/spec-gap-validation-evidence/ods-formula-text-functions/native/`.
Those candidate paths and receipts provide implementation/observation
evidence for review; this contract still defines the semantic profile and the
independent validation required before claiming conformance. The older
`litchi-eval` text code is compatibility evidence only: it includes host
aliases and a legacy DBCS byte approximation, and it does not establish ODS
conformance for the ODF ASC/JIS tables, Unicode-category CLEAN rules, or this
invariant formatter.

Three genuine profile/validation gaps remain before claiming full-family
conformance:

1. The built-in invariant `DOLLAR`/`FIXED`/`TEXT` formatter and its bounded
   grammar need independent conformance fixtures covering date, logical,
   complex, malformed-format, and resource cases. An ambient process locale
   or host formatter is prohibited by the source and resource contracts.
2. §6.7 leaves byte units implementation-dependent. This contract records a
   provisional UTF-8 proposal and requires a fresh byte-profile decision and
   fixtures when that gap is opened; importing a code-page/DBCS unit would be
   a separate compatibility profile.
3. Unicode normalization, host search regex/wildcard properties, and locale
   case mappings are permitted host choices. The base profile selects no
   normalization, literal searches, no regex/wildcards, and Unicode default
   case mappings. A host variation must be explicit in the evaluation context
   and cache key.

The native receipt is bounded evidence rather than a complete semantic or
resource proof: its selected closures and host profile exclude several
computed, locale, date, and error cases. Independent tests must close those
gaps. Legacy host results must not narrow the normative family or silently add
non-standard aliases.

## Validation requirements

Validation for the current batch must cover all 26 §6.20 names in scalar and
matrix contexts, including:

* finite Number, Logical, Text, Empty, Missing, Complex, formula Error,
  malformed/non-finite conversion, scalar Reference, multi-cell Reference,
  ReferenceList, inline Array, and lazy-branch inputs;
* Unicode ASCII, BMP, astral, combining-mark, normalization-variant,
  surrogate-rejection, U+0000, control, unassigned, format, non-breaking-space,
  full-width/half-width, voiced-Katakana, and Yen/reverse-solidus cases;
* all scalar-position boundaries, empty strings, zero lengths, out-of-range
  starts, empty needles, overlapping/non-overlapping substitutions, and REPT
  checked multiplication;
* exact FIND/EXACT case behavior, SEARCH case behavior, literal treatment of
  regex/wildcard metacharacters in the base profile, Unicode case expansions,
  and leftmost match/error precedence;
* CHAR/CODE 1..255 round trips, Unicode CODE values above 127, UNICHAR
  boundaries including U+0000 and surrogates, and UNICODE astral first-scalar
  results;
* exact ASC/JIS Table 33/34 mappings and unmapped-character preservation;
* CLEAN category `Cc`/`Cn` removal versus retained `Cf`/NBSP, TRIM's exact
  four-space set, PROPER letter boundaries, and normalization-sensitive
  no-normalization behavior;
* DOLLAR/FIXED default and negative decimal counts, invariant grouping and
  currency, TEXT sign/zero/literal-only sections, numeric `@` selection,
  positional mixed `0`/`#`/`?` placeholder padding (including `0?0` and
  `0.?0`), improper versus mixed fraction presentation (`#/#` and
  `# ?/?`) with rational `?` padding, normalized scientific mantissas with
  extra integer slots, fraction denominator and checked u64 whole/improper
  numerator bounds (including the exact `2^64` refusal and its represented
  predecessor), signed date serials, unsupported internal punctuation, invalid
  formats, and formatting resource limits;
* formula-error propagation, typed resolver/resource/cancellation precedence,
  source-version fences, borrowed Text, cumulative matrix read and byte
  budgets, fallible reservation cleanup, and structural scalar refusals; and
* matrix output shape/broadcasting. If demand-cache integration is later
  enabled, add complete argument/cache identity, invariant result reuse,
  position-sensitive scalar criteria, and MUNIT producer-input `[0,0]`
  behavior to that integration's receipt.

The independent oracle must compare Unicode scalar results directly and use
the built-in invariant formatter. Its gate runtime uses Python's Unicode 16.0
tables for stable mappings plus explicitly pinned Unicode 17 additions; this
is sufficient for the selected corpus and does not claim complete Unicode
17-table coverage. The evaluator's generated Unicode 17 profile remains the
normative capability identity. Native or legacy host caches can corroborate
ordinary ASCII behavior but cannot redefine the 26-name scope, normalization
policy, formatting grammar, or error/resource behavior. The future §6.7 byte gap requires a separate
oracle covering its selected UTF-8 profile.
