# ODS OpenFormula radix-conversion review

Status: normative review of OpenDocument 1.4 Part 4 §6.19 and the frozen
bounded scalar profile for the fourteen BASE/DECIMAL and binary, decimal,
hexadecimal, and octal conversion functions. `ARABIC` (§6.19.2) and `ROMAN`
(§6.19.17) are intentionally outside this batch. This report does not claim
reference, array, workbook, locale, or recalculation support.

## Checked source and scope

The normative source is the checked-in specification artifact:

* `3rdparty/specs/OpenDocument-v1.4-os.zip`, SHA-256
  `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`.
* Entry `part4-formula/OpenDocument-v1.4-os-part4-formula.html`, SHA-256
  `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.

The primary anchors are `a_6_19_Number_Representation_Conversion_Functions`,
`a_6_19_1_General`, and `a_6_19_3_BASE` through `a_6_19_16_OCT2HEX`.
Conversion and evaluation rules are `a_3_2_3_Operator_and_Function_Evaluation`,
`a_3_6_Numerical_Models`, `a_3_7_Basic_Limits`, `a_4_11_5_Integer`,
`a_6_1_General`, `a_6_2_Common_Template_for_Functions_and_Operators`,
`a_6_3_1_General`, `a_6_3_5_Conversion_to_Number`,
`a_6_3_6_Conversion_to_Integer`, and `a_6_17_2_INT`.

## Function inventory and hard constraints

Section §6.19.1 says that `xxx2BIN`, `xxx2OCT`, and `xxx2HEX` return Text,
while `xxx2DEC` returns Number. All twelve small converters accept
`TextOrNumber`; a Number is interpreted as its digits when printed in base
10. The hexadecimal input alphabet is ASCII `0-9`, `A-F`, or `a-f`, and
hexadecimal output shall use that alphabet and should use uppercase. The
small converters use at most ten source or target digits in their
interoperable envelope.

| Function | Signature and direct constraints | Fixed-width sign and result |
| --- | --- | --- |
| `BASE` | `BASE(Integer X; Integer Radix[; Integer MinimumLength])`; `X >= 0`, `2 <= Radix <= 36`, `MinimumLength >= 0` | Text, digits `0-9` then uppercase `A-Z`; no fixed sign bit |
| `DECIMAL` | `DECIMAL(Text X; Integer Radix)`; `2 <= Radix <= 36` | Number, parsed in the requested radix |
| `BIN2DEC` | `BIN2DEC(TextOrNumber X)`; binary digits only, at least one; Number must satisfy `INT(X)=X` | Number; source bit 10 is sign, two's complement |
| `BIN2HEX` | `BIN2HEX(TextOrNumber X[; Number Digits])`; same binary input constraints | Text; source bit 10, target bit 40; negative input ignores `Digits` |
| `BIN2OCT` | `BIN2OCT(TextOrNumber X[; Number Digits])`; same binary input constraints | Text; source bit 10, target bit 30; negative input ignores `Digits` |
| `DEC2BIN` | `DEC2BIN(TextOrNumber X[; Number Digits])`; decimal digits, at least one; Number must satisfy `INT(X)=X` | Text; signed target width 10, recommended `-512..511` |
| `DEC2HEX` | `DEC2HEX(TextOrNumber X[; Number Digits])`; same decimal input constraints | Text; signed target width 40, recommended `-2^39..2^39-1` |
| `DEC2OCT` | `DEC2OCT(TextOrNumber X[; Number Digits])`; same decimal input constraints | Text; signed target width 30, recommended `-2^29..2^29-1` |
| `HEX2BIN` | `HEX2BIN(TextOrNumber X[; Number Digits])`; hexadecimal digits, at least one; Number must satisfy `INT(X)=X` | Text; source bit 40, target bit 10; negative input ignores `Digits` |
| `HEX2DEC` | `HEX2DEC(TextOrNumber X)`; hexadecimal digits, at least one; Number must satisfy `INT(X)=X` | Number; source bit 40 is sign |
| `HEX2OCT` | `HEX2OCT(TextOrNumber X[; Number Digits])`; hexadecimal digits, at least one; Number must satisfy `INT(X)=X` | Text; source bit 40, target bit 30; negative input ignores `Digits` |
| `OCT2BIN` | `OCT2BIN(TextOrNumber X[; Number Digits])`; octal digits, at least one; Number must satisfy `INT(X)=X` | Text; source bit 30, target bit 10; negative input ignores `Digits` |
| `OCT2DEC` | `OCT2DEC(TextOrNumber X)`; octal digits, at least one; Number must satisfy `INT(X)=X` | Number; source bit 30 is sign |
| `OCT2HEX` | `OCT2HEX(TextOrNumber X[; Number Digits])`; octal digits, at least one; Number must satisfy `INT(X)=X` | Text; source bit 30, target bit 40; negative input ignores `Digits` |

The clauses saying an evaluator *may* process no more than ten source digits
or the listed signed decimal range describe the standard's interoperable
small-value envelope. The profile adopts that envelope for the twelve small
converters. Values outside a source or target signed width are reported as
profile `#NUM!`; this is a documented boundary for the specification's
implementation-defined cases. The general `BASE` and `DECIMAL` functions do
not acquire an arbitrary ten-digit or `u64` limit from these small helpers.

## Integer and Text conversion

The `Integer` pseudotype is a subtype of Number. Section §6.3.6 first applies
Conversion to Number and then says that a function-specific conversion from a
non-integer Number is implementation-defined when the function gives no
rounding rule. The twelve `xxx2yyy` constraints add a function-specific
integrality requirement: when X is considered as a Number, `INT(X) = X`.
Therefore a numeric `5.9` is invalid for `DEC2BIN`, `BIN2DEC`, `HEX2DEC`, and
the other small converters; it must not be silently truncated. Text such as
`"5.9"` is also invalid because the decimal-digit grammar has no decimal
point. This is separate from the profile choice for Integer parameters of
`BASE` and `DECIMAL`, which may truncate toward zero after Number conversion
because those functions do not impose the small-converter `INT(X)=X` rule.

Text accepted by the small converters is ASCII and strict after any permitted
leading minus sign:

* Binary source text is one or more `0` or `1` characters. Spaces, tabs,
  `+`, `-`, underscores, prefixes, suffixes, and other characters are invalid.
* Octal source text is one or more `0` through `7` characters with the same
  rejection of whitespace and signs.
* Hexadecimal source text is one or more ASCII hexadecimal digits. A leading
  sign is not part of the §6.19 hex input grammar.
* Decimal source text for `DEC2BIN`, `DEC2HEX`, and `DEC2OCT` is one or more
  decimal digits and may have one leading `-`, as their semantics explicitly
  allow. A leading `+`, spaces, decimal point, exponent, or internal sign is
  not accepted.

For a Number input, the evaluator uses a canonical base-10 integer digit
rendering after the required integrality check. It must not use a locale
decimal separator or an exponent spelling that changes the digit grammar.
Logical input is implementation-defined by the specification; this profile
converts it through Conversion to Number (`TRUE` = 1, `FALSE` = 0) and then
applies the relevant digit rules. References, arrays, names, and labels remain
typed capability refusals in the scalar evaluator.

## Fixed-width signed conversions

The source and target widths are fixed by the bit positions in §6.19:

| Radix | Width | Sign bit | Signed range when interpreted as a decimal value |
| --- | ---: | ---: | ---: |
| Binary | 10 bits | bit 10 | `-512..511` |
| Octal | 30 bits | bit 30 | `-2^29..2^29-1` |
| Hexadecimal | 40 bits | bit 40 | `-2^39..2^39-1` |

The sign bit is the high bit of the fixed width, not the high bit of a short
spelling. Thus a short binary `"1"` is positive one, while ten binary digits
beginning with `1` are negative; ten octal digits beginning with `4` through
`7` are negative; and ten hexadecimal digits beginning with `8` through `F`
are negative. A source negative value is interpreted as unsigned raw value
minus `2^width`. A source positive value is converted normally. When a
negative value is rendered at a wider target width, it is sign-extended and
encoded in two's complement. When a value cannot fit the target signed width,
the profile returns `#NUM!` rather than truncating high bits into a different
sign.

For text outputs, a positive value uses the shortest representation unless
`Digits` requests padding. A negative value is rendered in the target's full
fixed width and ignores `Digits`, as specified by each conversion entry. The
hexadecimal output alphabet is uppercase `A-F`. This preserves the source
sign-width identity: for example, `DEC2BIN(-1)` is ten `1` digits,
`DEC2HEX(-1)` is ten `F` digits, and `DEC2OCT(-1)` is ten `7` digits.

## Optional `Digits` and `MinimumLength`

The standard specifies optional `Digits` parameters for the nine text-output
small converters but gives no direct constraint or rounding operation for
that Number parameter. It says that positive results are padded with leading
zeroes up to the requested count, negative results ignore the count, and that
behavior is implementation-defined when the natural result needs more digits
than requested. The profile makes these choices explicit:

* Convert `Digits` through the existing finite Number conversion and truncate
  toward zero. Require a finite nonnegative integer in `0..=10`; negative,
  non-finite, and greater-than-ten values produce `#NUM!` before allocation.
  `Digits = 0` means no padding.
* If a positive natural result is already longer than `Digits`, return the
  natural result rather than truncating it or producing a width-dependent
  sign. This is the profile's choice for the standard's implementation-defined
  too-small case.
* If the source value is negative, the optional expression is still evaluated
  eagerly under §3.2.3; a formula Error in it propagates even though its value
  is ignored by the conversion semantics. After successful evaluation, its
  width value is ignored and the target's full negative width is emitted.
* One child means the optional parameter was omitted and uses the shortest
  positive result. An explicit missing trailing child, as in `DEC2BIN(5;)`,
  has no §6.19 default and is a formula `#VALUE!` in this profile. This keeps
  missing arguments distinct from omitted syntax; a future compatibility
  profile may document accepting the empty slot as omission.

`BASE` uses `MinimumLength`, which is not the same as the small-converter
`Digits` width. An omitted minimum emits the shortest text. A supplied
nonnegative integral minimum pads with leading zeroes to exactly that length,
but is ignored when the natural text is longer. The profile permits finite
minimum lengths beyond ten, subject to the evaluator's text, work, storage,
and allocation limits. A length that cannot be admitted is a typed resource
failure rather than an unchecked allocation.

## `BASE` and `DECIMAL`

`BASE` uses digits `0-9` followed by uppercase `A-Z`, so
`BASE(45745;36)` is `"ZAP"`. Its only direct value constraints are a
nonnegative X, radix 2 through 36, and nonnegative minimum length. The profile
accepts every finite integral value representable by the evaluator's Number
model after its documented Integer conversion; it must not impose an
arbitrary `u64` or 48-bit cap. In particular, powers and finite values above
`2^64-1` remain eligible. Conversion should use checked fixed-limb or
equivalent arithmetic, not a saturating float-to-`u64` cast. The resulting
digits and requested padding are admitted only after checked byte/work and
storage limits, so a huge finite Number cannot force an unbounded string or
loop.

`DECIMAL` accepts ASCII digits and letters whose numeric value is below the
radix. For radix greater than ten, `A-Z` and `a-z` are equivalent; for radix
ten or below, alphabetic characters are invalid. The following lexical
exceptions are exact and local to the specified radix:

* Leading ASCII space U+0020 and tab U+0009 are always ignored. Trailing or
  internal spaces/tabs are invalid, as are other Unicode whitespace characters.
* With radix 16, one leading regular-expression `0?[Xx]` is ignored (`XFF`
  and `0xFF` are accepted), and one trailing `H` or `h` is ignored.
* With radix 2, one trailing `B` or `b` is ignored. There is no general `0b`
  prefix rule.

The profile applies these strips once, then requires at least one digit and
validates every remaining ASCII character. It returns `#VALUE!` for an empty
or invalid post-strip string. The specification explicitly defines invalid
character behavior but leaves the empty-string result unspecified; rejecting
empty text is the deterministic profile choice. Leading-only stripping means
`DECIMAL(" 101b";2)` is 5, while trailing whitespace, `DECIMAL("0b101";2)`,
and `DECIMAL("0xFF";10)` are errors. Radix is converted to a finite Integer
through the profile's truncate-toward-zero choice, then checked against 2..36.

The Number result of `DECIMAL` follows the evaluator's documented numeric
model. Values within its finite range are returned as finite Number values;
an accumulated result that is non-finite produces profile `#NUM!`. Text
length and arithmetic work are charged incrementally, and parsing does not
stop at a smaller `u64` boundary when the Number model can still represent the
result.

## Evaluation, errors, and limits

All parameters are eagerly evaluated under §3.2.3. This includes an optional
`Digits` or `MinimumLength` expression that will later be ignored for a
negative fixed-width input. Formula Error values remain values; §6.1
recommends returning the leftmost source-order Error when several are present.
Cancellation, resource-limit, allocation, and unsupported-capability results
remain evaluator failures and are not converted into formula Errors.

The profile maps invalid digits, invalid text conversion, missing/incorrect
arguments, and malformed radix/width syntax to `#VALUE!`; numeric domain,
fixed-width overflow, non-finite result, and out-of-profile width values map
to `#NUM!`. The standard only requires an Error for constraints and leaves the
specific Error spelling to the implementation, so these mappings are profile
choices. Wrong arity is a formula error after preserving the eager child/error
policy. No conversion function performs workbook, cache, package, network, or
automatic recalculation work.

Every input byte and emitted digit should be charged to the existing evaluator
work budget and checked for cancellation in bounded chunks. `BASE` and
`DECIMAL` may have work proportional to the input/output digit count, but no
step may be proportional to the numeric value itself. Output strings require
fallible reservation and retained-result accounting. Fixed-width converters
have at most ten output digits; they must not allocate based on an unchecked
`Digits` Number.

## Frozen implementation disposition

The reviewed implementation matches this profile. `radix.rs` uses a fixed
1024-bit magnitude for `BASE` and `DECIMAL`, which covers every finite integer
represented by the evaluator's `f64` Number model without narrowing through
`u64`; it converts accumulated `DECIMAL` values back with IEEE nearest-even
rounding. The direct converters enforce the 10-, 30-, and 40-bit signed
domains, strict source alphabets, integral numeric inputs, eager optional
arguments, two's-complement sign extension, and checked output reservation.
Formula errors remain scalar values, while cancellation, resource limits,
allocation failures, and unsupported references/arrays remain typed evaluator
failures. Focused coverage contains twelve radix integration tests, including
large finite values, lexical exceptions, all fixed-width boundaries, eager
errors, lazy branches, cancellation, storage/text limits, and retained output
lifetime.

The frozen source and test inputs are bound by these SHA-256 digests:

* `crates/litchi-ods/src/codec/formula/evaluation/radix.rs` —
  `48cf638a72a4946f8f3c8284b34ff817e4c5ba77c962776ad6c2ae981991cf9d`.
* `crates/litchi-ods/src/codec/formula/evaluation.rs` —
  `15ff312b132e6d3d0d97a2cd6ff897a8b33fcbd8b366b35e20737c6d7cfeaa84`.
* `crates/litchi-ods/tests/ods_formula_radix_evaluation.rs` —
  `c2d0ee542fdc4e465b6236b1e6234d1b7980791a82995d7c315239af48ba32bf`.

The retained `gates/results.json` receipt is for head
`e1976c59d5c4785e0c73a5d27e9349ba082cca2f` and records unchanged source
hashes, 915 passing tests across 55 targets, two passing doctests, and passing
all-target Clippy, warning-denied documentation, and formatting checks. No
additional blocker was found within this radix-conversion scope. The stated
reference/array, locale, workbook, and recalculation boundaries remain
intentional follow-up work.

## Profile fixtures

| Expression | Expected result or boundary |
| --- | --- |
| `BASE(45745;36)` / `BASE(10;2;6)` | Text `"ZAP"` / `"001010"` |
| `BASE(0;10)` / `BASE(18446744073709551616;16)` | Text `"0"` / `"10000000000000000"` (finite value above `u64`) |
| `DECIMAL("zap";36)` | Number `45745` |
| `DECIMAL("  \t0xFFh";16)` / `DECIMAL("101b";2)` | Number `255` / Number `5` |
| `DECIMAL("1 0";2)` / `DECIMAL("0b101";2)` / `DECIMAL("0xFF";10)` | formula `#VALUE!` |
| `BIN2DEC("101")` / `BIN2DEC("1111111111")` | Number `5` / Number `-1` |
| `BIN2HEX("1111111111")` / `BIN2OCT("1111111111")` | ten-digit negative sign extension (`FFFFFFFFFF` / `7777777777`) |
| `DEC2BIN(-1)` / `DEC2HEX(-1)` / `DEC2OCT(-1)` | ten `1` / ten `F` / ten `7` digits |
| `DEC2BIN(5;8)` / `DEC2HEX(255;4)` | Text `"00000101"` / `"00FF"` |
| `DEC2BIN(5.9)` / `DEC2HEX("5.9")` | formula `#VALUE!` for required integral/digit grammar |
| `DEC2BIN(-1;1)` | ten-digit negative result; `Digits` ignored after eager evaluation |
| `DECIMAL("";10)` / `DECIMAL("1";1)` | formula `#VALUE!` for empty text / `#NUM!` for radix outside 2..36 |
| Radix-function reference/array operands | typed capability refusal, never implicit cell selection or flattening |
| `BIN2DEC(#N/A)` | formula `#N/A` propagated as a value |

This report records the normative contract, explicit profile choices, frozen
implementation digests, and gate disposition without weakening the strict
small-converter grammar or the explicit large-`BASE`/`DECIMAL` limits.
