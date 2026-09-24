# ODS OpenFormula Roman-number review

Status: normative review of OpenDocument 1.4 Part 4 §6.19.2 and §6.19.17,
with a deterministic profile for `ARABIC` and `ROMAN`. This report covers
Roman-number conversion only. It does not claim support for references,
arrays, workbook recalculation, locale-dependent number formats, or cache and
package mutation.

## Checked source and scope

The normative source is the checked-in specification artifact:

* `3rdparty/specs/OpenDocument-v1.4-os.zip`, SHA-256
  `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`.
* Entry `part4-formula/OpenDocument-v1.4-os-part4-formula.html`, SHA-256
  `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.

The function anchors are `a_6_19_Number_Representation_Conversion_Functions`,
`a_6_19_1_General`, `a_6_19_2_ARABIC`, and `a_6_19_17_ROMAN`. The common
evaluation and conversion anchors are `a_3_2_3_Operator_and_Function_Evaluation`,
`a_6_1_General`, `a_6_2_Common_Template_for_Functions_and_Operators`,
`a_6_3_1_General`, `a_6_3_5_Conversion_to_Number`,
`a_6_3_6_Conversion_to_Integer`, `a_4_11_5_Integer`, and
`a_6_17_2_INT`.

## `ARABIC` contract

`ARABIC(Text X)` returns a Number. The direct constraint permits only the
seven ASCII Roman symbols, or an empty string. The exact accepted code points
are uppercase/lowercase `M`, `D`, `C`, `L`, `X`, `V`, and `I`; Unicode Roman
numeral characters, signs, spaces, punctuation, and other letters are not
substitutes for those code points. Case is ignored.

For each symbol, subtract its value if a strictly larger symbol occurs to its
right, directly or indirectly; otherwise add it. A right-to-left scan that
keeps the largest value already seen is equivalent to this rule. It accepts
noncanonical spellings because §6.19.2 defines the value rule rather than a
canonical Roman grammar. Thus `ARABIC("IV")` is 4, `ARABIC("IIX")` is 8,
`ARABIC("IXX")` is 19, `ARABIC("IC")` is 99, and
`ARABIC("MIM")` is 1999. The empty string returns 0. The
`ARABIC(ROMAN(x;any)) = x` identity applies when `ROMAN` succeeds; it does
not require every accepted `ARABIC` spelling to be canonical.

The scanner must charge the input bytes, check cancellation in bounded
chunks, and use checked accumulator arithmetic. With the evaluator's finite
Number model, a representable finite Number result is returned; an arithmetic
overflow or non-finite conversion is a formula numeric error under this
profile. Text coercion uses the evaluator's existing scalar `Text` conversion
before scanning, and formula Errors remain formula Error values.

## `ROMAN` contract and conversion choices

`ROMAN(Integer N [; Integer Format = 0])` returns Text. §6.19.17 constrains
`0 <= N < 4000` and `0 <= Format <= 4`. The optional second parameter defaults
to format 0 when it is omitted. The standard permits a legacy evaluator with
a distinct Logical type to accept a scalar Logical format: `TRUE` means
format 0 and `FALSE` means format 4. This profile accepts those mappings even
though ordinary Number `1` is not itself Logical.

The `Integer` pseudotype first performs Conversion to Number. §6.3.6 leaves
the conversion of a non-integer Number implementation-defined when the
function gives no function-specific rounding rule. This profile makes the
choice explicit: finite Numbers are truncated toward zero, then the Roman
domain constraint is checked. Consequently `ROMAN(4.9)` is `IV`,
`ROMAN(3999.9)` is the representation of 3999, and `ROMAN(-0.9)` becomes
signed zero and follows the zero result rule. A non-finite value, a value
whose converted integer is negative, or a value at least 4000 returns the
profile's numeric formula Error. Text and Logical arguments follow the
evaluator's existing scalar Number conversion (`TRUE` = 1 and `FALSE` = 0);
the locale-dependent alternatives allowed by §6.3.5 are outside this
locale-free profile.

Zero has an explicit normative rule: §6.19.17 says that `N = 0` returns an
empty string. A later nonnormative `should` sentence says evaluators that
accept zero should return `"0"`. The profile follows the explicit rule so
that the stated identity also holds for zero: `ROMAN(0;f) = ""` and
`ARABIC("") = 0` for every valid format. The later `"0"` sentence is a
compatibility choice for another profile, not a reason to violate the direct
identity in this one.

§6.19.17 does not define a rounding rule for `Format`; this profile applies
the same truncate-toward-zero Integer choice and then checks `0..=4`. An
explicit missing trailing child is distinct from an omitted optional
component and is reported as a formula value error by the existing scalar
arity policy. All supplied child expressions are evaluated eagerly under
§3.2.3, including a Format expression whose value is later irrelevant to an
implementation-defined negative case. Formula Errors preserve source order;
cancellation, resource-limit, allocation, and unsupported-capability failures
remain evaluator failures.

## Format levels

The table in §6.19.17 is the contract for formats 0 through 3. It is not a
length-monotonicity theorem: format 1 permits `V` as a subtractor while
format 2 forbids `V`. Each implementation must apply the row's explicit
permissions rather than infer that every larger format has a representation
no longer than every smaller format.

| Format | Allowed subtraction and trailing rule | Deterministic profile |
| --- | --- | --- |
| `0`, omitted, or `TRUE` | Only powers of 10 (`I`, `X`, `C`) may subtract; the larger value may be at most ten times the subtractor. A symbol after the larger value must be smaller than the subtractor. | Greedy greatest admissible single symbol or pair, repeatedly. This yields classic forms such as `XLIX` and `CDXCIX`. |
| `1` | Powers of 10, `L`, and `V` may subtract; the larger value may be at most ten times the subtractor. The same trailing rule applies. | Greedy greatest admissible chunk. `VL` is permitted for 45 and `LDVLIV` for 499. |
| `2` | Powers of 10 and `L`, but not `V`, may subtract even when the larger value is more than ten times the subtractor. The same trailing rule applies. | Greedy greatest admissible chunk under this literal table. This permits `ID` for 499 and `IL` for 49; Excel-style `XDIX` is another compatibility spelling, not a requirement of the checked-in ODF prose. |
| `3` | Powers of 10, `L`, and `V` may subtract even when the larger value is more than ten times the subtractor. The same trailing rule applies. | Greedy greatest admissible chunk. This also permits `ID` for 499; `VDIV` is a compatibility spelling, not required by the ODF table. |
| `4` or `FALSE` | Produce the fewest Roman digits possible. | Use the bounded global minimum described below; tie choices are documented profile behavior because the standard does not prescribe one. |

The wording “powers of 10” includes `I`, `X`, and `C`. The table does not
limit a permitted power-of-10 subtractor to the immediately adjacent Roman
denomination. Therefore, under the checked-in ODF wording, `ID` (500 - 1) is
valid in formats 2 and 3: `D` is more than ten times `I`, the subtractor is a
power of 10, and there is no following value that could violate the trailing
rule. No normative output examples in the local Part 4 HTML override this
reading. If Excel-compatible spellings are desired, they need a separately
named compatibility mode and must not be presented as the ODF requirement.

The greedy chunk rule satisfies the trailing condition: if a pair consists of
`small` before `large` and a later chunk began with a symbol of value at least
`small`, then the residual would be at least `small`. The single symbol
`large` would then have been an admissible chunk at least as large as
`large - small`, contradicting selection of the pair as the greatest chunk.
The next chunk therefore starts below `small`.

## Format 4 minimum and bounded construction

Use the coin values `[1, 5, 10, 50, 100, 500, 1000]` with symbols
`IVXLCDM`. A representation is an integer coefficient vector `c` satisfying

```
N = c0*1 + c1*5 + c2*10 + c3*50 + c4*100 + c5*500 + c6*1000
length = |c0| + |c1| + ... + |c6|
```

After cancelling opposite coefficients at one denomination, emit negative
coefficients in increasing denomination order, then positive coefficients in
decreasing denomination order. For every negative coin, a larger positive
coin is to its right, so the §6.19 indirect-subtraction rule subtracts it;
the positive coins add. For `N > 0`, the carry recurrence below guarantees
that any negative coefficient is followed by a positive higher denomination,
so the emitted text evaluates back to `N` under `ARABIC`.

Process the adjacent ratios `[5, 2, 5, 2, 5, 2]` from low to high. At each
level, if the current carry is `u` and the ratio is `q`, its coefficient must
be congruent to `u mod q`. An L1-minimal vector never needs a coefficient
whose magnitude is at least `q`: transfer `q` units to the next denomination,
which preserves the represented value and lowers the total coefficient count
by at least `q - 1` (and more when opposite signs cancel). Therefore only
these two representatives need consideration:

* the nonnegative residue `r`, and
* when `r != 0`, the negative carry `r - q`.

Set `u = (u - c_i) / q` and continue. Six binary choices produce at most 64
candidate vectors. Keep the lowest `length`; the top carry is nonnegative.
This is an exhaustive finite search over all possible minimum-length
representations, not a claim that a greedy choice is globally minimal.

The emitted order is important. A negative higher denomination placed after a
positive lower denomination would be interpreted as an addition by `ARABIC`;
the ordered negative-then-positive emission prevents that. The recurrence also
shows that when a negative coefficient is selected, the next carry is at least
one. If later choices reduce the carry, some positive higher coefficient has
already occurred; otherwise the final `M` carry is positive. Thus every
negative emitted coin has a strictly larger positive symbol to its right.

The ODF requirement specifies the minimum length but does not specify which
spelling wins a tie. The profile chooses the first minimum in deterministic
mask order. For example, it may emit `IMM` for 1999 while `MIM` has the same
three-character length; both evaluate to 1999. Other useful shortest results
are `IIX` for 8, `IIM` for 998, `ID` for 499, `IM` for 999, and `IMMMM` for
3999. Repeated `M` and indirect subtractive forms are allowed by the local
Roman symbol/value rules; format 4 does not impose an unmentioned canonical
Roman grammar.

The search uses fixed-size arrays and no value-indexed table or recursion.
It should charge each candidate/level operation to the existing work budget,
check cancellation through that budget, and perform one checked output-size
admission before constructing the result. The final Text reservation must be
fallible and retained with the evaluated result; a failed reservation must
not leave budget or partial output state behind. The 64-candidate search is a
constant bound independent of `N`; scanning an `ARABIC` input remains bounded
by input bytes.

## Error, capability, and side-effect boundary

Wrong arity, invalid Roman characters, invalid text conversion, and failed
format constraints are formula errors in the scalar profile. Numeric domain
failures use the profile's numeric formula error. §6.1 allows one of several
input Errors and recommends the leftmost source-order Error; this profile
retains that behavior. A reference, array, name, label, or other unsupported
function remains a typed evaluator capability refusal. Conditional functions
may avoid evaluating an unselected Roman branch, but an evaluator refusal is
not a formula Error for `IFERROR`/`IFNA` to catch.

No Roman conversion performs workbook reads, formula-cache mutation, package
publication, network access, or automatic recalculation. Input scans and
output construction are governed by the existing text, work, memory, and
cancellation limits. Numeric value does not control loop count.

## Profile fixtures

| Expression | Expected result or boundary |
| --- | --- |
| `ARABIC("")` / `ARABIC("iv")` | Number `0` / Number `4` |
| `ARABIC("IIX")` / `ARABIC("IXX")` / `ARABIC("IC")` | Number `8` / `19` / `99` |
| `ROMAN(0;0)` through `ROMAN(0;4)` | empty Text `""` for all valid formats |
| `ROMAN(4)` / `ROMAN(8;0)` / `ROMAN(8;4)` | `IV` / `VIII` / `IIX` |
| `ROMAN(49;0..3)` | `XLIX`, `VLIV`, `IL`, `IL` under the greedy table profile |
| `ROMAN(499;0..4)` | `CDXCIX`, `LDVLIV`, `ID`, `ID`, `ID` under the literal ODF table profile |
| `ROMAN(998;0..4)` | `CMXCVIII`, `CMXCVIII`, `XMVIII`, `VMIII`, `IIM` |
| `ROMAN(1999;4)` / `ROMAN(3999;4)` | a three-digit shortest spelling (profile `IMM`) / `IMMMM` |
| `ARABIC(ROMAN(n;f))` | Number `n` for every `0 <= n < 4000` and valid `f` |
| `ROMAN(4.9)` / `ROMAN(499;1.9)` | `IV` / `LDVLIV` under truncate-toward-zero conversion |
| `ROMAN(-1)` / `ROMAN(4000)` / `ROMAN(1;5)` | numeric formula Error |
| `ARABIC("A")` / `ARABIC("IV ")` / `ARABIC("Ⅰ")` | `#VALUE!` formula Error |
| `ROMAN(#N/A;#DIV/0!)` | leftmost formula Error `#N/A` |
| Roman call with a reference or array argument | typed capability refusal |

The `ID` entries above are deliberate: they follow the exact checked-in
§6.19.17 table. `XDIX` and `VDIV` are useful interoperability fixtures for
an Excel-compatible profile, but should not be used as ODF conformance tests
unless that profile explicitly chooses those spellings.

## Implementation review disposition

The current `evaluation/roman.rs` snapshot implements the profile with a
right-to-left `ARABIC` scan, fixed-size greedy chunks for formats 0 through 3,
and the six-ratio bounded search for format 4. The current source and focused
test digests are:

* `crates/litchi-ods/src/codec/formula/evaluation/roman.rs` —
  `3441233624f68e5efc72157953d3b7137bbb86f6c9835fd809eee40674342da1`.
* `crates/litchi-ods/src/codec/formula/evaluation.rs` —
  `5b2320a6204b96185b6e91eaa27d3f166ac19555e71b307c79a0600329975cf4`.
* `crates/litchi-ods/tests/ods_formula_roman_evaluation.rs` —
  `4bcc00fd7278787c1ae77a4bd58118677a743a92bf023f6f725309bd6b8877ff`.

The implementation's fixed arrays and reservation paths fit the bounded
profile. The focused tests are aligned with the literal ODF table: formats 2
and 3 may return `ID` for 499, whereas the Excel-compatible `XDIX` and `VDIV`
spellings are not normative requirements. The independent shortest-length
oracle and all-value round trips are the right evidence for format 4; exact
tie spelling remains a profile choice. The evaluator's string scanner now
checks cancellation at 4096-byte windows while allowing a doubled quote to
cross a window boundary; the focused concatenation test verifies that this
windowing preserves literal contents. The final gate receipt records 926
tests across 56 targets, zero failures or ignored tests, and two doctests; the
performance receipt is a separate pending artifact and is not inferred here.
