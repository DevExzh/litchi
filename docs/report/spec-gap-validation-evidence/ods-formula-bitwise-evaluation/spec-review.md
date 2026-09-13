# ODS OpenFormula bit-operation review

Status: normative review of OpenDocument 1.4 Part 4 §6.6 and read-only review
of the frozen bounded scalar implementation. This report covers the five
standard bit-operation functions and does not claim reference, array,
workbook, or recalculation support.

## Checked source and scope

The normative source is the checked-in specification artifact:

* `3rdparty/specs/OpenDocument-v1.4-os.zip`, SHA-256
  `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`.
* Entry `part4-formula/OpenDocument-v1.4-os-part4-formula.html`, SHA-256
  `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.

The relevant anchors are `a_6_6_Bit_operation_functions` and
`a_6_6_1_General` through `a_6_6_6_BITXOR`. Type and conversion rules are
`a_4_11_5_Integer`, `a_6_1_General`, `a_6_2_Common_Template_for_Functions_and_Operators`,
`a_6_3_1_General`, `a_6_3_5_Conversion_to_Number`,
`a_6_3_6_Conversion_to_Integer`, and `a_6_17_2_INT` for the `INT` definition.
Evaluation ordering comes from
`a_3_2_3_Operator_and_Function_Evaluation`; numeric-model and minimum-limit
boundaries are `a_3_6_Numerical_Models` and `a_3_7_Basic_Limits`.

The reviewed implementation is
`crates/litchi-ods/src/codec/formula/evaluation.rs`, SHA-256
`a150855aa43b64753ba90ba98322b213d1dfa0bc33b0e317da0f3e7e11a959db`.
The focused integration source is
`crates/litchi-ods/tests/ods_formula_bitwise_evaluation.rs`, SHA-256
`a64fcb6976541001c3ce33a51e9bfa44283717db14ea61cf853dd5325a65d8b8`.
It contains nine integration groups, including the high-bit left-shift
regressions and typed cancellation, budget, and capability cases.
Compared with the previously reviewed evaluator digest, this source changes
only the `#[inline(never)]` annotation on `apply_bitwise`; the bitwise
semantics, bounds, and error paths are unchanged.

## Normative function contract

| Function | Syntax and constraints | Result and semantics |
| --- | --- | --- |
| `BITAND` | `BITAND(Integer X; Integer Y)`; `X >= 0`, `Y >= 0` | Number whose bits are 1 exactly where both operands have a 1 bit. |
| `BITOR` | `BITOR(Integer X; Integer Y)`; `X >= 0`, `Y >= 0` | Number whose bits are 1 where either operand has a 1 bit. |
| `BITXOR` | `BITXOR(Integer X; Integer Y)`; `X >= 0`, `Y >= 0` | Number whose bits are 1 where exactly one operand has a 1 bit. |
| `BITLSHIFT` | `BITLSHIFT(Integer X; Integer N)`; `X >= 0` | If `N < 0`, `BITRSHIFT(X; -N)`; if `N = 0`, `X`; otherwise `X * 2^N`. |
| `BITRSHIFT` | `BITRSHIFT(Integer X; Integer N)`; `X >= 0` | If `N < 0`, `BITLSHIFT(X; -N)`; if `N = 0`, `X`; otherwise `INT(X / 2^N)`. |

All five calls have exactly two parameters. The shift count `N` has no
nonnegative constraint: negative counts reverse the direction. The `X >= 0`
constraint applies after the expected-type conversion, because the common
template converts parameters before the function operates on them. A
constraint violation produces an Error under §6.2. Function names ignore case
under §6.1.

Section §6.6.1 requires unsigned integer values and results through at least
48 bits, the inclusive range `0..=2^48-1` (`MAX48 = 281474976710655`). If an
operation receives or produces a value that cannot be represented within 48
bits, its behavior is implementation-defined. Bit operations are unsigned;
negative bit operands must not be interpreted with sign extension or masked
into a two's-complement word.

## Integer conversion decision

`Integer` is a Number pseudotype. Section §4.11.5 says an integer has no
fractional value and that an integer `X` equals `INT(X)`. Section §6.3.6 then
states that conversion first applies Conversion to Number and, when the
function/operator supplies no particular rounding operation, conversion of a
non-integer Number is implementation-defined. None of the five §6.6 entries
specifies a rounding operation. The profile for this batch therefore chooses
**truncate toward zero** after Number conversion and documents that as the
implementation-defined §6.3.6 choice; it must not be described as a universal
OpenFormula rounding rule.

Under this profile, `5.9` becomes integer `5`, `-1.9` becomes `-1`, and
`-0.9` becomes signed zero, which compares as `0` and therefore passes the
nonnegative constraint. A profile choosing `INT`/floor would reject that last
case, so conversion behavior must be explicit in the API/docs and fixtures.
The helper should perform the conversion in this order:

1. Convert Number, Logical, or Text according to §6.3.5. This evaluator's
   locale-independent scalar profile accepts its existing finite decimal Text
   parser; failed Text conversion is formula `#VALUE!`, and non-finite values
   are formula `#NUM!`.
2. Apply the documented truncate-toward-zero Integer conversion without a
   saturating float-to-integer cast.
3. Check the function's nonnegative constraint and the selected 48-bit profile
   before using a fixed-width integer operation.

Logical values follow Conversion to Number (`TRUE` = 1 and `FALSE` = 0).
Formula Error values remain Error values under §6.3.1. References, arrays,
names, and labels stay capability refusals in the current scalar evaluator;
future reference support must use the §6.3.2 implied-intersection and empty
cell rules rather than silently selecting or flattening data here.

## Deterministic 48-bit profile

The recommended interoperable window is `0..=MAX48` for every bit operand and
every bit-operation result. For `BITAND`, `BITOR`, and `BITXOR`, conversion or
the nonnegative constraint is checked before the operation. The profile
returns formula `#NUM!` for a negative converted operand, an operand outside
the 48-bit window, or a result outside the window. This is a documented choice
for §6.6.1's implementation-defined cases, not a claim that every evaluator
must use `#NUM!`.

The shift count is signed and should be handled by sign and magnitude rather
than by recursively calling the opposite public function or negating an
unchecked minimum integer. There is no normative maximum shift count. The
profile may accept any finite integer shift count representable by the input
Number model, including magnitudes outside the 48-bit value window, as an
explicit implementation-defined extension. Such counts must be classified
without a loop or a lossy cast:

* A positive right shift of a 48-bit `X` by magnitude at least 48 is `0`.
* A positive left shift of `X = 0` is `0` for any finite positive count. For
  nonzero `X`, a count at least 48, or any count whose checked result exceeds
  `MAX48`, produces profile `#NUM!`.
* A negative `N` reverses these two cases. In particular,
  `BITRSHIFT(1; -48)` is a left shift and therefore `#NUM!`, while
  `BITLSHIFT(1; -48)` is `0`.
* `N = 0` returns `X` exactly. Fractional counts are converted before this
  sign test, so `1.9` is one and `-0.9` is zero under the selected profile.

Implementations may instead reject a shift count outside a signed 48-bit
window, but must document that choice. They must not silently wrap the count,
use a platform-dependent shift mask, call floating `powf` for an unbounded
exponent, or iterate once per requested bit. For in-window values, a checked
`u64`/`u128` bit operation is exact because all values through 48 bits are
exactly representable by the evaluator's `f64` Number. The positive right shift
uses integer division, which agrees with `INT` because `X` is nonnegative.

## Evaluation, errors, and limits

The §3.2.3 default is eager evaluation: both argument expressions are
computed before conversion and application. A formula Error is a value and
the function returns one input Error; the §6.1 recommendation is the leftmost
source-order Error when there is more than one. Cancellation, resource-limit,
allocation, and unsupported-capability failures remain evaluator failures and
are not converted to formula Errors. An invalid text-to-number conversion is
the profile's `#VALUE!`; a valid but domain-invalid integer or an
out-of-window bit result is `#NUM!`.

The exact-two-argument contract should reject wrong arities as a formula
`#VALUE!` after preserving the evaluator's eager child/error policy. A missing
argument node is therefore a formula `#VALUE!` in this profile, while a
malformed immutable AST is an evaluator `InvalidExpression` failure. No
bitwise function has lazy branches or a short-circuit exception.

Each admitted function application and argument visit must consume the
existing evaluation work budget; existing text conversion must charge input
bytes and check cancellation. The bit operation itself is O(1), allocates no
storage, and must use checked arithmetic. A huge shift must finish with a
bounded comparison and cannot cause a shift panic, integer overflow, or work
growth proportional to `N`. The operation has no workbook/cache/package/network
side effect.

## Required profile fixtures

| Expression | Expected result or boundary |
| --- | --- |
| `BITAND(6;3)` / `BITOR(4;1)` | Number `2` / Number `5` |
| `BITXOR(6;3)` | Number `5` |
| `BITLSHIFT(3;2)` / `BITRSHIFT(13;2)` | Number `12` / Number `3` (`INT(13/4)`) |
| `BITLSHIFT(8;-2)` / `BITRSHIFT(3;-2)` | Number `2` / Number `12` |
| `BITLSHIFT(9;0)` / `BITRSHIFT(9;0)` | Number `9` / Number `9` |
| `BITAND(5.9;3.1)` | Number `1` under truncate-toward-zero profile |
| `BITAND(-0.9;1)` | Number `0` under the documented profile choice |
| `BITAND(TRUE();1)` / `BITOR("3.9";1)` | Number `1` / Number `3` |
| `BITAND(-1;0)` / `BITLSHIFT(-1;1)` | formula `#NUM!` |
| `BITLSHIFT(1;48)` / `BITRSHIFT(1;-48)` | formula `#NUM!` |
| `BITRSHIFT(281474976710655;48)` | Number `0` |
| `BITAND("not-a-number";1)` | formula `#VALUE!` |
| `BITOR(#N/A;1)` | formula `#N/A` |
| `BITAND(1)` / `BITOR(1;2;3)` | formula `#VALUE!` for wrong arity |
| `BITAND(#N/A;#DIV/0!)` | formula `#N/A` under leftmost-error retention |

## Frozen implementation disposition

The frozen evaluator admits all five names in its eager-function scheduler and
routes them through `apply_bitwise`. Every call requires exactly two
arguments; wrong arity consumes already-scheduled child values and returns
the profile's formula `#VALUE!`, while capability failures in those children
remain typed evaluator failures. The pairwise path converts both operands
through the existing Number conversion before truncating toward zero, checks
the unsigned 48-bit window, and applies `&`, `|`, or `^` on a fixed-width
integer. Formula Errors remain values and are retained in source order.

The shift path converts X to the same checked unsigned operand and N to the
profile Integer value, then branches directly on N's sign. It handles signed
zero as zero and reverses direction for negative counts without recursive
evaluator calls. Large finite counts are classified by comparison rather than
cast into a native shift width or iterated. Positive right shifts at or above
48 return zero; positive left shifts first check
`X <= BIT_MAX >> amount` before `checked_shl`. This precheck closes the
high-bit-loss case where `checked_shl` alone would accept a count below the
native width while discarding bits beyond the selected 48-bit result. The
focused fixtures cover `1 << 47`, overflowing `MAX48 << 47`, overflowing
`2^47 << 17`, reverse-direction large counts, and zero-valued large shifts.

The operation is allocation-free after AST evaluation, performs bounded
constant-time arithmetic independent of the requested shift magnitude, and
keeps conversion byte work and cancellation checks in the existing evaluator
paths. The test source covers eager formula-error propagation, unselected
conditional bitwise work, uncaught reference/array refusals, cancellation
before work, local work and storage limits, result lifetime, and explicit
stack traversal for nested calls. No additional semantic or safety blocker
was found in this bounded review. The full repository gates remain owned by
the parent; any later source change requires refreshing both digests and
rechecking the profile fixtures above.
