# Rounding-function contract

This contract defines the complete ODF 1.4 Part 4 §6.17 batch. It is an
implementation input and does not claim complete OpenFormula conformance. The
normative source is the local
`part4-formula/OpenDocument-v1.4-os-part4-formula.html` member of
`3rdparty/specs/OpenDocument-v1.4-os.zip`:

| Source | SHA-256 |
| --- | --- |
| archive `OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| HTML member | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |

The current evaluator snapshot used for this gap review was:

| File | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `a0b577f441ef56baa5a364bc45d7b9889af503173bfdb653044309f4d82ff789` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `35d441d137123348f7b57a7fbb222e960c7e4bc5137eeeec25077c45bf581187` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/scalar.rs` | `ee5767b1be57c4c10b3b347d7ce3638df8fa06fb4ac75cad1ec70c51c171b747` |
| `crates/litchi-ods/src/codec/formula/functions.rs` | `1d736271e9f9cf3743295f8895f184e32842640eaba48084dfafc9ca7333e2a1` |
| `crates/litchi-ods/docs/FEATURE_MATRIX.md` | `2a9b6e072c3facf8219921414df6ae4db64a82856441524b4e0a686942e01aa4` |
| `docs/adr/0005-io-memory-and-performance.md` | `34a6148a8fe77b3e90212810996667b654e25209fb49c56aaecd8d8fbe83f770` |

## Function set and signatures

The complete family is `CEILING`, `INT`, `FLOOR`, `MROUND`, `ROUND`,
`ROUNDDOWN`, `ROUNDUP`, and `TRUNC` (§6.17.1–§6.17.8). The parameter types in
the source are significant; they must not all be normalized to an integer
helper.

| Function | ODF syntax | Return | Additional constraints |
| --- | --- | --- | --- |
| `CEILING` | `CEILING(Number N [; [Number Significance] [; Number Mode]])` | `Number` | `N` and `Significance` have the same sign when neither is zero |
| `INT` | `INT(Number N)` | `Number` | none |
| `FLOOR` | `FLOOR(Number N [; [Number Significance] [; Number Mode]])` | `Number` | `N` and `Significance` have the same sign; zero is handled by its semantics |
| `MROUND` | `MROUND(Number A; Number B)` | `Number` | none stated |
| `ROUND` | `ROUND(Number X [; Number Digits = 0])` | `Number` | none |
| `ROUNDDOWN` | `ROUNDDOWN(Number X [; Integer Digits = 0])` | `Number` | none |
| `ROUNDUP` | `ROUNDUP(Number X [; Integer Digits = 0])` | `Number` | none |
| `TRUNC` | `TRUNC(Number A; Integer B)` | `Number` | none; the prose also permits an absent `B` |

Thus `CEILING`/`FLOOR` significance and `MROUND`'s two operands are finite
`Number` values and may be fractional. `Mode` is also a `Number`: its only
normative distinction is zero versus nonzero, so a non-integer such as `0.5`
must not be truncated to zero. `ROUND`'s `Digits` is a `Number`, unlike the
`Integer` `Digits`/`B` parameters of `ROUNDDOWN`, `ROUNDUP`, and `TRUNC`.

## Normative semantics

* `INT` returns the greatest integer less than or equal to `N` (toward negative
  infinity).
* `ROUND` rounds to the nearest multiple of `10^-Digits`; halfway values round
  away from zero. Positive digits retain decimal places and negative digits
  round to the left of the decimal point.
* `ROUNDDOWN` rounds `X` toward zero and `ROUNDUP` rounds it away from zero at
  the requested decimal position. `TRUNC` truncates toward zero at its
  requested position.
* `CEILING` with omitted or syntactically empty significance uses `+1` when
  `N` is non-negative and `-1` when `N` is negative. With omitted or zero
  `Mode`, it rounds toward positive infinity to the smallest signed multiple
  not below `N`. With nonzero `Mode`, it rounds `abs(N)` away from zero to a
  multiple of `abs(Significance)` and reapplies the sign.
* `FLOOR` has the same significance default. With omitted or zero `Mode`, it
  rounds toward negative infinity to the greatest signed multiple not above
  `N`. With nonzero `Mode`, it rounds `abs(N)` toward zero to a multiple of
  `abs(Significance)` and reapplies the sign.
* If `N` or `Significance` is zero, `CEILING` and `FLOOR` return zero. Nonzero
  `N` and `Significance` with opposite signs violate the stated constraint and
  return a formula error under the selected profile.
* `MROUND` chooses a multiple of `B` nearest to `A`. If two multiples are
  equally distant, it chooses the numerically greater multiple. There is no
  same-sign constraint in §6.17.4. Consequently `MROUND(-5; 2)` is `-4`, and
  `MROUND(5; -2)` is `6`; the tie rule is numeric ordering, not greater
  absolute value.

The examples above also lock the negative-mode distinction:
`CEILING(-5.3; -2; 0) = -4`, `CEILING(-5.3; -2; 1) = -6`,
`FLOOR(-5.3; -2; 0) = -6`, and `FLOOR(-5.3; -2; 1) = -4`.

## Conversion and optional Empty policy

The common evaluation rules in §§3.2.3, 6.2, 6.3.1, and 6.3.5 require
implicit conversion to the declared parameter type and propagation of formula
errors. This profile uses finite `f64` Numbers, converts Logical to `0`/`1`,
uses the existing locale-independent decimal Text-to-Number policy, and
converts an evaluated Empty value to numeric zero. Conversion failure is
`#VALUE!`; non-finite input or a non-finite final result is `#NUM!`.

Syntactic omission is distinct from an explicit empty parameter under §5.6.
For `CEILING` and `FLOOR`, the special “empty parameter (two consecutive
semicolons)” rule applies to significance and selects the sign-dependent
default above. A blank cell supplied as an expression is an evaluated Empty
value and therefore converts to Number zero in this profile, producing the
zero-significance result. Treating a blank cell as omitted is a possible host
extension, but must not be conflated with the syntax rule. An empty `Mode`
converts to zero and has the same result as an omitted mode. Omitted or empty
`ROUND`, `ROUNDDOWN`, `ROUNDUP`, and `TRUNC` digit parameters use zero.

For the `Integer` parameters, §6.3.6 leaves conversion of a non-integer Number
implementation-defined unless the function specifies a direction. This
profile chooses truncation toward zero for `ROUNDDOWN`, `ROUNDUP`, and `TRUNC`
digits. `ROUND` retains its declared `Number Digits` and does not silently
apply that integer conversion: finite fractional `Digits` are accepted and
used as the exponent in `10^-Digits`. A kernel that cannot support fractional
decimal exponents must return a documented typed refusal rather than silently
truncating them. `Mode` never uses integer conversion.

The `TRUNC` syntax table requires `B`, while its semantics explicitly says that
an absent `B` means zero. The profile accepts one or two arguments and treats
the omitted argument as zero, recording the prose/syntax inconsistency rather
than rejecting the documented one-argument form. Invalid arity otherwise
returns `#VALUE!`. A `MROUND` divisor of zero has no mathematical multiple
definition; this profile returns `#DIV/0!`, an explicit choice because the
section lists no constraint or error kind.

## Resource, numerical, and test boundaries

The scalar kernels should allocate no heap storage. They must check the
retained execution context and charge the existing evaluator work budget for
each operation and conversion. Matrix-mode iteration and reference projection
must retain the existing finite shape, stack, storage, cancellation, and
reference-cell limits; they must not turn a scalar rounding call into an
unbounded materialization.

Power-of-ten scaling, quotient formation, and multiplication need checked
paths: an intermediate overflow or underflow is not by itself proof that the
final rounded result is non-finite. Conversely, a non-finite final result must
become `#NUM!`. Avoid unchecked casts of large or non-finite digit values. This
profile canonicalizes a zero result, including `-0.0`, to `+0.0`.
The ODF numerical model is implementation-defined (§3.6), so binary `f64`
tie behavior and decimal exactness are profile limits, not a reason to claim
universal decimal conformance.

For integer decimal digit counts, the implementation interprets the finite
input through its shortest round-trip decimal representation. It quantizes
that representation using a bounded integer coefficient and decimal exponent,
then converts the selected decimal result back to `f64`. Fixed-size formatting
buffers and checked integer arithmetic avoid heap allocation. This explicit
profile keeps ordinary decimal inputs such as `0.07` stable while retaining
the distinction from an adjacent input such as `0.07000000000000002`.
It does not snap quotients using an epsilon or truncate inputs to an arbitrary
number of significant figures. Fractional `ROUND` exponents and the
significance-based functions retain their documented binary arithmetic.

Tests should cover every function, positive and negative values, zero and
negative significance, all mode states, fractional Number significance and
Mode, ties on both signs, negative MROUND divisors, omitted/explicit-empty
parameters versus blank-cell parameters, integer-argument truncation, the
one-argument TRUNC form, formula-error precedence, non-finite and extreme
finite values, scalar implicit intersection, matrix broadcasting, cancellation,
work/storage refusals, and release of any retained reservations. Expected
values should come from an independent decimal/reference oracle rather than a
candidate-generated checksum.

The profile intentionally does not add locale, `HOST-PRECISION-AS-SHOWN`,
worksheet lookup, date epochs, or external services. Those host choices belong
to later evaluator families and must not be smuggled into this scalar batch.
