# ODS rounding implementation review

This is the final independent review of the frozen OpenFormula 1.4 Part 4
§6.17 candidate. The checked normative artifact is the local
`part4-formula/OpenDocument-v1.4-os-part4-formula.html` member of
`3rdparty/specs/OpenDocument-v1.4-os.zip`:

* archive SHA-256:
  `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`;
* HTML member SHA-256:
  `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`;
* section anchors: `a_6_17_1_CEILING` through `a_6_17_8_TRUNC`, extracted
  HTML lines 12169–12274.

The reviewed production and test source hashes are:

| File | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation/rounding.rs` | `f0bdde7dad0e5d07fe2edb493be3bb8451d157b9617af87e10c34d1a1b7cdf46` |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `6eab8e35a69fa806f43c2be8fcd7a338c93ddef55837819223ac20a1da685fa3` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/scalar.rs` | `fcc077c551bba2ce5fa66207b8b952541f33c9f4e2e48f67668c4b211d806388` |
| `crates/litchi-ods/tests/ods_formula_rounding_evaluation.rs` | `1866fc167779c08d6421ec247f18ea5b84f7ac47b73ecd98847ce0c037249604` |
| `crates/litchi-ods/tests/ods_formula_rounding_arrays.rs` | `9478f81f634ee52514f6f51a3e34019ac87ce466f64ba0f3cce9dea8ba66a51b` |
| `docs/report/spec-gap-validation-evidence/ods-formula-rounding/contract.md` | `48746b3c7e11bd141ec05693729a4fa11cf82f2ac678f50d75a16d62933493b4` |

The focused validation completed with 19 scalar tests and 9 matrix/value tests:

```text
cargo test -p litchi-ods --test ods_formula_rounding_evaluation -- --nocapture
19 passed
cargo test -p litchi-ods --test ods_formula_rounding_arrays -- --nocapture
9 passed
```

The root gate receipt additionally reports 1,168 ordinary ODS tests, five
doctests, strict all-target Clippy, rustdoc, formatting, and the rounding
profile verifier passing against the hashes above.

## Current correctness assessment

The implementation follows §6.17.4's nearest-multiple definition and its
numerically greater tie rule. `mround` compares the lower and upper
mathematical distances before constructing the upper candidate. A selected
upper candidate that overflows therefore returns `#NUM!`; an overflowing
upper neighbor is not substituted with a farther lower neighbor. The checked
boundary expectations are:

| Formula | Result |
| --- | --- |
| `MROUND(1.1e308;1e308)` | `1e308` |
| `MROUND(1.7e308;1e308)` | `#NUM!` |
| `MROUND(0;0)` | `#DIV/0!` under this profile |

Integer digit counts use a fixed 32-byte scientific representation of the
finite input's shortest round-trip decimal value. The coefficient and decimal
exponent are checked, quotient/remainder decisions are integer operations,
and the selected decimal result is parsed back to `f64`. The integer path
does not use epsilon snapping and does not silently fall back to binary
arithmetic when formatting or checked arithmetic fails; the caller reports a
typed Number error. This policy keeps ordinary decimal values and adjacent
binary values separate. The regression vectors include:

* `ROUNDDOWN(0.3;1) = 0.3` and `TRUNC(0.3;1) = 0.3`;
* `ROUNDUP(0.07;2) = 0.07`, while
  `ROUNDUP(0.07000000000000002;2) = 0.08`;
* `ROUNDDOWN(1.15;2) = 1.15` and
  `ROUNDDOWN(1.1499999999999997;2) = 1.14`;
* `ROUNDUP(1e-20;0) = 1` and `ROUNDUP(-1e-20;0) = -1`.

The implementation uses a separate checked binary scale path for fractional
digits. This is the explicit profile interpretation of the §6.17.5 `Number`
signature and `10^-Digits` definition. The same paragraph's statement that a
Number with `Digits <= 0` always produces an integer is recorded as a
specification tension; the profile and README do not silently claim both
interpretations at once. The boundary vectors cover finite `-308.1` scales,
subnormal `323.1` scales, overflowing integer scales, and digits beyond the
representable range. Non-finite final results become `#NUM!`, and intermediate
overflow does not force an error when the selected result remains finite.

All result publication passes through the zero canonicalization boundary, so
`-0.0` becomes `+0.0`. Tests inspect the zero sign bit for decimal, MROUND,
and CEILING cases. Omitted and explicit-empty optional arguments remain
distinct from blank referenced cells. The suites cover omitted one-argument
`TRUNC`, scalar reference intersection, matrix broadcasting, lazy branches,
formula-error precedence, work/cancellation refusal, storage refusal, and
reservation release.

## Resolved review history

Earlier candidate snapshots had four concrete issues. They are retained here
as resolution evidence rather than open findings:

1. `Err(upper) => lower` in MROUND could return a farther finite neighbor;
   distance selection now precedes checked upper construction.
2. Division by an inexact `10^-digits` moved values such as `0.3` below an
   integer quotient; the shortest-decimal integer path replaces that quotient
   and avoids the unsafe `near_integer` tolerance.
3. Several negative directed results exposed `-0.0`; `push_result` now
   canonicalizes both zero signs and tests assert the bit pattern.
4. The earlier evidence omitted storage/intersection/omitted-TRUNC cases and
   had only historical 14-scalar/8-matrix receipts; the frozen suites and
   profile now contain the expanded coverage and the retained source hashes.

The normative and resource choices are recorded in the local
[rounding contract](contract.md), [performance evidence](performance.md),
[array/reference review](../ods-formula-array-reference-evaluation/spec-review.md),
and [ADR 0005](../../../adr/0005-io-memory-and-performance.md). No material
correctness blocker remains in the frozen snapshot reviewed above.
