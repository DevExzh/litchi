# Independent review of the ODF 1.4 discrete-math batch

This review covers the eleven functions in ODF 1.4 Part 4 §6.16:
`COMBIN`, `COMBINA`, `FACT`, `FACTDOUBLE`, `GCD`, `LCM`,
`MULTINOMIAL`, `EVEN`, `ODD`, `DELTA`, and `GESTEP`. It is
independent of the production and test changes. The normative source identity
and the selected profile are recorded in [contract.md](contract.md).

The review accepts the frozen implementation profile described below. The
acceptance applies to the source snapshot identified by the gate freeze and to
the read-only formula evaluators. It does not claim that native cached
formula results are normative or that the targeted oracle proves every
possible binary64 input.

## Frozen source and receipts

The gate freeze is based on commit `782339a2c467ba77def74839066ba9f9ef7eb67d`.
The relevant frozen source digests are:

| File | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation/discrete.rs` | `0b4a6f85b2e1c9131f4a9d62752da0bc59549a39cf8e320ea5a912240f437af9` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/aggregate.rs` | `ef77863ad9fed2a6f869df0ed5f819ed30ba1a5163b9017acd4dd9f3b7fb3874` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/scalar.rs` | `36fa84a7821639434b4f2a47db9025b16bcec937ab567cff29f2a9ea809ddbdc` |

The isolated gate receipt reports stable sources and seven successful checks:

- `cargo test --locked -p litchi-ods`: 1,271 passed, 0 failed, 0 ignored;
- all-target Clippy with `-D warnings`;
- rustdoc with warnings denied;
- crate format and selected batch format;
- crate-boundary checks; and
- `git diff --check`.

The retained numerical oracle contains 361 observations across all eleven
functions, including exact represented-binary64 fractional `MULTINOMIAL`
cases. Both evaluator APIs match those observations at 0 ULP where the
profile publishes a result; rows marked with a selected profile error are
checked as errors. The native reproduction retains 77 cached observations
from 11 pinned LibreOffice source files and records 15 excluded host/profile
variances. Those observations corroborate ordinary cases and document host
behavior; they do not override the contract choices.

## Normative disposition

The following choices resolve the source's explicit permissions and the
implementation-defined conversion points:

| Topic | Accepted disposition | Basis |
| --- | --- | --- |
| Integer conversion | `INT`, floor toward negative infinity | §4.11.5 defines `Integer` as `INT(X)`; the COMBIN and COMBINA clauses also name `INT` explicitly. |
| `MULTINOMIAL` fractions | Sum raw admitted values first; apply Integer conversion independently through the numerator and each denominator `FACT` | §6.16.43 gives the nested `FACT(a1+...+an)/(FACT(a1)*...)` expression and gives no pre-truncation rule. |
| `LCM` fractions | Reject a selected fraction with `#NUM!` | §6.16.38 states `INT(X)=X` and `X>=0` as constraints. The later `INT` wording does not remove that constraint. |
| `COMBINA(0;0)` | Return `1`; reject `N=0,M>0` | The displayed binomial is undefined at zero/zero; the profile selects the empty-multiset identity while retaining `N>=M`. |
| `COMBINA(N<M)` | Reject with `#NUM!` | The standard permits, but does not require, a non-error extension. This profile does not enable it. |
| `GESTEP` non-Number sentence | Use the common Number conversion; conversion failures are `#VALUE!` | The Number signature and §6.2 conversion rules accept numeric Text, Logical, and Empty reference values. |
| all-zero `GCD` | Return `0` | §6.16.36 permits either Error or `0`; the profile selects `0`. |
| empty sequence | `GCD`/`LCM` `#NUM!`; `MULTINOMIAL` `1` | GCD cannot satisfy its positive-member constraint, LCM has no source identity, and the empty multinomial sum/product gives `FACT(0)/1`. |
| large parity | Publish only a finite result that preserves the requested even/odd parity; otherwise `#NUM!` | The exact candidate is checked after binary64 rounding. Thus `ODD(2^53)` is a profile error rather than an even rounded result. |

The ordinary native variance set is intentionally excluded from conformance
claims: LibreOffice caches `COMBINA(0;0)` as `0`, accepts fractional
GCD/LCM inputs differently, and has other host-specific behavior around
fractional `MULTINOMIAL` and text conversion.

## Conversion and sequence review

The frozen implementation preserves the distinction between `NumberSequence`
and `NumberSequenceList`:

- GCD and LCM admit `ReferenceList` values. Areas are consumed in reference
  list order and cells in each area in row-major order.
- MULTINOMIAL has `NumberSequence`, so an explicit `ReferenceList` is a
  structural `#VALUE!` mismatch even though one rectangular reference is
  admitted. Rejected areas are not dereferenced to search for hidden provider
  errors. Direct scalar or inline-array formula Errors already materialized in
  other arguments retain source-order precedence.
- Referenced Empty, Text, and distinguished Logical cells are omitted from a
  NumberSequence. Referenced Number and formula Error cells are admitted.
- Inline arrays are value arrays and use the scalar bridge: Empty becomes zero,
  Logical becomes zero or one, numeric Text is parsed, malformed Text and
  Missing are `#VALUE!`, and formula Errors propagate.
- Scalar Empty uses the Number bridge and becomes zero. An omitted optional
  argument receives its default; a supplied empty slot is Missing and remains
  `#VALUE!`.

The scalar evaluator rejects references and arrays at its existing typed API
boundary. The value evaluator consumes scalar functions elementwise and keeps
the existing array/reference shape and broadcast rules. No path uses a formula
cache as a reference read.

The scalar and value sequence paths both keep the first formula Error separate
from generated conversion/domain errors. They scan all eagerly evaluated
arguments and sequence members so a later formula Error can supersede an
earlier generated error. If no formula Error exists, the first generated error
in argument, area, cell, or inline-array traversal order is returned. Typed
provider, cancellation, resource, and allocation failures remain evaluator
failures and are not converted into formula Errors.

## Exact arithmetic and boundary review

The shared `DiscreteFold` uses a fixed 1024-bit unsigned limb state for all
finite binary64 integer operands. It computes factorials, combinations,
GCD/LCM, and multinomial integer ratios exactly, then performs one
round-to-nearest, ties-to-even conversion at the Number boundary. Overflow is
reported as `#NUM!`; low bits are not discarded to make an intermediate fit.

`MULTINOMIAL` stores each raw finite non-negative binary64 operand's integer
part and an exact fractional residue at the binary64 minimum exponent. The
residue is accumulated in a fixed limb state, with each carry increasing the
raw-sum floor by one. The integer multinomial ratio uses the independently
floored denominator parts, and the final rising factors account for the
fractional carries. This gives the required profile results
`MULTINOMIAL(1.5;1.5)=6` and `MULTINOMIAL(1.4;0.6)=1` without relying on a
rounded binary64 intermediate sum.

The binomial kernel selects the smaller side and uses divide-before-multiply.
Its `r >= 1024` refusal is derived from `C(n,r) >= 2^r` when `n >= 2r`, so it
is a finite-result proof rather than a host argument cap. FACT and FACTDOUBLE
use the exact finite-result boundaries of their products. LCM reduces by the
exact GCD before multiplication and treats zero as absorbing; a zero after an
intermediate overflowing nonzero LCM restores the mathematically correct zero
result.

For wide GCD, the u64 fast path handles ordinary values. The fixed-width
Euclidean path checks for a zero divisor before `div_rem`; `div_rem` returns
a value for every nonzero fixed-width divisor. For positive remainders, every
two Euclidean steps reduce the current divisor below half its previous value:
if the next remainder is already at most half, this is immediate; otherwise
the following subtraction leaves less than half. A positive 1024-bit operand
therefore reaches zero in at most `2 * MAX_BITS` divisions. The loop includes
an extra terminal guard iteration. Its `None` and exhausted-loop branches are
unreachable under these checked invariants; the proof is why the bound is a
defensible fixed-width limit rather than a guessed threshold.

EVEN and ODD construct the exact required candidate before output. If the
finite binary64 rounding would change its required parity, the implementation
returns `#NUM!`; ordinary finite candidates are rounded once at the output
boundary. DELTA uses exact binary64 equality, and GESTEP uses the converted
binary64 `>=` comparison.

## Work, cancellation, and error accounting

Each sequence operand is admitted only after charging its bounded kernel work.
The charge invokes the execution check before the fixed-width operation. Work
units describe logical kernel operations rather than CPU instructions: the
limb arrays have fixed 32-limb width, the fractional residue has fixed width,
and all recurrences have mathematically derived finite bounds. The admission
therefore remains proportional to the number of sequence operands and
recurrence rounds without allowing an input value to request an unbounded
allocation or loop.

The scalar and resolver-backed value reducers use the same work estimate.
Reference traversal charges each cell before reading it, and structural
`ReferenceList` rejection occurs before provider reads. Formula errors do not
short-circuit eager traversal, while typed cancellation and resource failures
are returned immediately. The fixed-width kernels are bounded atomic units;
their precharge check provides the cancellation boundary and their maximum
limb/round widths are recorded in the implementation comments and contract.

## Final disposition

The source, test, documentation, native, oracle, and gate receipts agree with
the contract. The frozen Scope11 implementation is accepted for this selected
finite-binary64 profile. Future changes to conversion order, sequence shape,
limb width, error precedence, or work admission require a new normative review
and a new source freeze; the current receipts must not be reused silently.
