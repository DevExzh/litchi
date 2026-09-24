# ODF 1.4 financial cash-flow and annuity test plan

Status: planning only. This file defines the cases and independent evidence
needed before the twenty functions in the financial contract are advertised as
implemented. It contains no production or Rust test changes and makes no
validation claim.

The owned slice is `CUMIPMT`, `CUMPRINC`, `EFFECT`, `FV`, `FVSCHEDULE`, `IPMT`,
`IRR`, `ISPMT`, `MIRR`, `NOMINAL`, `NPER`, `NPV`, `PDURATION`, `PMT`, `PPMT`,
`PV`, `RATE`, `RRI`, `XIRR`, and `XNPV`. The remaining §6.12 security, coupon,
depreciation, and fraction functions stay outside this plan.

The first independent scalar oracle is deliberately narrower than that
twenty-function inventory. Its Decimal corpus covers these eleven finite,
non-iterative kernels: `EFFECT`, `FV`, `IPMT`, `ISPMT`, `NOMINAL`, `NPER`,
`PDURATION`, `PMT`, `PPMT`, `PV`, and `RRI`. The pending nine are
`CUMIPMT`, `CUMPRINC`, `FVSCHEDULE`, `IRR`, `MIRR`, `NPV`, `RATE`, `XIRR`,
and `XNPV`; they remain explicitly outside the scalar corpus and no all-twenty
coverage claim follows from it. The generator is
[`oracle_scalar.py`](oracle_scalar.py), and its generated vectors are
[`oracle-scalar-vectors.json`](oracle-scalar-vectors.json).

## Authority and source review

The primary source is the repository-local
`3rdparty/specs/OpenDocument-v1.4-os.zip`, member
`part4-formula/OpenDocument-v1.4-os-part4-formula.html`. The checked source
hashes are:

| Artifact | SHA-256 |
| --- | --- |
| Archive | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| Part 4 HTML member | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |

The direct entries are §§6.12.11–6.12.12, 6.12.19–6.12.21, 6.12.23–6.12.25,
6.12.27–6.12.30, 6.12.35–6.12.37, 6.12.41–6.12.42, 6.12.44, and
6.12.51–6.12.52. Shared conventions come from §6.12.1 and the conversion and
sequence rules in §§4.3.3–4.3.6, 4.11.1–4.11.2, 4.11.5, 4.11.7,
4.11.12–4.11.13, and 6.2–6.3.

The HTML entries provide signatures, constraints, prose, and MathML equations;
they do not provide a complete set of worked decimal examples. The positive
vectors below are therefore direct evaluations of those normative equations,
with exact or high-precision expected values independently calculated in the
oracle. The MIRR MathML is authoritative where flattened HTML text loses the
fraction or exponent layout. LibreOffice files and outputs are compatibility
evidence after the normative vectors, never the source of the contract.

The plan follows accepted ADR 0001 (typed, panic-free public behavior), ADR
0004 (bounded semantic values without archive/provider types), ADR 0005
(hierarchical work, storage, reference, cancellation, source, and fallible
allocation budgets), ADR 0006 (explicit providers, formula Errors as values,
typed failures preserved), ADR 0008 (compile-tested positive and negative
paths before support claims), and ADR 0023 (the ODS family owns its formula
tests and semantic boundary).

## Exact arity and baseline vectors

The harness will run every row in scalar evaluation and in resolver-backed
value evaluation where the formula is context-free. Numbers below are
independent expected values; approximate rows use the oracle tolerance recorded
with the vector. Inline-array examples use the repository formula grammar:
`|` separates rows and `;` separates columns. Traversal assertions use row
major order.

| Function | Arity | Baseline vector and expected result | Constraint/error vectors |
| --- | ---: | --- | --- |
| `CUMIPMT` | 6 | `CUMIPMT(0.1;2;100;1;2;0)` = `-15.238095238095238`; type `1` is checked separately | `Rate=0`, `Value=0`, `Start=0`, `End>Periods`, and `Type=2` must use the selected numeric constraint mapping |
| `CUMPRINC` | 6 | `CUMPRINC(0.1;2;100;1;1;0)` = `-47.619047619047619`; the full `End=Periods` case remains a pending reducer diagnostic | `Type=2`, reversed period bounds, and the selected per-term PPMT `Period>=Nper` `#NUM!` boundary |
| `EFFECT` | 2 | `EFFECT(0.1;2)` = `0.1025`; `EFFECT(0;12)` = `0` | Negative rate and zero/nonpositive payments are domain failures |
| `FV` | 3–5 | `FV(0.1;1;-110)` = `110`; `FV(0;10;-10)` = `100`; beginning-payment type is compared with end-payment type | Zero rate, `PayType=0/1`, invalid pay type, and non-finite intermediate cases |
| `FVSCHEDULE` | 2 | `FVSCHEDULE(100;{0.1|0.2})` = `132` | Empty/malformed schedule, formula Error member, and reference-list admission are tested |
| `IPMT` | 4–6 | `IPMT(0.1;1;2;100)` = `-10`; period 2 = `-5.238095238095238`; the selected beginning-payment profile gives `IPMT(0.1;2;2;100;0;1) = -100/21` | Period boundary, type 2, zero rate, and optional FV/Type slots |
| `IRR` | 1–2 | `IRR({-100|110})` = `0.1`; explicit guess `.2` converges to the same root | All-positive/all-negative, no root, multiple-root, non-finite guess, and bounded nonconvergence |
| `ISPMT` | 4 | `ISPMT(0.1;1;2;100)` = `-5`; period 2 = `0` under the selected conventional profile | Period 0, period beyond Nper, zero/negative/non-finite values, and sign behavior |
| `MIRR` | 3 | `MIRR({-100|110};0.1;0.1)` = `0.1` under the MathML equation; reducer execution remains pending | No positive or no negative value, Text/Empty omission, and the selected scalar Reference/ReferenceList shape refusal |
| `NOMINAL` | 2 | `NOMINAL(0.1025;2)` = `0.1`; `NOMINAL(EFFECT(0.1;12);12)` round-trips to `.1` | Effective rate `<=0` and compounding count `<=0` |
| `NPER` | 3–5 | `NPER(0.1;-110;100)` = `1`; `NPER(0;-10;100)` = `10` | Negative-rate Medium/Large profile, zero-rate branch, singular balances, and optional FV/Type |
| `NPV` | 2+ | `NPV(0.1;{-100|110})` = `0`; split `NPV(0.1;{-100|0};{110})` proves argument order | Missing sequence, multiple sequence-list areas, formula Error retention, and invalid rate/intermediate |
| `PDURATION` | 3 | `PDURATION(0.1;100;121)` = `2` | Rate `<=0`, current/specified value `<=0`, equal values, and overflow |
| `PMT` | 3–5 | `PMT(0.1;1;100)` = `-110`; `PMT(0;10;100)` = `-10` | `Nper<=0`, zero/nonzero rate branches, and PayType 0/1 |
| `PPMT` | 4–6 | `PPMT(0.1;1;2;100)` = `-47.619047619047619` | `Rate<=0`, `Present<=0`, Period `<=0` or `>=Nper`; `Period>=Nper` is the selected `#NUM!` boundary |
| `PV` | 3–5 | `PV(0.1;1;-110)` = `100`; `PV(0;10;-10)` = `100` | Zero branch, PayType 0/1, singular/non-finite results |
| `RATE` | 3–6 | `RATE(1;-110;100)` = `0.1`; explicit guesses `.1` and `.2` are compared | `Nper<=0`, no root, multiple root, negative root, non-finite guess, and iteration bound |
| `RRI` | 3 | `RRI(2;100;121)` = `0.1`; `RRI(1;100;110)` = `.1` | Negative ratio with exact binary64 reciprocal integer (`Nper=1` or `.5`), `RRI(3;100;-800) -> #NUM!`, zero bases, and overflow |
| `XIRR` | 2–3 | `XIRR({-100|110};{43831|44196})` = `0.1` because the dates differ by 365 days | Unequal lengths, missing sign, date errors/order, guess/nonconvergence, and direct ReferenceList refusal |
| `XNPV` | 3 | `XNPV(0.1;{-100|110};{43831|44196})` = `0`; a nonannual date interval checks the `days/365` exponent | Rate `<=-1`, unequal lengths, non-number elements, dates before the first date, and list/shape refusal |

The sign convention is checked through identities rather than only isolated
outputs. For the two-period loan at 10%, the oracle records:

* `PMT(0.1;2;100)` = `-57.619047619047619`;
* `IPMT` periods 1 and 2 sum to `-15.238095238095238`;
* `CUMIPMT` period 1 plus `CUMPRINC` period 1 equals the period-1 payment;
  the full-range CUMPRINC case reaches the selected PPMT `Period>=Nper`
  `#NUM!` boundary and remains a pending reducer case.

The zero-rate equations are separate vectors. They must not be evaluated by a
nonzero-rate formula that divides by zero. Declared `Integer` slots—CUM
`Start`/`End`/`Type`, EFFECT `Payments`, NOMINAL compounding periods, PMT
`Nper`, and PPMT `Period`/`Nper`—include positive and negative fractional
values so truncation toward zero is observable and checked before range
validation. Number-valued `Nper` slots in FV/IPMT/NPER/RATE and Number-valued
`Type`/`PayType` controls in the FV/IPMT/NPER/PMT/PPMT/PV/RATE signatures are
tested as exact 0/1 selectors and with nonintegral values such as 0.5 and 1.9;
they are not covered by a blanket Integer-truncation assertion.

The scalar corpus records `PPMT(0.1;2;2;100)` as the selected `#NUM!`
precondition boundary. `CUMPRINC(0.1;2;100;1;2;0)` remains a pending
sequence-reducer diagnostic because its result depends on the reducer's
per-term scan and typed/error precedence.

## Arity, omission, and explicit-empty matrix

Each function gets one below-arity and one above-arity case where the arity
permits it. Required missing arguments produce the established formula value;
too many arguments are rejected at the function boundary. `NPV` gets a missing
value-list case and a multiple-list case because its arity is open on the
right. `NOW`/`TODAY` are outside this slice.

Optional defaults are tested against explicit values and explicit empty AST
slots. The expected profile is recorded per vector rather than silently
assuming that an empty slot is omission:

| Family | Omitted form | Explicit-empty forms |
| --- | --- | --- |
| `FV`, `NPER`, `PMT`, `PV` | `FV(.1;1;-110)`, `PMT(.1;1;100)` | `FV(.1;1;-110;;)`, `PMT(.1;1;100;;)` and explicit `0`/`1` controls |
| `IPMT`, `PPMT` | `IPMT(.1;1;2;100)` | `IPMT(.1;1;2;100;;)`, `PPMT(.1;1;2;100;;)` |
| `RATE` | `RATE(1;-110;100)` | Empty FV, PayType, and Guess slots separately, including `RATE(1;-110;100;;; )` |
| `IRR` | `IRR({-100|110})` | `IRR({-100|110};)` and explicit guesses |
| `XIRR` | `XIRR({-100|110};{43831|44196})` | `XIRR({-100|110};{43831|44196};)` and explicit guesses |

The test fixture also distinguishes a syntactically empty slot from a resolver
cell whose value is `Empty`. Required arguments use both forms to confirm that
missing required input is not silently replaced by zero. For all optional
parameters, the test records whether the repository's existing Empty-to-Number
bridge applies or the function returns its selected formula error; this is a
profile decision that must be frozen before numerical receipts are accepted.

## Independent oracle and numerical profile

The scalar oracle is the small independent Python implementation in
`oracle_scalar.py`; it imports no Rust evaluator and calls no spreadsheet host.
It:

1. evaluates hand-authored scalar arguments with `Decimal` precision of 100
   digits after converting each formula literal through the evaluator's
   binary64 input profile;
2. uses explicit signed-integer power handling for finite negative-base FV and
   RRI cases;
3. applies the selected exact Number Type/PayType controls and
   truncate-toward-zero Integer slots; and
4. emits the expected finite value, formula error, absolute/relative
   tolerance, and vector identity without depending on the implementation's
   intermediate representation.

The generated scalar rows use zero absolute tolerance and a relative tolerance
of `1e-12` without a unit-scale floor, so a small nonzero target cannot be
accepted as zero merely because its magnitude is below a fixed absolute
threshold.

The selected RRI profile admits a negative ratio only when the binary64
reciprocal of `Nper` is an exact positive integer. The corpus covers
`Nper=1` (odd exponent), `Nper=0.5` (even exponent and result zero), and the
explicit refusal `RRI(3;100;-800) -> #NUM!`; it does not infer a real root from
the mathematical cube root when the profile rejects that input.

The pending nine require separate sequence traversal and/or deterministic root
solver profiles. They are listed in the corpus for scope accounting only and
are not evaluated by this scalar generator.

The oracle records exact decimal values for zero-rate, integer, and simple
ratio rows. For nonzero annuity rows it records a binary64 target plus a stated
absolute and relative tolerance. Root functions use a tighter residual check
(`abs(NPV)`) in addition to the output-rate tolerance. A result that is finite
but selects a different valid root is not silently accepted; multiple-root
vectors are separate profile cases and require a documented root-selection
rule.

The scalar oracle pass uses the equations from the local HTML and the selected
repository profiles:

* `EFFECT` and `NOMINAL` are reciprocal compounding relations;
* `FV`, `PV`, `PMT`, and `NPER` share the ordinary/due annuity balance with
  an explicit zero-rate branch;
* `IPMT` and `PPMT` use the selected amortization profile, including the
  beginning-payment divisor after period one;
* `PDURATION` and `RRI` use the logarithmic and power equations; and
* `ISPMT` uses `Rate * Pv * (Period / Nper - 1)` because §6.12.25 supplies no
  equation.

The scalar corpus supplements fixed vectors with the algebraic identities that
are safe for this slice: EFFECT/NOMINAL round trips, PV/FV reversal at the
same rate and term, PMT balance closure, and the IPMT/PPMT payment split.
Sequence and root-solver metamorphic checks remain pending with their nine
functions and never become an implicit all-twenty claim.

## Scalar, resolver, array, and projected evaluation

The semantic target will maintain one resolver fixture with stable source
version and cancellation controls. It contains numbers, logicals, text,
explicit empties, formula Errors, and out-of-domain numbers in several sheets.
Every scalar literal vector is run through the scalar evaluator and the value
evaluator in scalar mode. A scalar evaluator given a reference or a sequence
requiring a resolver must return typed `Unsupported(Reference)` rather than
inventing an origin.

The value target covers these shapes for each sequence-bearing family:

* inline one-dimensional and rectangular arrays in row-major order;
* one rectangular Reference with distinct values per cell;
* a three-dimensional Reference whose sheet/area order is observable;
* a ReferenceList where the pseudotype admits it (`NPV`); and
* a known wrong ReferenceList where the pseudotype rejects it (`IRR`,
  `FVSCHEDULE`, `XIRR`, `XNPV`, and the selected `MIRR` boundary).

`NPV(.1;[.A1:.A2]~[.B1:.B2])` must consume the first area before the second,
while `IRR([.A1:.A2]~[.B1:.B2])` must return the selected shape/type error with
zero resolver reads. A computed child such as an `INDIRECT`-produced reference
may be evaluated until its type is known; a statically known mismatch must not
read any cells.

Projected scalar consumers are tested with formulas such as:

```text
=IF({TRUE|TRUE};NPV(0.1;[.A1:.A2]);0)
=IF({TRUE|FALSE};FVSCHEDULE(100;[.B1:.B2]);0)
=IF({TRUE|TRUE};XNPV(0.1;[.C1:.C2];[.D1:.D2]);0)
```

The selected result shape, complete sequence reads, and per-position scalar
controls are asserted. A scalar financial reducer remains scalar at each
projected coordinate even when its sequence argument is rectangular. A
computed scalar criterion such as `SUM(MUNIT(...))` remains position-sensitive
and is not admitted to a cache entry that assumes a complete invariant
argument.

Formula Errors inside sequences are retained while the complete admitted
sequence is scanned. A later provider, cancellation, source-version, or
resource failure supersedes the retained formula Error. The test fixture logs
read order and verifies that a selected NPV/FVSCHEDULE/IRR/XIRR sequence does
not stop at the first formula Error.

## Error and constraint coverage

Every function receives a formula Error in its first argument and, where it
has controls or sequences, in a later argument. The expected value is the
original formula Error in source order. `IFERROR` and `IFNA` are wrapped around
both formula-error and typed-provider cases: only the formula-error value may
be handled.

Constraint vectors are grouped so each named condition is observable:

* CUMIPMT: positive rate/value, ordered integer period bounds, and Type 0/1;
* CUMPRINC: Type 0/1 plus the contract's deliberate distinction from the
  CUMIPMT positivity/period wording;
* EFFECT/NOMINAL: nonnegative or positive rates and positive integer counts;
* PMT/PPMT/RATE/NPER: positive period requirements, negative fractional
  truncation, zero rate, and negative-rate Medium/Large behavior;
* PDURATION/RRI: positive base/rate constraints and checked logarithm/power
  domains;
* IRR/RATE/XIRR: no root, multiple roots, guess sensitivity, nonfinite guess,
  and bounded nonconvergence;
* MIRR: at least one positive and one negative value, with Text/Empty omitted;
* NPV/FVSCHEDULE: empty or malformed sequences and source-order errors; and
* XIRR/XNPV: equal-size vectors, date ordering, non-number members, and the
  `rate > -1` XNPV boundary.

The exact formula error category for a violated financial constraint is still
an open contract decision. The tests will name the expected category only
after that mapping is frozen; they must not infer it from LibreOffice.

## Resource, cancellation, source, and cache evidence

The limits target uses a resolver that records every read and exposes typed
failures at a selected cell. The planned cases are:

* a large NPV/FVSCHEDULE/IRR/XIRR Reference stream under a low reference-cell
  limit, proving failure before the next read and no retained cell vector;
* a ReferenceList shape refusal for `IRR`, `FVSCHEDULE`, `XIRR`, and `XNPV`,
  proving zero resolver reads when the mismatch is statically known;
* cancellation before and after a resolver read and during root iteration;
* source-version changes before publication and after a complete sequence;
* a formula Error followed by a typed provider failure, proving typed failure
  precedence;
* bounded work for long NPV/IRR/XIRR sequences and long root iterations;
* fallible output/sequence scratch reservation with release after both success
  and failure; and
* borrowed text numeric conversion without cloning an entire referenced cell.

One-pass reducers must stream admitted cells. Root solvers may retain bounded
numeric terms for repeated evaluation only after checked capacity reservation;
they must not retain borrowed cell objects or source text. Any result cache key
must include the complete sequence descriptor/values, all scalar financial
parameters, source identity/version, cancellation fence, and projected shape.
Computed sequence expressions are cacheable only when the complete descriptor
is stable. A cached scalar result must not erase position sensitivity of a
projected guess, Type, PayType, or MUNIT-derived control.

## Native and compatibility evidence

The normative oracle is primary. Native LibreOffice/ODF files will be used to
check ordinary compatibility vectors for FV/PV/PMT, EFFECT/NOMINAL, NPV/IRR,
MIRR, and XIRR/XNPV with matching date serials and sign conventions. Native
results are recorded with version, locale, formula text, and recalculation
provenance. A native disagreement triggers review of the local ODF equation or
the implementation profile; it does not silently change the expected vector.

The final evidence bundle must bind the exact contract hash, oracle source and
vector hashes, focused Rust target hashes, isolated Cargo receipt, and native
provenance. A green parser/function-name test is insufficient because the
functions may still dispatch to typed Unsupported.

## Open decisions to resolve before implementation evidence

The contract records the selected MIRR, CUMPRINC/PPMT, XIRR/XNPV sign and
shape, Type/PayType, and formula-error profiles. The remaining implementation
profile is deterministic root selection, convergence tolerance, maximum
iterations, and nonconvergence mapping for `IRR`, `RATE`, and `XIRR`. The
sequence/reducer functions in the pending-nine list also need their complete
scan, shape, and resource receipts before they can be advertised; their
absence from the scalar corpus is intentional.

No financial implementation or validation disposition should be reported until
the pending profiles, independent oracle, focused semantic/resource tests, and
source-freeze receipts are complete.
