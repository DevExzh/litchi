# ODS financial contract semantic review

Status: **PASS for the current contract's semantic profile; implementation and
regression validation remain pending**. This review covers the twenty
cash-flow and annuity functions in the current contract. It is not a
production, test, numerical, resource, or native-parity claim. No
production, test, Cargo, or contract files were changed by this review.

The reviewed contract is [`contract.md`](contract.md), SHA-256
`fbe2283886326582b0e34fd277b38e6fb3e1bdeb2ecfbec5522da0f6f1467e4b`.
The primary source is the repository-local ODF 1.4 archive and formula member:

| Input | SHA-256 |
| --- | --- |
| `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| `part4-formula/OpenDocument-v1.4-os-part4-formula.html` | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |

I checked §§3.3, 4.3.1, 4.3.7, 4.5–4.7, 4.10, 4.11.12, 6.3.2,
6.3.4–6.3.9, 6.4.6, 6.12.1, 6.12.11–6.12.12, 6.12.19–6.12.21,
6.12.23–6.12.25, 6.12.27–6.12.30, 6.12.35–6.12.37, 6.12.41–6.12.42,
6.12.44, 6.12.51–6.12.52, and the shared POWER definition in §6.16.46.

The current amendment resolves the earlier findings: it records the exact
MIRR MathML equation; admits scalar Number, Text, and Logical values in
DateSequence through Number conversion; declares `PPMT`'s `Type` as Number
and keeps it out of the truncating Integer controls; applies the selected
Type-1 IPMT beginning-payment recurrence; and defines strict XIRR original
slot masks. The XIRR profile retains admitted formula Errors, drops slots
skipped on both sides, and generates `#VALUE!` after complete scans when only
one side skips a slot, avoiding silent cash-flow/date reassignment.

## Resolved RRI profile

### RRI negative-ratio POWER profile

§6.12.44 gives `RRI = (Fv / Pv)^(1 / Nper) - 1` and only constrains
`Nper > 0`. The shared POWER rule in §6.16.46 says that `POWER(A; B)` for
`A <= 0` and non-integer `B` is implementation-defined. The contract now
chooses the repository's shared real POWER profile for that boundary.

The deterministic recommendation is to follow the repository's shared real
POWER behavior: for a negative ratio, a finite exponent `1/Nper` is accepted
only when it is exactly integral and is evaluated with checked signed integer
power; otherwise the result is `#NUM!`. Thus
`RRI(0.5;100;-100)` is zero, while `RRI(3;100;-800)` is `#NUM!`; the contract
explicitly excludes an additional odd-root extension. Zero and non-finite
intermediate domains are also mapped without exposing NaN or infinity.

## Resolved decisions checked against the primary source

* **MIRR:** the source names plain `Array`, not `Reference|Array` or a
  sequence pseudotype. The contract's scalar reference refusal, matrix Array
  admission, Text/Empty omission, Logical 0/1 conversion, retained errors,
  same-position sign masks, and exact MathML equation are explicit profile
  choices and are internally consistent.
* **DateSequence and XIRR:** §6.3.9 inherits the scalar NumberSequence
  conversion rule. The contract now states scalar Text/Logical conversion and
  the strict original-slot mask rule, including complete scans and formula
  error precedence.
* **XNPV:** §6.12.52's equal-count row-wise pairing for different rectangular
  geometries is preserved, with Values-first scanning and generated
  `#VALUE!` for non-Number non-error elements.
* **Controls and CUMPRINC:** Number-declared controls use exact 0/1; only
  CUMIPMT/CUMPRINC Integer controls truncate. CUMPRINC retains its literal
  per-term PPMT preconditions and reversed-empty interval behavior.
* **Amortization:** IPMT, PPMT, and ISPMT equations are labeled repository
  profiles because their ODF clauses do not print equations. The Type-1
  period-one zero-interest and post-payment recurrence now agrees with the
  selected profile.
* **Solvers:** the bounded transformed-coordinate IRR/RATE/XIRR profile is
  recorded, including branch bounds, bracketing, tie-breaking, caps, and
  failure mapping. It is no longer an open contract item.

## Disposition

**PASS for contract semantics.** All reviewed normative/profile decisions are
now explicit, including the negative-ratio RRI boundary. Implementation,
focused regression tests, resource evidence, and source-freeze evidence remain
outside this contract review and are not claimed here.
