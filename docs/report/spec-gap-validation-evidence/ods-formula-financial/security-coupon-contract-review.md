# Security and coupon contract source review

Status: **PASS for source transcription; HOLD for implementation support until
the contract's explicitly open profiles are separately accepted**. This is an
independent read-only comparison of the security/coupon contract with the
repository-local ODF 1.4 source. It makes no production, test, native-parity,
resource, performance, or implementation-support claim.

## Reviewed inputs

| Input | SHA-256 |
| --- | --- |
| `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| `part4-formula/OpenDocument-v1.4-os-part4-formula.html` | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |
| `ods-formula-financial/security-coupon-contract.md` | `f559f51b199c4f9ec655908b85889218f2066e32f80475372001dae2d4c9fc02` |
| `ods-formula-financial/scope.json` | `53099c347632d3be8332443cb85ad8bac44e4489ff39b2d6dc42d78754216d1d` |
| `ods-formula-date-time/contract.md` | `cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f` |

I read the HTML member directly from the ZIP, checked the 27 `security-and-
coupon` entries against `scope.json`, and inspected §§4.11.1–4.11.7,
6.3.2, 6.3.5–6.3.6, 6.3.15, 6.12.1, and each of §§6.12.2–6.12.10,
6.12.15, 6.12.18, 6.12.22, 6.12.26, 6.12.31–6.12.34, 6.12.38–6.12.40,
6.12.43, 6.12.47–6.12.49, and 6.12.53–6.12.55.

## Function-by-function result

All 27 inventory names and section numbers match `scope.json`. The contract's
signatures, return types, and explicit source constraints also match the
archive. The table records the equation status that was checked:

| Functions | Source comparison |
| --- | --- |
| `ACCRINT` | Signature, `Currency` result, ordering/positivity constraints, frequency set `{1,2,4,12}`, and per-period `Par * Coupon * YEARFRAC(...)` equation match. |
| `ACCRINTM` | Signature is correctly normalized to `ACCRINTM`; the source actually prints `ACCRINT` in the syntax line. Constraints and `Currency` result match; no displayed equation is supplied. |
| `AMORLINC` | Signature and constraints match, including `PurchaseDate <= FirstPeriodEndDate`, `Salvage >= 0`, `Period >= 0`, `Rate > 0`; period-zero, full-period, last-period, and beyond-life equations match. Equality of the two dates is source-defined as implementation-defined. |
| `COUPDAYBS`, `COUPDAYS`, `COUPDAYSNC` | Signatures, `Number` results, `Settlement < Maturity`, and frequency set `{1,2,4}` match. The source syntax prints `COUPDAYNC` for `COUPDAYSNC`; no closed equation is supplied. |
| `COUPNCD`, `COUPNUM`, `COUPPCD` | Signatures, return types, frequency sets, and the source's omission of an explicit ordering constraint for `COUPNUM` match. Schedule meanings match; no closed equations are supplied. |
| `DISC` | Signature, `Percentage` result, and only the printed `Settlement < Maturity` constraint match. The `((Redemption - Price) / Redemption) / YEARFRAC(...)` equation matches. |
| `DURATION`, `MDURATION` | Signatures preserve the source's `Date` spelling and `Number Frequency`; constraints and `Number` results match. `MDURATION = DURATION / (1 + Yield / Frequency)` matches. The source supplies no displayed cash-flow equation for `DURATION`. |
| `INTRATE` | Signature, `Number` result, and only the printed `Settlement < Maturity` constraint match. The `(Redemption - Investment) / Investment / YEARFRAC(...)` equation matches. The repeated `Basis Basis` spelling is correctly recorded as an anomaly. |
| `ODDFPRICE`, `ODDFYIELD`, `ODDLPRICE`, `ODDLYIELD` | Signatures, `Number` results, positive-value wording, frequency meanings `{1,2,4}`, and the one printed `ODDFYIELD` date ordering match. The source supplies no displayed price/yield equations. |
| `PRICE` | Signature, `Number` result, positivity wording, frequency set, `Bas` alias, and the displayed coupon-price equation match. The contract correctly identifies `AnnualYield`/`Yield` as the same input. |
| `PRICEDISC`, `PRICEMAT` | Signatures, `Number` results, and constraints match. `PRICEMAT`'s explicit `Rate = 0` and `AnnualYield = 0` result of `100` is retained. No displayed equation is supplied. |
| `RECEIVED` | Signature, `Number` result, positivity/order constraints, and `Investment / (1 - Discount * YEARFRAC(...))` equation match. |
| `TBILLEQ` | Signature, `Number` result, “less than one year beyond settlement” and positive-discount wording, Actual/360 `DSM`, and `365 * Discount / (360 - Discount * DSM)` equation match. |
| `TBILLPRICE`, `TBILLYIELD` | Signatures, results, maturity-window constraints, and positive Discount/Price wording match. Neither section supplies a displayed equation. |
| `YIELD` | Signature, `Number` result, positive Rate/Price/Redemption wording, and frequency meanings `{1,2,4}` match. No displayed inverse/solver equation is supplied. |
| `YIELDDISC` | Signature, `Number` result, positive Price/Redemption wording, and `(Redemption / Price - 1) / YEARFRAC(...)` equation match. |
| `YIELDMAT` | Signature, `Number` result, and positive Rate/Price wording match. No displayed equation is supplied. |

The source-defined frequency restrictions for the four irregular functions and
for `YIELD` appear in their parameter descriptions rather than the short
`Constraints` sentence. The contract includes them without turning the
source's positive “should” language into a stronger normative inequality.

## Shared profile comparison

The contract correctly carries over the archive's following points:

* §6.12.1's annuity terminology and outgoing-negative/incoming-positive cash
  flow convention do not get mixed into this security/coupon family.
* `Basis` is the §4.11.7 subtype with the five procedures 0 through 4. The
  contract's day-count table and its treatment of counter-factual February 30
  intermediate dates match the source.
* `DateParam` is Number or Text, and scalar conversion of a multi-cell
  Reference follows implied intersection under §6.3.2. The contract's
  ReferenceList refusal and zero-read preflight are an explicit repository
  profile, not a claim that the source calls ReferenceList a DateParam.
* §6.3.6 leaves generic non-integral Integer conversion implementation-defined;
  the contract explicitly chooses checked truncation toward zero. Its
  rejection of fractional or unsupported coupon frequencies is documented as
  a profile choice rather than presented as a universal host rule.
* The contract does not invent settlement ordering for `COUPNUM`, `DISC`,
  `PRICEDISC`, `YIELDDISC`, `YIELDMAT`, `ODDFPRICE`, or `ODDLYIELD` where the
  individual source sections do not print one.

Two profile details still need an explicit acceptance record before support:

1. §6.3.15 says Logical-to-DateParam conversion is implementation-defined.
   The contract's DateParam bullet specifies Number and Text behavior but does
   not say whether a raw Logical DateParam is admitted as 0/1 or refused.
   Select one outcome and bind it to focused tests; do not infer it from the
   Number conversion bullet.
2. `DURATION`, `MDURATION`, and `INTRATE` print `Date`, while neighboring
   functions print `DateParam`. The contract records this spelling and says it
   will bind the functions to the shared DateParam/date-time profile. That
   binding should be made explicit in the implementation contract and tests,
   including raw Number, Text, Logical, and reference behavior, before these
   three functions are advertised.

The selected text-to-number/date profile and the error mapping are repository
choices. The ODF source leaves text-number conversion and several DateParam
paths implementation-defined; those choices are acceptable here because the
contract labels them as profile behavior, but they must remain distinct from
normative source claims.

## Equations and intentionally open profiles

The archive gives displayed equations for exactly the functions called out in
the contract: `ACCRINT`, `AMORLINC`, `DISC`, `INTRATE`, `MDURATION`, `PRICE`,
`RECEIVED`, `TBILLEQ`, and `YIELDDISC`. I found no displayed equation in the
archive sections for `ACCRINTM`, the seven coupon schedule functions,
`DURATION`, the four irregular bond functions, `PRICEDISC`, `PRICEMAT`,
`TBILLPRICE`, `TBILLYIELD`, `YIELD`, or `YIELDMAT`. The contract correctly
keeps those equations, zero-denominator behavior, coupon roll conventions,
inverse/root choice, and nonconvergence outcomes as open profile work. No
Excel, LibreOffice, or financial-library formula is treated as normative by
this review.

## Source anomalies

The contract correctly records the material syntax and anchor anomalies:

* §6.12.3 is headed `ACCRINTM` but its syntax line says `ACCRINT`.
* §6.12.7 is headed `COUPDAYSNC` but its syntax line says `COUPDAYNC`.
* The HTML places an extra `COUPNCD` anchor on the §6.12.7 heading, while the
  actual §6.12.8 heading is `COUPNCD`.
* Coupon “See also” links repeatedly identify `COUPNCD` as §6.12.7; those are
  navigation defects and do not change the numbered heading or body.
* `DURATION`/`MDURATION` use `Date`, `INTRATE` repeats `Basis`, and `PRICE`
  uses `Bas`; the contract preserves these spellings while defining aliases.
* The §6.12.40 `PRICEMAT` body links to itself in “See also” (`PRICEMAT
  6.12.40`). This additional navigation defect is harmless but should be
  added to the contract's anomaly inventory if that inventory is intended to
  be exhaustive.

The “should be greater than 0” wording for the bond/security sections and the
missing date-order constraints are faithfully preserved as open profile
decisions. They should not be silently converted into hard source gates.

## Disposition

The contract is **source-aligned for all 27 functions**: inventory, section
numbers, signatures, result types, printed constraints, displayed equations,
and the known archive naming defects all match the local ODF member. The two
profile clarifications above and the unlisted `PRICEMAT` self-link are the
only review follow-ups. Implementation remains on hold until the contract's
open equations, schedule boundaries, Date/DateParam binding, formula-error
mapping, and bounded solver profiles receive separate acceptance and tests.

No production, contract, test, Cargo, archive, or native evidence file was
changed by this review.
