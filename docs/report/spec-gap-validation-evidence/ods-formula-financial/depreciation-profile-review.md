# Depreciation profile review

Status: source review complete; the implementation gate remains **PENDING
CONTRACT RESOLUTION** for the profile choices below. This review owns no
production or test changes. It records executable repository-profile
recommendations, the source conflicts they repair, and the negative evidence
required before these seven functions are advertised.

## Authority and method

The review used the repository-local ODF 1.4 archive, not a spreadsheet host or
the legacy financial implementation:

| Artifact | SHA-256 |
| --- | --- |
| [`OpenDocument-v1.4-os.zip`](../../../../3rdparty/specs/OpenDocument-v1.4-os.zip) | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| `part4-formula/OpenDocument-v1.4-os-part4-formula.html` | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |

I read §§6.12.13 (`DB`), 6.12.14 (`DDB`), 6.12.16 (`DOLLARDE`), 6.12.17
(`DOLLARFR`), 6.12.45 (`SLN`), 6.12.46 (`SYD`), and 6.12.50 (`VDB`), with
§6.17.8 (`TRUNC`) for the two fractional monetary functions. The boundary
review follows [ADR 0001](../../../adr/0001-priorities-and-api-layers.md),
[ADR 0004](../../../adr/0004-semantic-api-design.md),
[ADR 0005](../../../adr/0005-io-memory-and-performance.md),
[ADR 0006](../../../adr/0006-validation-security-and-compatibility.md), and
[ADR 0008](../../../adr/0008-migration-and-verification.md).

The companion [depreciation contract](depreciation-contract.md) is a useful
transcription of the source, but its open-profile section is still binding.
The source's `SLN` typo and `VDB` prose/table conflict cannot be repaired by
copying behavior from Excel, LibreOffice, or the existing host code.

## Disposition

| Function | Source finding | Profile disposition |
| --- | --- | --- |
| `DB` | The rate is rounded to three decimals; the source gives no tie rule or fractional-`Period` rule. Omitted `Month` is distinguished from supplied `Month`. | Recommend deterministic ties-to-even rate rounding, a continuous residual extension for finite fractional `Period`, and preserved optional-argument presence. These are contract-profile choices, not claims about ODF wording. |
| `DDB` | The source supplies an executable fractional-`Period` algorithm and the `rate >= 1` branch, despite its integral-period description. | Recommend retaining the printed Number signature and source algorithm. Validate the integer-period relation to `VDB`; do not replace the fractional path with host behavior. |
| `DOLLARDE`/`DOLLARFR` | The equations and positive Integer denominator are complete; negative operands depend on `TRUNC`. | Recommend the existing truncation-toward-zero Integer bridge and §6.17.8 `TRUNC`; denominator/domain errors use the shared financial profile. |
| `SLN` | The heading and summary describe straight-line depreciation, but the syntax says `DDB(...)`, constraints say none, and no equation is printed. | Recommend an explicit repository repair: recognize `SLN`, use the ordinary straight-line equation, and disclose that the archive syntax/equation is defective. This is a profile decision, not a claim that the printed syntax is correct. |
| `SYD` | The equation is printed, while the section says `Constraints: None`; prose calls `LifeTime` positive integer. | Recommend preserving `Constraints: None` in admission until the contract owner resolves the prose tension. Use the printed equation and finite-result checks. |
| `VDB` | Fractional periods and the `initialPeriod` rule are explicit, but there is no standalone interval equation. The table allows equality while prose requires `EndPeriod > StartPeriod`. | Recommend retaining the explicit Number/equality domain: equality is an empty interval with result zero, while reversed bounds are `#NUM!`. Use the executable piecewise DDB/straight-line profile below and retain fractional values. The table/prose conflict still needs a contract note. |

## Resolved profile proposals

### `DB`

The source computes

```text
rate = 1 - (Salvage / Cost)^(1 / LifeTime)
```

and says that `rate` is rounded to three decimal places before it is used in
the residual recurrence. The recommended repository profile is decimal
round-to-nearest, ties-to-even at exactly three decimal places. This matches
the deterministic rounding convention already used by the independent
financial oracles and avoids depending on a host language's `round` rule. It
applies once to `rate`; it does not round each residual or allowance.

`Period` is a `Number` and the only printed constraint is `Period > 0`. A
profile that rejects every fractional Number would narrow an explicit source
domain, so the recommended profile keeps finite fractional periods and makes
their residual calculation executable. Define a residual curve `R(t)` for
`t >= 0`:

```text
R(0) = Cost
R(t) = Cost * (1 - (Month / 12) * rate * t)                 for 0 <= t <= 1
R(t) = R(1) * (1 - rate)^(t - 1)                             for 1 < t <= LifeTime
R(t) = R(LifeTime) * (1 - (1 - Month / 12) * rate
                       * (t - LifeTime))                    for LifeTime < t <= LifeTime + 1
```

The last branch is used only when `Month` is explicitly supplied and the
source's `LifeTime + 1` residual is in range. The allowance at a requested
period `p` is `Cost - R(p)` when `0 < p <= 1`, and `R(p) - R(p - 1)` for
`p > 1`, using the subtraction orientation already transcribed in the
contract. The source zero-tail rule is applied before evaluating the last
allowance. This piecewise-geometric extension agrees with every integral
source residual, preserves the Number domain, and makes fractional periods
deterministic without silently selecting a host's floor or ceiling. It is a
repository proposal: the contract must accept the extension and its allowance
orientation before tests treat fractional `DB` as conformance evidence.

Optional `Month` presence is semantic state. When the slot is omitted, the
source supplies its default 12. When the slot is explicitly supplied,
including an explicit 12, the source's “measured in years” wording and the
`LifeTime + 1` residual branch apply. The evaluator and demand-cache key must
retain this presence bit; evaluating a default and then pretending that the
argument was absent changes the source-defined branch. `Month` remains a
finite Number subject to `0 < Month < 13`; it is not silently truncated to an
integer.

The explicit source domains (`Cost > 0`, `Salvage >= 0`, `LifeTime > 0`, and
`Period > 0`) are numeric/domain failures. The recurrence and the zero-tail
rule must use checked arithmetic. A lifetime-sized loop or vector is not an
acceptable implementation, even though the mathematical recurrence is stated
period by period.

### `DDB`, `DOLLARDE`, `DOLLARFR`, and `SYD`

`DDB` is the one depreciation entry with a complete non-integral algorithm in
the archive. The `rate >= 1` branch must be retained exactly, including the
`Period = 1` special case. The integer-period identity with `VDB(...; TRUE)`
is a post-implementation invariant, not a license to borrow VDB's unresolved
fractional profile before that profile is recorded. `DDB`'s explicit constraints
remain the source table;
non-finite intermediate values are numeric formula errors.

The two `DOLLAR*` equations use `TRUNC`, so a negative fractional operand is
truncated toward zero. A denominator that becomes non-positive after the
existing Integer conversion is a numeric/domain error. The functions must not
replace `TRUNC` with floor and must not use a host currency formatter.

`SYD` uses the printed equation:

```text
(Cost - Salvage) * (LifeTime + 1 - Period) * 2
    / ((LifeTime + 1) * LifeTime)
```

The archive explicitly says `Constraints: None`. Until an amendment gives the
prose sentence “LifeTime ... positive integer” executable force, the admission
layer must not import the stricter `DB`/`DDB`/`VDB` domains. A zero denominator
is `#DIV/0!`; a non-finite result is `#NUM!`. A finite result for an otherwise
unconstrained Number is not converted to `#NUM!` merely because another
depreciation function has a lifetime constraint.

### `VDB` equality and fractional boundaries

The source has two statements about an empty interval:

* the constraint table permits `StartPeriod = EndPeriod`; and
* the parameter prose says `EndPeriod` is greater than `StartPeriod`.

The recommended profile gives precedence to the explicit constraint table for
admission and treats equality as an empty interval: after scalar arguments are
successfully evaluated, `StartPeriod == EndPeriod` returns numeric zero without
walking an interval. This preserves the printed Number domain and the natural
empty-sum identity. `EndPeriod < StartPeriod` remains `#NUM!`. The prose
conflict must be disclosed in the contract; changing equality to `#NUM!` would
be a deliberate narrowing profile choice rather than a direct transcription.
Formula Errors or typed failures in the scalar arguments still follow normal
source-order precedence before the empty result is published.

Fractional `LifeTime`, `StartPeriod`, and `EndPeriod` remain admitted finite
Numbers within the source bounds. No integer conversion or implicit ceiling is
permitted. For `x >= 0`, record the source's fractional part as
`x - floor(x)`; the first partial boundary uses the nonzero fractional part of
`StartPeriod` or `EndPeriod`, and when both are fractional it uses the
`StartPeriod` fraction. Omitted `DepreciationFactor` is 2; zero is admitted;
omitted `NoSwitch` is false. Both optional-presence bits must survive argument
classification and caching.

The following is an executable repository profile for the missing interval
equation. It is derived from the source's DDB relation and switch sentence; it
is not presented as hidden ODF text:

1. Let `B(0) = Cost`. For each complete unit interval beginning at `k`, let
   `declining_k = max(0, min(B(k) * factor / LifeTime,
   B(k) - Salvage))`. Let
   `straight_k = max(0, (B(k) - Salvage) / (LifeTime - k))` when
   `k < LifeTime`.
2. If `NoSwitch` is false and `straight_k > declining_k`, use
   `d_k = straight_k`; otherwise use `d_k = declining_k`. Advance
   `B(k + 1) = B(k) - d_k`, clamping only at `Salvage` through the same
   checked `min` operation. This is the source's “switch when straight line is
   greater” rule expressed at a period boundary.
3. Define cumulative depreciation `C(0) = 0` and extend each interval
   linearly: `C(k + u) = C(k) + u * d_k` for `0 <= u <= 1`, truncating the
   final interval at `LifeTime`. The result is
   `C(EndPeriod) - C(StartPeriod)`. Thus `[0, .75]` is 75% of the first
   boundary-selected period, and an interval crossing an integer boundary is
   the exact sum of its two fractional overlaps. In this proposal, when both
   endpoints are fractional the first boundary is the `StartPeriod` boundary;
   that is the explicit interpretation of the source's `initialPeriod` rule.
   If the contract instead means the selected fraction as a multiplier, it
   must replace this step with vectors before implementation evidence is
   accepted.

The implementation may replace the unit walk with closed-form geometric sums
and a checked switch-boundary search, but it must produce the same profile and
must not allocate a vector proportional to `LifeTime`. All intermediate
periods, powers, and remaining-life divisors are finite and work-charged. The
contract should accept this algorithm, or replace it with another explicit
one, before VDB tests become conformance evidence. A host-specific VDB routine
is not acceptable evidence.

## Error and precedence profile

The seven functions should use the completed financial-family mapping so that
their errors are reviewable and consistent:

| Condition | Result |
| --- | --- |
| Wrong arity, pseudotype, known shape, or refused conversion | Formula `#VALUE!` |
| Printed domain violation, undefined fractional profile, non-finite power/result, or bounded numeric failure | Formula `#NUM!` |
| Direct zero divisor in an otherwise admitted arithmetic branch | Formula `#DIV/0!` |
| Resolver/provider `Unsupported`, `ResourceLimit`, `Allocation`, `Cancelled`, `SourceChanged`, or source-version failure | Typed evaluator failure |

The source sections do not assign these formula-error categories themselves;
the first three rows are the repository profile. The proposed SLN repair adds
its explicit positive-lifetime check before division, so `LifeTime <= 0` is
`#NUM!`; the direct zero-divisor row applies to an admitted arithmetic branch
such as `SYD` under its `Constraints: None` profile. Fractional positive SLN
lifetimes remain Numbers rather than being silently truncated.

Arguments are observed in source order. A formula Error is retained as a value
and wins over a later generated formula error according to the shared financial
precedence. A typed resource, cancellation, source, allocation, or provider
failure aborts and supersedes an earlier formula Error. These scalar functions
do not admit a complete reference sequence, so no implementation may scan a
range to “find a better” depreciation argument or hide a typed read failure.
Known shape/type refusal must happen before a resolver read; computed children
may run until their resulting type is known. Source-version and cancellation
fences surround the complete evaluation and result publication.

## `SLN` decision record

The archive defect is concrete, not an implementation detail:

1. §6.12.45 is headed `SLN` and describes straight-line depreciation.
2. Its syntax line says `DDB(Number Cost; Number Salvage; Number LifeTime)`.
3. It declares `Constraints: None`.
4. It says only that the lifetime is a positive integer and supplies no equation
   or MathML.

A repository profile can repair this defect explicitly: recognize
`SLN(Number Cost; Number Salvage; Number LifeTime)`, evaluate
`(Cost - Salvage) / LifeTime`, require a finite positive `LifeTime`, and retain
fractional positive lifetimes because the printed pseudotype is `Number` and
the formal constraint line is `None`. `LifeTime <= 0` is a numeric-domain
`#NUM!` under this repaired profile. The prose adjective “positive integer”
remains a disclosed source conflict and should be covered by a profile test,
rather than silently truncating a Number.

This is an explicit repository repair inferred from the heading and ordinary
meaning of “straight-line”; it does not turn the missing source equation into
normative ODF text. Once the contract records that repair, the evaluator may
advertise the profile and must not route `SLN` to `DDB`.

## Evidence required to close the hold

Before implementation support is advertised, the contract owner should amend
the open points and add focused evidence for:

* DB rate ties exactly at the third decimal, fractional residual periods,
  omitted versus explicit `Month = 12`, fractional Month, and the
  `LifeTime + 1`/zero-tail boundary;
* DDB fractional periods, `rate >= 1`, salvage stopping, and the integer
  DDB/VDB identity;
* DOLLARDE/DOLLARFR negative `TRUNC`, denominator truncation, and round trips;
* the accepted SLN signature/equation and zero/negative/fractional lifetime
  errors;
* SYD finite unconstrained values, zero denominator, and explicit handling of
  the prose-versus-`Constraints: None` conflict; and
* VDB equality-as-zero and reversed-bound refusal, fractional initial-period
  choices, factor zero, no-switch versus switch, cross-boundary intervals, and
  checked work and cancellation limits.

The resource evidence must also show fixed-size or checked scratch, no
lifetime-sized allocation, no hidden clock/provider, and typed failure
precedence after an earlier formula Error. Until these proposed repairs and
the executable fractional interval profile are recorded in the contract, the
appropriate disposition is source/profile review complete but implementation
support not established.
