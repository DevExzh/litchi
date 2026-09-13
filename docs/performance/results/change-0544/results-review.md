# 0544 numerical results review

This is a bounded, read-only review of the sealed 0544 numerical evidence.
The disposition is **reject**: `decision.json` records `final_stage: baseline`,
`runtime_restored: true`, `native_passed: false`, `allocation_passed: true`,
`guard_passed: true`, and `cap_boundary_passed: false`. No runtime speedup is
retained.

## Gate disposition

The native primary gate fails only the medium repeat 2 workflow row: total
p50 improves 2.8573315% and total mean improves 2.9169942%, both below the
required 3%. Its planning p50 improves 20.4443832%; the passing planning
diagnostic does not override the every-shape/repeat workflow gate.

The supplemental valid cap gate also rejects the candidate. Cap 164 repeat 1
exceeds the 1.05 limit at p50/mean by 5.4490429%/5.3580591%; cap 256 repeat 2
exceeds it by 6.9398511%/6.8872736%. The other cap rows pass. Allocation has
no allocation-analysis adverse flags and passes its planning gates; the
refusal guard passes its stated envelopes. Conditional profile, hardware, and
eager lanes were correctly not used to turn a failed pilot into an admission.

## Adverse-row coverage

`adverse-review.json` is complete, has disposition `reject`, and covers all
450 flags with unique exact IDs `0544-review-0001` through
`0544-review-0450`; the markdown companion retains an individual entry for
each row, while the JSON retains each row's original metric vector. The
source-group counts are:

| Evidence source | Rows |
| --- | ---: |
| Main native adverse flags | 68 |
| Main same-build drift | 102 |
| Allocation-analysis adverse flags | 0 |
| Refusal normal-lane adverse flags | 32 |
| Refusal allocation-lane adverse flags | 221 |
| Refusal same-build drift | 23 |
| Cap adverse rows | 4 |
| Cap same-build drift | 0 |
| **Total** | **450** |

The rows are evidence to preserve and interpret, not permission to discard
the failed candidate. Same-build drift is retained as a stability diagnostic;
it is not dismissed as noise or presented as a candidate effect.

## Claim boundaries

The scanner/preflight statements are scoped to the frozen 0544 candidate versus
its matched 0544 baseline and the separately identified valid cap fixtures.
They do not claim a cross-campaign or same-build speedup, and no 0543 timing is
reused. Same-build rows remain repeat diagnostics rather than candidate
causality. The next-priority correction is design-only: the proposed
one-scan bound `1 + initial_text + 2 * count(< or &)` gives the current 128
numeric grid a coarse bound of 131,593, above the 131,072 cap, so that path
would decline the coarse 128 grid. This is an explicit coverage cost in the
next design, not a 0544 measurement or admission claim.

Allocation-instrumented timing, allocation calls/reallocations, and child RSS
remain diagnostics under their declared scopes; none is converted into a
native latency claim. The baseline-valid invalid-input envelope passing does
not establish unchanged same-invalid cost: the individually reviewed
late-validator rows retain their timing and allocation increases, including
the reported same-invalid cost. The review therefore keeps valid-envelope
correctness separate from invalid-input performance.

## Evidence bindings

The reviewed files are bound by these SHA-256 values:

| File | SHA-256 |
| --- | --- |
| `decision.json` | `2ab4891f056623904cae6882d082a1910bf4ce5ca2b283a400b096dcf629eb4e` |
| `comparison.json` | `e87456ab9169beef2318497e5d2032bc95cc4fc645b3c6a003d41af6a5ef531a` |
| `allocation-analysis.json` | `2f9ac59d84ca0db67822f459389cf9af305d599cf48fb4fa916a13ce930335c9` |
| `guard-analysis.json` | `c958c735e949ed99d98f04972f7979d76d206f592c0c6ecdec4bbad34bb3a60b` |
| `cap-boundary/cap-analysis.json` | `5504ba11cccc1a818c66714268d9f05465e216d7ff5770fad777e818ef8cda88` |
| `adverse-review.json` | `96ea6c967130f51c87707d505a18bf373d3ee3fbfd6be06a00d01109fe06468a` |
| `next-priority.md` | `b70897f57230b250039b7a90949773ec90332482be56fac8998ffaf1ab23ebc4` |
