# Initial financial scalar-kernel review

Disposition: **changes required**. This reviews the unintegrated kernel in the
isolated financial development checkout, not production support. The reviewed
`evaluation/financial/kernel.rs` has SHA-256
`c6e515b291b689ff46499ad3f638c4afc09c147fd0d1ad34c682e0e84b1228ac`.

The root reviewer compiled the unchanged kernel with `rustc --edition 2024
--test` in a temporary harness. The harness supplied only the two formula-error
variants used by this pure module. This checks kernel arithmetic and its unit
tests; it is not a Cargo, evaluator integration, coercion, or resource gate.
The four existing tests passed and all four independent probes below failed.

| Probe | Expected | Observed |
| --- | --- | --- |
| `FV(-2;2;0;100;0)` | `-100`, from the finite integer-power balance equation | `#NUM!` |
| `RRI(1;100;-100)` | `-2`, from the source ratio-to-power equation | `#NUM!` |
| `NPER(1e-20;-10;100;0;0)` | approximately `10` | `0` |
| `IPMT(0.1;2;2;100;0;1)` | `-100/21`, under the selected beginning-payment amortization profile | `-110/21` |

The common growth helpers unconditionally use `ln_1p`, excluding real finite
integer powers with a negative base. RRI likewise rejects every negative
ratio, even when its exponent is an integer. NPER rounds the growth ratio to
one before taking its logarithm, losing the small-rate limit. IPMT computes
the ordinary prior balance without the required beginning-payment adjustment
for periods after the first.

These are implementation defects in the draft kernel. They require corrected
arithmetic and persistent regression tests before integration. The coder also
needs to align the misspelled `ispmnt` identifier with `ispmt` and the selected
zero-divisor error mapping with the scalar adapter. No universal accuracy or
numerical-domain claim follows from these eight probes. The temporary harness
and executable were removed after recording the results.
