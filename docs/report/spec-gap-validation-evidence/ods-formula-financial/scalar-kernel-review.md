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

## First correction follow-up

Root repeated the isolated `rustc` probe against kernel SHA-256
`1271770ec5522c34bcf15c05fc8da463a4b4ee5c371548942d7fe3a7af3b7cc1`.
All four original probes now pass, together with the four kernel tests.
Two additional RRI probes still return `#NUM!` and need correction:

* `RRI(0.5;100;-100)` has exponent 2 and should return 0. Requiring an odd
  integer NPER excludes this real integer-exponent case.
* `RRI(2;1e-300;1e300)` has a finite result approximately `1e300`, but the
  intermediate ratio overflows. Log-magnitude arithmetic can avoid forming
  that ratio.

The follow-up result is 8 passed and 2 failed; disposition remains changes
required. Its temporary harness and executable were also removed. These are
arithmetic probes, not integrated evaluator or Cargo evidence.

## Second correction follow-up

Kernel SHA-256
`bf48b77f683e36e42d75798d24dffa35bc870d7485035f89328bacfb1024983c`
passes all six independent probes and the four kernel tests: 10 passed,
0 failed. Root checked that the source hash was unchanged across compilation
and execution, then removed the temporary harness and executable. The six
reported defects are resolved in this isolated version. Broader oracle,
coercion, solver, resolver, and integration validation remain outstanding;
this result does not promote the financial batch to supported.
