# Dispersion oracle comparison review

This review records why the retained comparison policies in
`numeric-goldens.json` are bounded and observable. The independent reference
uses exact represented-binary64 fractions for variance and a 600-digit Decimal
square root for standard deviation. It makes no correct-rounding claim about
the production scaled compensated first-offset kernel. The tolerances below
are acceptance limits for this retained corpus, not universal error bounds
for every admitted sequence or input length.

The ordinary policy is eight ULP. A four-ULP trial was insufficient: the
`VARA` `typed-mixed` case produced `7.428571428571422` (`401db6db6db6db66`)
against the exact reference `7.428571428571429`
(`401db6db6db6db6e`), an eight-ULP difference. The same trial produced six
ULPs for `VARPA` `typed-mixed`: `6.499999999999995`
(`4019fffffffffffa`) versus `6.5` (`401a000000000000`). The eight-ULP bound
therefore records the largest observed ordinary error instead of silently
loosening a relative epsilon. It is also the existing fixed-width numeric
aggregate allowance used for scaled product observations.

Adversarial cancellation rows use the existing database numeric profile's
explicit `2e-12` relative bound. The largest observed adversarial error was
`VAR` on `majority-unit-forward-100k`: actual
`9.99990000098478e-6` (`3ee4f8a7ca7c5ac4`) versus exact
`9.99990000099999e-6` (`3ee4f8a7ca7c7dd7`), with relative error
`1.52111697769693914e-12`. The reverse-order and 10,000-element large-value
fixtures remain separately retained, so order-dependent `n*epsilon`
cancellation is visible rather than hidden by a wider bound. Expected zero
results are checked as exact positive zero.

The serial diagnostic covered all 512 observations in both Scalar and Matrix
modes. Eight focused oracle tests passed, with no numeric mismatch beyond the
documented policies. The diagnostic instrumentation was removed after the
measurement; the owned test remains rustfmt-clean. The retained checks are:

```text
python3 numeric_oracle.py --check
cargo test --locked --offline -p litchi-ods --test ods_formula_dispersion_oracle
```

No production source, contract, or golden input was changed for this review,
and the temporary metric log, diff, and Python cache were removed.
