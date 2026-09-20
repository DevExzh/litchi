# Financial integration development findings

These are development observations from the isolated checkout, not frozen
acceptance gates. Sources are under active correction.

The production library compiled during the focused test-target build.
`cargo test --locked --offline -p litchi-ods --test
ods_formula_financial_oracle -- --quiet` then ran both oracle tests: the
65-vector scalar replay passed; the reducer replay failed.

The reducer replay exposed these value-adapter mismatches:

- Beginning-payment CUMIPMT: expected approximately `-4.7619047619`,
  observed approximately `-15.2380952381`.
- CUMIPMT's Integer Type conversion: expected zero for the selected
  beginning-payment case, observed `-10`.
- CUMPRINC with an included period equal to Nper: expected `#NUM!`,
  observed a numeric principal payment.
- CUMPRINC reversed empty interval: expected zero, observed `#NUM!`.
- MIRR's mixed Text/Empty/Logical reference fixture in matrix mode:
  expected approximately `0.05403984744`, observed `#VALUE!`.

The same replay also had harness defects: it attempted the scalar-only API
on unsupported array syntax and read the compact repeated-zero descriptor
incorrectly. Those must be fixed before the remaining reducer cases yield
meaningful implementation evidence. Unsupported-array failures are not
counted as financial semantic failures here.

The limits target executed 18 tests: **16 passed, 2 failed**. Both failing
projection fixtures supplied a horizontal `{TRUE();TRUE()}` condition with
vertical scalar inputs. The intended vertical condition is
`{TRUE()|TRUE()}`, as in the existing conditional regressions. The owner is
correcting geometry while retaining the two-row output and five-read
expectations. Cache correctness is pending that rerun.

The evaluation target still has a test-only dead-code diagnostic for an
unused Logical fixture variant. It has not produced runtime results yet.
All findings were handed to their implementation or test owners. No full
financial, resource, native-parity, or performance PASS is claimed.
