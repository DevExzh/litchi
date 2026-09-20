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

## Rich reducer integration follow-up

After migrating production adapters to rich reducer failures, the focused
evaluation/limits/oracle build stopped on four dead-code diagnostics: the
old `formula_value` conversion and the NPV/FVSCHEDULE/MIRR lossy `push` and
`finish` wrappers. No runtime integration result was produced by that run.
The kernel owner is removing the obsolete production seams.

The separate command `cargo test --locked --offline -p litchi-ods --lib
financial -- --quiet` passed **32 tests**, with 467 unrelated tests filtered
out. It ran in `.codex-tmp/ods-financial-development`; the observed reducer
and solver hashes were `7c0f57f24322dc244dec09763a1f4c057adb63e6e76bc9260ba25e3331127ebe`
and `3b39b8a270f3b5eb4aef4fd1df9d0f67b981e7078d681b8d340f1ca65dccc98d`.
These unit results do not clear the integration build diagnostics.

The earlier five-read projection expectation above has been superseded by
contract review: differing scalar rates require two scalar reads plus two
complete three-cell scans, or **eight reads**. Reusing cells across distinct
rate keys would require range materialization contrary to the streaming
contract. The updated limits fixtures assert exact read order and eight reads,
and add an invariant literal-rate case requiring only three reads to verify
whole-result cache reuse. Sixteen reads remain a defect, not an accepted
threshold. Public execution of those revised fixtures is still pending.

Source review also found that the intermediate XNPV adapter retained both
Values and Dates. The contract requires retaining Values slots and streaming
Dates against them; the integration owner is correcting that memory path.
