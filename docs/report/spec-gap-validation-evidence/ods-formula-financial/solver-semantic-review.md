# Financial solver semantic review

Status: **HOLD for implementation fixes**. This is a read-only source review of
the resolver-free root solver and the MIRR reducer. It does not invalidate the
separate contract-semantic PASS in `contract-semantic-review.md`, and it makes
no Cargo, production, native-parity, resource, or performance claim.

## Reviewed inputs

| Input | SHA-256 |
| --- | --- |
| `evaluation/financial/solver.rs` | `6624fce0af61ed0b6d7dddcad0e7bf0623fe17821e53e236a6d741bc3eb8d4b3` |
| `evaluation/financial/reducers.rs` | `de2afc7d993f78e9a6c39ed2f223f94c4cfca0ccd05e14cda63d469fa27ffd6d` |
| `ods-formula-financial/contract.md` | `fbe2283886326582b0e34fd277b38e6fb3e1bdeb2ecfbec5522da0f6f1467e4b` |
| `ods-formula-financial/numerical-design.md` | `fe91e17caae719784653fce7f74e0bcbd892e2c5749b06abc5bfc4e7d71f8bbe` |

The source snapshots are from the isolated financial development checkout.
The source files were not changed during this review, and no Cargo command was
run.

## Findings requiring correction

1. **Cross-bucket compensation is discarded.** `LogSum::finish` first reduces
   each bin with `sum + correction` (lines 154 and 170), and `scale` does the
   same for magnitude bins at line 214. The fixed signed accumulator therefore
   rounds away a retained Neumaier correction before rescaling and cancellation
   against another bucket. This conflicts with the numerical-design requirement
   to retain a term that may become visible after later cancellation. A small
   direct-bin probe is:

   ```text
   high bucket: add_log_term(false, ln(nextafter(1,+inf)))
               add_log_term(true,  ln(2))
               add_log_term(true,  ln(5))
   ```

   The high bin reaches approximately
   `sum = 5.999999999999999` and
   `correction = -2.220446049250313e-16`; `sum + correction` rounds back to
   the stored sum. Add several `add_log_term(false, -epsilon)` terms in the
   immediately lower bucket. Their rescaled contributions are each just under
   `-1`, so the lower-bucket cancellation makes the discarded correction
   observable. The existing same-bucket cancellation test does not exercise
   this path. Cross-bin finalization must preserve both components (or return
   `#NUM!` when that cannot be done), with a regression straddling a bucket
   boundary.

2. **Equal-distance bracket tie-breaking is reversed.** In
   `endpoint_preference`/`choose_closer_bracket` (lines 493–526), a positive
   branch candidate `[start, right]` satisfies `left < right` and is selected,
   even though it has the higher resulting rate. A negative branch candidate
   on the right is rejected, even though larger negative-branch coordinates
   produce lower rates. Equal-distance brackets must select the lower
   resulting rate: the left coordinate bracket on the positive branch and the
   right coordinate bracket on the negative branch.

3. **Bracket endpoint probes are not in distance order.** `find_bracket`
   always evaluates the left endpoint before the right endpoint (lines
   792–809). The contract requires increasing distance from the supplied
   guess. When a bound clips one side, this can return or fail on a farther
   endpoint before the nearer endpoint is inspected.

4. **`IRR` accepts an exact `guess == -1`.** The branch selection and fallback
   at lines 1019–1028 route this value to the positive lower bound. The fixed
   profile requires exact `IRR` guess `-1` to return `#NUM!`; only the explicit
   `RATE` boundary has special handling.

5. **`RATE` does not search both branches after a non-root `-1` boundary.**
   Lines 1066–1105 evaluate the guarded boundary, then always start the
   positive branch. The contract requires adjacent positive and negative
   integral branches in deterministic distance order, with the negative branch
   winning an equal-distance tie. A nonintegral `Nper` at `-1` also needs the
   normal numeric-error path rather than silently entering the positive search.

6. **Newton underflow/stall does not force bisection.** `signed_ratio` can
   exponentiate a very negative log quotient to zero (lines 409–425). The
   solve loop then accepts an unchanged `current` or `previous` candidate
   (lines 883–904) and evaluates it again unless the bracket has already
   converged. An unchanged or non-finite Newton step must be discarded and
   replaced by the interior bisection point.

## MIRR reducer defect

`MirrReducer::finish_rich` rejects every `ratio <= 0` at
`reducers.rs:407`, while `mirr_ratio` rejects a zero growth factor as outside
the positive-real profile at lines 542–545. The accepted MIRR equation has no
positive-rate restriction. Apply the selected real `POWER` profile instead:

- a zero ratio with a positive final exponent produces zero (and hence MIRR
  `-1`) when the denominator is nonzero;
- a negative ratio is valid when the final exponent is an admitted integer,
  and otherwise maps to `#NUM!` under the no-complex/odd-root-extension
  profile; and
- signed integer powers must preserve the signs of `(1 + rate)^n` and of the
  final root.

For `Values = [110, -100]`, `Investment = 0`, and `ReinvestRate = -2`, the
positive-mask NPV is `-110`, the negative-mask NPV is `-100`, and `n = 2`.
The equation gives a ratio of `-1.1`, final exponent `1`, and result `-2.1`.
The current blanket `ratio <= 0` refusal incorrectly returns `#NUM!`.

These findings keep the implementation review at **HOLD** until the solver and
reducer owners land fixes and add focused regressions. No validation result is
claimed for the reviewed snapshot.
