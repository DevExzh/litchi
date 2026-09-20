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

## Corrected-handoff re-review

The implementation owners supplied a corrected handoff. The reviewed source
hashes are now:

| Input | SHA-256 |
| --- | --- |
| `evaluation/financial/solver.rs` | `3b39b8a270f3b5eb4aef4fd1df9d0f67b981e7078d681b8d340f1ca65dccc98d` |
| `evaluation/financial/reducers.rs` | `7c0f57f24322dc244dec09763a1f4c057adb63e6e76bc9260ba25e3331127ebe` |
| `evaluation/financial/kernel.rs` | `142ab9fb8d4b6200ebbe31ba1177354cfba2dc116835e492645aa035014a85d2` |
| `ods-formula-financial/contract.md` | `fbe2283886326582b0e34fd277b38e6fb3e1bdeb2ecfbec5522da0f6f1467e4b` |
| `ods-formula-financial/numerical-design.md` | `fe91e17caae719784653fce7f74e0bcbd892e2c5749b06abc5bfc4e7d71f8bbe` |

The previous six solver findings are resolved in this handoff:

- `LogSum::finish` retains each bin's `sum` and `correction` separately while
  rescaling lower bins (lines 152–196). The focused
  `scaled_sum_preserves_compensation_across_buckets` regression constructs a
  lower-bucket cancellation and checks that the retained residue remains
  visible.
- Equal-distance bracket selection compares the lower-rate endpoint for the
  active branch (lines 498–505), and `probe_left_first` orders endpoints by
  distance with the lower-rate equal-distance tie (lines 508–519).
- `find_bracket` now uses that ordering for endpoint evaluation (lines
  812–832), rather than unconditionally evaluating the left side first.
- `IRR` rejects an exact `guess == -1` before branch selection (lines
  1047–1049).
- `RATE` rejects a nonintegral `Nper` at the exact boundary, evaluates the
  guarded boundary, and searches both adjacent integral branches when the
  boundary is not a root (lines 1104–1139).
- `signed_ratio` rejects an underflowed zero step, and the solve loop falls
  back to the bracket midpoint after an unchanged Newton candidate (lines
  414–425 and 905–932).

The MIRR reducer now follows the selected signed-power profile. It returns
`-1` for a zero ratio with a positive exponent and admits a negative ratio only
when `1 / (n - 1)` is an exact integer (reducers lines 483–510). The focused
`mirr_allows_negative_ratio_for_exact_integral_root` and
`mirr_zero_ratio_power_returns_minus_one` regressions cover the prior defect,
including the `-2.1` example from the original finding.

This re-review is **PASS for the six prior solver findings and the MIRR source
defect**, pending ordinary integration/gate evidence. The exact-boundary RATE
implementation chooses the positive side first because its nearest
representable rate is `2^-53` above `-1`, versus `2^-52` below it; this is the
source's explicit rate-space interpretation of the contract's deterministic
distance order, with the negative branch reserved for the equal-distance tie.
If the contract is later interpreted as transformed-coordinate distance across
the disconnected branches, that profile choice must be made explicit before
acceptance. The absolute-magnitude `scale()` still combines its positive
compensation pair before `log_add_exp`; because that state has no signed
cross-bucket cancellation, I record it as a low-order accuracy note rather
than a blocker to the signed residual finding.

No Cargo command, production edit, or final performance/gate claim was made in
this re-review.

## Reducer wrapper-only follow-up

The reducer handoff then removed the obsolete lossy scalar wrappers. The
current hashes are:

| Input | SHA-256 |
| --- | --- |
| `evaluation/financial/solver.rs` | `3b39b8a270f3b5eb4aef4fd1df9d0f67b981e7078d681b8d340f1ca65dccc98d` |
| `evaluation/financial/reducers.rs` | `ecda9ba493af93ac7fb8f7da24048b3dd6888a52e18ae8b2d45ee3bc8a6c9e27` |
| `ods-formula-financial/contract.md` | `fbe2283886326582b0e34fd277b38e6fb3e1bdeb2ecfbec5522da0f6f1467e4b` |
| `ods-formula-financial/numerical-design.md` | `fe91e17caae719784653fce7f74e0bcbd892e2c5749b06abc5bfc4e7d71f8bbe` |

The delta is wrapper/API hygiene only: the scalar `npv`, `fvschedule`, and
`mirr` convenience functions and `test_scalar_result` are test-only, while
production call sites use the rich `push_rich`/`finish_rich` API. The reducer
state, signed-power MIRR path, formula-error behavior, and typed accumulator
boundary are unchanged. This follow-up remains **PASS** for the reviewed
semantic findings; no new source blocker was found.

## RATE zero-guard follow-up

The solver handoff adds one scoped correction for the guarded `RATE` boundary.
The current solver hash is
`387f9791103bc97fb85fa99e46502e0b5e74f8da59586ba755a8785fd18761c8`; the
reducers remain `ecda9ba493af93ac7fb8f7da24048b3dd6888a52e18ae8b2d45ee3bc8a6c9e27`.

After the existing boundary work charge and cancellation check, the code now
recognizes `boundary == 0.0` before calling `residual_from_sums`. This covers
the valid annuity-due case where both guarded terms are zero and the temporary
log accumulator has no scale. The new
`rate_due_boundary_accepts_empty_guarded_residual` regression asserts an exact
`-1.0` result; the public roots replay also covers this path. The change is
localized to guarded-boundary publication and preserves the work/cancellation
fences. No new semantic blocker was found, so the implementation disposition
remains **PASS**, pending ordinary integration/gate evidence.
