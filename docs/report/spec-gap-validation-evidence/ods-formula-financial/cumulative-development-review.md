# Cumulative stability development findings

The first positive-rate, zero-future balance rewrite passes the new 12-row,
256-digit cumulative oracle through scalar and value public APIs. The broader
15-test evaluation target and scalar/reducer/root oracle targets also passed
against the first rewrite. These are development results, not final acceptance.

Root then compiled a temporary standalone probe against kernel SHA256
`4514131064b761b9c1fd02941b42ce415164004fe486ddc80adca63e387e8485`.
It reproduced two remaining finite-result failures:

| Formula | Independent 256-digit Decimal result rounded to f64 | Observed kernel result |
| --- | --- | --- |
| `PPMT(2;1;2;1e308;0;1)` | `-7.5e307` | `#NUM!` |
| `IPMT(0.1;512;512;100;1;0)` | `-0.8181818181818183` | `-1048576.0` |

The PPMT helper multiplies principal by rate before division, overflowing an
intermediate even though the due-payment result is finite. The IPMT helper
still uses the cancellation-prone balance formula for nonzero future value.
The owner is generalizing the positive-rate stable calculation; negative-base,
zero-rate, and domain semantics must remain intact. Public regressions now
cover both formulas with relative `1e-12` tolerance and no absolute floor.

Source review additionally found the cumulative Kahan accumulator used a
subtract-on-next-step compensation convention but added that compensation at
finalization. The owner must fix the sign convention and retain a focused
regression. Temporary probe sources and binary were removed automatically.
No performance acceptance or frozen-source claim is made here.

## Generalized positive-rate checkpoint

Isolated commit `ab93dec951` scales the payment equation by inverse growth,
computes remaining balance from the terminal value, and subtracts the Kahan
rounding excess at finalization. The previous two finite-result regressions
and all twelve long-span oracle rows now pass through public APIs. The
focused suites pass 16 evaluation, 23 resource, two scalar/reducer oracle, and
two roots tests. The 35 financial library tests and all-target Clippy with
warnings denied also pass. Debug profiling preflight passes all 91 cases; its
receipt is `performance/development-preflight-stability-fixed.json`.

This is a tested development checkpoint, not final numerical acceptance.
Independent edge review, final frozen gates, and timing capture remain open.
