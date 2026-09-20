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
