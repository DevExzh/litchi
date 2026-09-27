# 0782 results review

Disposition: **reject the `Option<&str>` candidate and retain the restored
baseline**. This is an independent, bounded read of the frozen 0782 packet at
base `f1df64d15a8f2711e554f58acb413b06d56ef1ca`. I did not run Cargo, native
captures, profilers, or any hardware work.

The frozen policy in `adoption-policy.json` defines a persistent control
regression as a paired native process-p50 after/before median ratio above
`1.05` with a bootstrap 95% CI lower bound above `1.0`, in any of all ten
cases. The native plan has six alternating blocks, 30 samples plus three
warmups, CPU 12; `analysis.json` records the bootstrap as 10,000 median
resamples with seed `782078`.

## Guard arithmetic

The packet's ratios are consistent with `1 + change_percent / 100`, and the
independent decision receipt in `decision-audit.json` identifies exactly four
guard failures:

| case | p50 change | after/before ratio | bootstrap 95% CI for ratio | guard |
| --- | ---: | ---: | ---: | --- |
| tiny/write | -3.562% | 0.964383 | [0.957443, 0.987059] | pass |
| tiny/lifecycle | -3.677% | 0.963233 | [0.945745, 0.973443] | pass |
| many/write | **+12.897%** | **1.128968** | **[1.110164, 1.135918]** | **fail** |
| many/lifecycle | **+11.089%** | **1.110891** | **[1.108498, 1.123324]** | **fail** |
| payload/write | -24.328% | 0.756719 | [0.735985, 0.795278] | pass |
| payload/lifecycle | -21.642% | 0.783584 | [0.771905, 0.818333] | pass |
| unicode/write | **+20.650%** | **1.206502** | **[1.202215, 1.209476]** | **fail** |
| unicode/lifecycle | **+19.069%** | **1.190694** | **[1.189756, 1.192544]** | **fail** |
| rich/write | +0.121% | 1.001213 | [0.990073, 1.019772] | pass |
| rich/lifecycle | -0.865% | 0.991346 | [0.984315, 0.999672] | pass |

The four failing p50 changes are positive in every paired block. Their block
ranges are +9.678% to +13.612% (`many/write`), +10.720% to +13.118%
(`many/lifecycle`), +20.173% to +21.179% (`unicode/write`), and +18.955% to
+19.327% (`unicode/lifecycle`). This satisfies the frozen guard directly;
the payload gains cannot be used to hide those common-workflow regressions.

## Resource and observer evidence

The allocation packet shows no `net_live` or `peak_above_entry` change in any
of the ten cases, and both allocation regression and spread flag lists are
empty (`analysis.json` under `allocation.analysis`). It also reports useful
allocation reductions in the target payload cases: allocated bytes fall
19.447% for write and 16.232% for lifecycle. Heaptrack independently reports
the plain conversion path falling from 192 events and 7,680,000 requested
bytes to zero (`observer-analysis.json` under `heaptrack.diagnostic`).

Those resource results do not override the latency guard. Heaptrack is a
whole-process allocation-site diagnostic. The observer `perf` lane is also
whole-process setup-plus-verification evidence; its cycles rise about 10.641%
to 14.246% while instructions change about -0.046%, so it cannot establish an
operation latency cause. I make no causal attribution from either observer.

## Correctness and final disposition

The packet reports exact source/output and semantic/raw-text checks, 1,287
tests passed, 11 ignored, zero failed, and all six quality gates exited zero
(`test-summary.json`, `quality.json`, and `analysis.json.verification`). The
correctness and measurement receipts are therefore sufficient to make the
performance decision; they do not turn the candidate into an adoptable
change.

`disposition.json` records `status: rejected`,
`production_change_retained: false`, and a restored-source receipt. Its reason
matches the four guard failures above and leaves the latency cause unproven.
`baseline-fixture-parity.json` explicitly scopes parity with the corrected 0781
qualification to source/output identity, not timing; I did not make a
cross-run 0781 timing comparison.
