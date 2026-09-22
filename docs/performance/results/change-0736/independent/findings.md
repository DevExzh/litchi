# Independent 0735 sample-order audit

This is a read-only analysis of the sealed 0735 packet. It starts no Rust,
Cargo, native, allocator, or profiler process. The companion
[`independent-order.py`](independent-order.py) verifies the 0735 artifact
manifest, reads all 36 native process reports, and writes
[`independent-analysis.json`](independent-analysis.json). The sealed packet
digest is `dbe26c50532792cba24871862c96dc0889b3fc21626cb57357fab6a08ff03955`.

The input contains 36 native processes: nine baseline/candidate pairs for
each fixture, with 50 measured samples after three warmups. The 1,800 sample
timings and all per-sample semantic oracles are present and pass. Samples
inside a process and overlapping windows are descriptive observations; the
nine paired processes per fixture remain the comparison units.

The independent full-window calculation reproduces the 0735 paired midpoint
p50 changes: primary −3.529943% and secondary +6.336312%. The corresponding
mean changes are primary −5.290004% and secondary +1.846582%. The window
analysis below computes a midpoint p50 within each process, then the ordinary
median of the nine paired percentage changes.

| Measured sample indices | Primary p50 change | Secondary p50 change |
| --- | ---: | ---: |
| 0–49 | −3.53% | +6.34% |
| 0–9 | +11.97% | +5.80% |
| 10–39 | −6.72% | +3.57% |
| 40–49 | −5.09% | +7.47% |
| 0–24 | +2.96% | +3.25% |
| 25–49 | −6.90% | +5.78% |

The primary comparison changes sign across sample windows. The secondary
comparison remains positive in the first and final ten samples, while the
middle thirty are smaller than the five percent review threshold. This is
sample-position sensitivity, not evidence of a causal allocator or cache
mechanism.

The final-ten versus first-ten within-process p50 change is consistent across
the nine processes in each cell:

| Cell | Median | Range |
| --- | ---: | ---: |
| Primary baseline | +18.35% | +15.89% to +19.53% |
| Primary candidate | +0.25% | −0.41% to +1.19% |
| Secondary baseline | −2.78% | −3.81% to −2.09% |
| Secondary candidate | −1.25% | −2.25% to −0.96% |

The secondary paired percentage also has repeated position bands when grouped
into five-sample bins. The medians for indices 00–04, 05–09, 10–14, 15–19,
20–24, 25–29, 30–34, 35–39, 40–44, and 45–49 are respectively −1.03%,
+9.36%, +2.22%, +1.30%, −0.96%, −7.40%, +5.89%, +4.36%, +8.60%, and −0.45%.
Each bin contains 45 observations formed from the nine process pairs; these
are not 45 independent experiments.

The sample sequence has descriptive lag dependence. The table is the median
within-cell Pearson correlation between timings separated by the indicated
sample lag, across the nine processes:

| Cell | lag 1 | lag 2 | lag 5 | lag 10 |
| --- | ---: | ---: | ---: | ---: |
| Primary baseline | +0.548 | +0.322 | −0.216 | −0.058 |
| Primary candidate | −0.617 | +0.263 | +0.132 | −0.359 |
| Secondary baseline | +0.199 | −0.114 | −0.034 | +0.184 |
| Secondary candidate | +0.395 | +0.204 | −0.046 | −0.002 |

The linear timing slope across sample index has median (range) +2,214.8
(+2,115.0 to +2,422.1) ns/sample for primary baseline, +122.7 (−15.4 to
+226.3) for primary candidate, −659.2 (−769.5 to −580.1) for secondary
baseline, and −878.1 (−1,189.3 to −855.8) for secondary candidate. These
statistics describe the recorded sequence and do not establish why it has
these bands.

Process order does not confine the secondary result to one adjacent ordering.
There are four secondary baseline-first pairs and five candidate-first pairs;
their full-window paired p50 medians are +5.64% and +6.61%, respectively.
For primary, the five baseline-first pairs give −3.24% and the four
candidate-first pairs give −3.64%. The groups are intentionally described as
small, unbalanced strata rather than an order-adjusted estimate.

The manifest's total process time is much larger than the sum of the 50
timed intervals. For secondary, the median baseline process is 0.194502 s,
of which 0.063213 s is the timed sum and 0.131240 s is the residual; the
candidate values are 0.195826 s, 0.064407 s, and 0.131509 s. The residual is
not an oracle-only duration: it includes expected-output preparation, oracle
controls, output inventory and validation, JSON serialization, process
startup, and other untimed work. It only shows that the untimed lifecycle
is substantial relative to the timed work.

The sealed probe source supplies the lifecycle interpretation. Its
`timed_format` measures `measured_public_format` and contains no
`output_sample` call. The warmup loop invokes the timed operation and drops
the result without calling `output_sample`; the measured loop calls
`output_sample` after the timer has stopped. The source bytes, expected output,
and oracle controls are prepared before the warmups. The input file is read
once before that loop, so the retained native timings do not measure cold file
I/O. Warmup timings are not retained, so the cold-to-warm transition itself
cannot be inferred from these reports. The returned output vector is
validated after each measured interval and then dropped; the retained sample
contains its digest, inventory summary, and oracle witness.

Before another optimization trial, qualify the lifecycle with fixed control
arms:

1. Keep the current after-each-sample oracle cadence as one arm. Add a
   deferred-oracle arm that performs the same complete oracle for every output
   after the measured block, with an explicit bound and separate accounting
   for retained output bytes. Compare both arms on the unchanged accepted
   owner before staging the archived candidate.
2. Match warmup and measured cadence in a separate arm by running the same
   post-timer validation path during warmups, while recording warmup timings.
   Pre-register the warmup count and retain a zero-oracle warmup control.
3. Balance and rotate which variant runs first in each pair, record monotonic
   process start and finish times, and keep independent process pairs. A fixed
   validation cadence such as every sample versus once after the block can be
   compared only with the same output-retention and allocation accounting.
4. Require every measured output to pass the existing byte, directory, and
   semantic oracle before accepting a result. Do not select a favorable sample
   window or infer a mechanism from these post-hoc bands.

This audit leaves the 0735 rejection and production source unchanged.
