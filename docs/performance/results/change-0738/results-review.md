# 0738 results review

## Scope and disposition

This review keeps the two measurement phases separate. The main phase tests the
0738 retained-`Sample` layout change against the archived 0735 probe and the
0737 controls probe. The supplemental phase tests startup argument controls
using the already-built archive and prior binaries. It does not pool samples,
processes, or statistics across phases.

Both phases leave production unchanged. Neither phase authorizes restoring the
rejected 0735 candidate or transfers a harness result to an ordinary
production path.

The main phase passed its analyzer and independent audit over 72 native
processes, 3,600 native samples, and 24 allocation processes. Its secondary
archive-to-prior p50 difference remains a failure of the 5% review gate, while
the prior-to-0738 layout contrast is small. The supplemental phase passed its
capture validation over another 72 native processes, 3,600 samples, and 24
allocation processes. It shows that harmless extra command-line setup can
change later owner timing within the same binary, with fixture-dependent
effects. This is a harness sensitivity result; it does not identify the
allocator, cache, executable-layout, or 0735 candidate cause.

## Main layout phase

The frozen main matrix used four legacy arms (`archive`, `prior`,
`restored-a`, and `restored-b`), both PPT fixtures, CPU 12, serial execution,
nine paired native repeats, and three allocation repeats. Every report had the
sealed 0735 output, inventory, and full visible oracle projection. The
independent validator reports `analysis_match: true`; all 17 deliberate
qualification-report corruptions were rejected.

The full-window paired p50 results are:

| Fixture and contrast | Median change | Bootstrap 95% interval | Pair range | p50 flags |
| --- | ---: | ---: | ---: | ---: |
| Primary archive → prior | +0.989% | [+0.802%, +1.505%] | [+0.355%, +1.580%] | none |
| Primary prior → restored-a | −2.005% | [−2.502%, −0.857%] | [−2.550%, −0.843%] | none |
| Primary restored-a → restored-b | +0.258% | [−0.999%, +0.513%] | [−1.086%, +0.625%] | none |
| Secondary archive → prior | +7.105% | [+5.192%, +7.591%] | [+5.112%, +7.604%] | all 9 |
| Secondary prior → restored-a | +0.275% | [−0.038%, +0.356%] | [−0.384%, +0.463%] | none |
| Secondary restored-a → restored-b | −0.065% | [−0.285%, +0.240%] | [−0.396%, +0.298%] | none |

The other full-window flags are limited to secondary tail outliers: prior →
restored-a has one p99/maximum flag (pair 0), and restored-a → restored-b has
p99/maximum flags on pairs 0, 1, and 7. No primary mean, p95, p99, or maximum
comparison is flagged; the secondary archive → prior p50 is the only repeated
full-window gate failure. Ordered first-ten, middle-thirty, and last-ten
windows remain descriptive diagnostics. Their flags are not additional
independent decisions: primary archive → prior has first-ten flags on all nine
pairs, primary prior → restored-a has first-ten flags on pairs 1 and 4 and
last-ten flags on pairs 1–8, and secondary archive → prior has first-ten flags
on all nine pairs plus middle-thirty flags on pairs 3, 4, and 7.

Removing `retained_witness_count` from the retained `Sample` and deriving it
at the receipt boundary therefore does not reproduce or remove the secondary
archive-to-prior effect: prior → restored-a is +0.275% with an interval that
crosses zero, while archive → prior is +7.105%. The source-equivalence proof
also binds the owner and oracle blocks; the result cannot be attributed to an
oracle-scope change.

All 24 main allocation reports have identical boundary-relative owner fields
within each fixture and across every arm. The values are:

| Fixture | Allocated bytes | Deallocated bytes | Calls | Peak live bytes | Retained bytes |
| --- | ---: | ---: | ---: | ---: | ---: |
| Primary | 11,391,768 | 11,001,624 | 5,658 | 1,992,885 | 390,144 |
| Secondary | 10,444,631 | 10,159,447 | 18,662 | 1,976,223 | 285,184 |

These counters exclude the external oracle and receipt work and do not measure
RSS, total probe heap, cache state, or a timing cause.

## Supplemental argument-control phase

The supplemental plan froze the same CPU, fixtures, serial order, sample
counts, warmups, nine native repeats, and three allocation repeats. It used
only the existing archive and prior binaries and changed no Rust or production
source. Its four arms were:

* `archive-standard`: archive binary with the ordinary command;
* `archive-extra`: archive binary with a redundant `--operation format` pair;
* `prior-default`: prior binary with the default lifecycle flag omitted; and
* `prior-explicit`: prior binary with the explicit `--lifecycle legacy` pair.

The redundant operation pair and lifecycle pair each use an option token of 11
characters and a value token of 6 characters. Within-binary comparisons test
the added argument and parser/setup path. Cross-binary comparisons with equal
argument counts remain confounded by the different probe binaries, source
history, code generation, and executable layout.

An independent replay of the 96 raw capture reports checked the frozen
manifest and command vectors, serial monotonicity, empty stderr, capture
hashes, report and sample schemas, every output hash/inventory/oracle value
against the sealed 0735 reference, and all retained-count values. A separate
scalar implementation recomputed the five full-window statistics, all eight
case/contrast groups, all 40 metric summaries, paired bootstrap intervals, and
the review flags. It matched `argv-control/analysis.json` exactly. The packet
preflight also rejected an altered command vector.

The p50 results are:

| Fixture and contrast | Median change | Bootstrap 95% interval | Pair range | p50 flags |
| --- | ---: | ---: | ---: | ---: |
| Primary archive-standard → archive-extra | −1.494% | [−2.120%, −0.933%] | [−2.166%, −0.515%] | none |
| Primary prior-default → prior-explicit | +7.043% | [+6.654%, +8.100%] | [+5.319%, +8.148%] | all 9 |
| Primary archive-standard → prior-default | −6.137% | [−6.360%, −5.850%] | [−6.422%, −4.886%] | 8 of 9 |
| Primary archive-extra → prior-explicit | +2.197% | [+1.765%, +2.572%] | [+1.188%, +2.734%] | none |
| Secondary archive-standard → archive-extra | +6.142% | [+5.659%, +6.423%] | [+5.068%, +6.847%] | all 9 |
| Secondary prior-default → prior-explicit | +7.781% | [+7.270%, +8.211%] | [+6.749%, +8.793%] | all 9 |
| Secondary archive-standard → prior-default | −1.045% | [−1.701%, −0.190%] | [−1.870%, −0.037%] | none |
| Secondary archive-extra → prior-explicit | +0.719% | [−0.007%, +1.045%] | [−0.129%, +1.141%] | none |

The complete supplemental full-window flag set is as follows. Unlisted
metrics have no flags.

* Primary archive-standard → archive-extra: p99 and maximum, pair 0.
* Primary prior-default → prior-explicit: p50, all pairs; mean, pairs 0, 4,
  7, and 8.
* Primary archive-standard → prior-default: p50, pairs 0, 2–8; p99 and
  maximum, pair 0.
* Secondary archive-standard → archive-extra: p50, all pairs.
* Secondary prior-default → prior-explicit: p50, all pairs; p99 and maximum,
  pair 2.
* Secondary archive-extra → prior-explicit: p99 and maximum, pair 2.

The eight supplemental allocation contrasts have zero difference for every
field on all three repeats. They therefore supply no allocator change that
could explain the timing differences.

The same-binary results establish that changing only the harmless command-line
setup changes later timed owner behavior in this harness: the archive
secondary p50 moves +6.142% when the redundant operation pair is added, and
the prior p50 moves +7.043% on primary and +7.781% on secondary when the
explicit lifecycle pair is added. They do not establish whether the mechanism
is argument parsing, process state, stack or allocator state, cache state, or
another startup effect. The equal-argument cross-binary secondary contrasts
are small (+0.719% with an interval crossing zero for extra → explicit, and
−1.045% for standard → default), but they are not binary equivalence proofs.
The primary equal-argument contrast remains flagged (−6.137%), so the
fixture-dependent and cross-binary confounding must remain visible.

## Statistical and causal limits

Each native comparison uses nine independent paired processes. Its bootstrap
resamples process-pair changes, not the 50 ordered samples inside a process.
The nearest-rank p99 of a 50-sample process is its maximum, so p99 and maximum
are reported separately for schema completeness but are not independent tail
measurements. The supplemental phase does not add order-window decisions.

Allocation processes have one owner sample and are reported only as
boundary-relative counters. No result here supports an RSS, heap-size, cache,
allocator-cause, or latency claim from those fields. The CLI controls also
cannot explain the original 0735 candidate regression: they use unchanged
archive/prior probe binaries, and their effects differ by fixture and command
pair. Production remains at the accepted source census, and the rejected
candidate remains absent.
