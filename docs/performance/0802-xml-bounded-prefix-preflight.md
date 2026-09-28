# 0802 — bounded linear names after the no-replay short prefix

The candidate is rejected: four protected syntax-error cases violate the
frozen regression guard. Production remains unchanged.

This candidate addresses the two costs exposed by 0801: empty consumption and
an ordered map started after only three successful attributes. It is tested
against the unchanged production source at `334a22f101`; historical timing
samples are not pooled and no public-workflow speedup is inferred.

Exactly empty raw tails start in the fused Done state. Whitespace and malformed
tails still reach lexical processing. The first two attributes and third-key
precheck retain the no-replay design. When a third successful item has a
non-whitespace tail, the candidate seeds one boxed fixed array of 32 borrowed
names. Linear checks occur before values, and successful names are stored
without a second linear search. A unique 33rd item either finishes the iterator
when only whitespace follows, or seeds the ordered map from the stored keys.
Later names use ordered checks. No values are copied or replayed.

The backend enum is stored directly in the phase, avoiding a redundant outer
box. Its array still has initialization/copy costs, and its linear comparisons
and eventual map seeding remain work to measure. The measured whole iterator
is 128 bytes versus the baseline's 120; these numbers do not measure heap use.
The bounded linear stage retains the 32-name cap, with ordered checking for
hostile long tags. Comparison tests instrument both stages.

## Frozen protocol and correctness

The [packet](results/change-0802/README.md) retains all 39 literal cases, 18
protected consume cases and acceptance thresholds from 0801 and 0799. Six
paired native blocks use 30 samples, three warmups and 4,096 iterations per
process on CPU 12. Process p50 uses nearest rank; median paired ratios use
10,000 bootstrap resamples with fresh seed 802080. Two separate Callgrind
repeats retain guest instruction/branch counters, not native time or allocation
API counts. The probe includes opaque iterator construction and checksum work.

Advancement requires at least 3% benefit for both one- and two-attribute consume
cases, with ratio intervals wholly below 1, and no protected consume regression
above 5% with its interval wholly above 1. All other regressions remain review
triggers. Preflight cannot authorize production adoption: fresh workflow,
resource and cross-format trials are still required even when it passes.

All six isolated helper gates pass: 70 baseline and 100 candidate tests, zero
failures or ignores, and warnings-denied Clippy for both legs. These exact
five-copy mirror crates exercise the shared unusual-key/error-position matrix,
clone/fusion transitions, differential/random cases and comparison bounds.
The release probe passes formatting, build, check, Clippy, the literal fixture
audit and semantic self-check, including clone advances 0, 1, 2, 3, 4, 5, 32
and 33. These are helper/probe gates, not full production-workspace verification.

The first quality attempt passed tests but candidate Clippy requested the
question-mark operator in one optional key extraction. The failed source,
patch, logs and mirror projects are retained; the five-copy correction is
semantically equivalent. There were no failed direct builds or failed native
captures. The frozen trait-method comment has one stale immediate-map sentence;
implementation, module documentation and the source review show the actual
bounded linear stage. The tested source archive is retained exactly.

## Results and decision

Ratios are candidate/baseline. Both required short-tag benefits qualify, and
the empty consume case improves, but four protected syntax cases veto
advancement. These are direct-helper measurements only.

| Case | Mode | Paired p50 ratio | Change | 95% bootstrap ratio interval |
|---|---|---:|---:|---:|
| `distinct-0` | construct | 3.177667 | +217.767% | [2.994513, 3.372681] |
| `distinct-0` | consume | 0.675011 | -32.499% | [0.674321, 0.676301] |
| `distinct-1` | consume | 0.799373 | -20.063% | [0.795323, 0.804328] |
| `distinct-2` | consume | 0.928650 | -7.135% | [0.908952, 0.939217] |
| `distinct-3` | consume | 1.024019 | +2.402% | [1.014697, 1.027424] |
| `distinct-4` | consume | 1.302379 | +30.238% | [1.287646, 1.323957] |
| `distinct-16` | consume | 1.275080 | +27.508% | [1.251021, 1.287138] |
| `distinct-32` | consume | 1.379962 | +37.996% | [1.375706, 1.394342] |
| `distinct-33` | consume | 0.769327 | -23.067% | [0.725713, 0.773088] |
| `distinct-64` | consume | 1.107857 | +10.786% | [1.095758, 1.130985] |
| `syntax-flag-after-0` | consume | 1.276295 | +27.629% | [1.243891, 1.315446] |
| `syntax-equals-value-after-0` | consume | 1.347073 | +34.707% | [1.296092, 1.359071] |
| `syntax-flag-after-2` | consume | 1.076772 | +7.677% | [1.059236, 1.080236] |
| `syntax-equals-value-after-2` | consume | 1.072755 | +7.275% | [1.052862, 1.082765] |

All 19 significant consume regressions remain visible: distinct counts 4, 8,
16, 17, 32 and 64; valid duplicates after 4 and 33; an unquoted duplicate after
33; flag syntax after 0, 2, 4 and 33; unique-tail syntax after 4 and 33; and
equals-value syntax after 0, 2, 4 and 33. The four protected vetoes are flag and
equals-value syntax after zero and two accepted attributes. The other fifteen
remain material review triggers. See the [complete table](results/change-0802/summary.md)
for all 78 case/mode pairs.

All 39 construction rows also flag significant regressions, with paired median
ratios from 3.176686 to 3.406122. Construction is diagnostic rather than an
independent advancement veto, but this is a material limitation. Its measured
scope includes the opaque iterator, its drop, dispatch and common checksum
work; it is not an isolated constructor instruction benchmark or a forecast
of public-workflow latency. Sixty-six of 156 case/mode/leg process-p50 groups
have spread above 5%; none are dropped from the summaries.

The empty shortcut removes parser advancement on the exact empty path, but
this combined design does not preserve protected malformed-input costs.
The bounded linear stage also leaves substantial middle-size regressions.
No historical native samples are pooled to assert improvement over 0801,
and the result does not justify another one-case fix followed by adoption.
Constructor/state handling and linear comparison costs need separate controls
before a further combined candidate. Source inspection shows that the linear
stage currently uses ordering comparison to test equality; whether changing
that helps must itself be measured.

All 312 counter owners qualify, each with one positive and one empty termination
dump; independent scalar conservation passes for all 624 dumps and all five
events. In both repeats, opaque construction uses 95 guest instructions versus
70 for the baseline, while empty consumption uses 134 versus 170. Flag syntax
before any attribute rises from 253 to 291 instructions; equals-value syntax
after two rises from 1,006 to 1,289. Thirty-two-attribute consumption rises from
25,562 to 35,997 instructions. These separate guest diagnostics identify work
changes but do not establish native-time causality, phase fractions, hardware
branch rates or allocation API counts.

The packet contains 936 native reports and 28,080 samples, plus 312 profile
reports with one measured sample each: 1,248 reports and 28,392 samples in total.
Independent input/error/checksum and paired-statistic audits agree with all 78
analysis rows. Post-cleanup replay passes. The owned target was removed,
freeing 178,414,599 logical bytes; its exact final executable identity is
retained, and no failed-build executable existed.

All 9,196 production source hashes, 35 architecture inputs, unrelated files
and other worktrees remain unchanged. The candidate is archived without
workflow advancement or production adoption. No memory, cold/range, concurrency,
producer or CRUD coverage is promoted. iWork remains excluded, and the broader
GOAL remains incomplete.
