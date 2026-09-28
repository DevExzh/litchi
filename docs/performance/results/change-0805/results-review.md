# 0805 results review

The captured evidence passes the frozen production-preflight gates. The
decision is `advance_to_workflow_trials: true` and
`production_adoption: false`. This is a fresh production-versus-combined
candidate comparison; no 0798, 0802, 0803, or 0804 timing ratio is pooled or
multiplied.

## Raw native audit

The capture contains 39 cases, two modes, two legs, six alternating blocks,
30 samples per process, three warmups, and 4,096 iterations per sample:
936 native reports and 28,080 native samples. The process p50 is the frozen
nearest-rank value, and each case/mode ratio is the median of six paired
after/before p50 ratios. The 95% intervals are the 10,000-resample median
bootstrap with seed 805080 and ranks 250 and 9749.

I independently recomputed the nearest-rank p50 from every raw native JSON
report, checked every sample's checksum, accepted count, and error marker
against its report oracle, and compared the resulting vectors with the frozen
analysis. All 936 reports and all 156 case/mode/leg p50 vectors matched. The
independent root audit also matches all 78 case/mode ratios and intervals.

The two required consume benefits pass:

| Case | Ratio | Change | 95% interval |
| --- | ---: | ---: | ---: |
| `distinct-1` | 0.613544 | -38.646% | 0.588405–0.617521 |
| `distinct-2` | 0.809200 | -19.080% | 0.791095–0.809967 |

Both intervals are wholly below 1 and both ratios are at most the 0.97
benefit threshold. `distinct-0` is also a large improvement at 0.318976
(95% interval 0.318807–0.336264), but it is not one of the two required
dominant-class gates.

The six diagnostic flags are exactly the following. The two construct flags
and four consume flags remain visible review triggers:

| Case | Mode | Ratio | 95% interval |
| --- | --- | ---: | ---: |
| `distinct-17` | construct | 1.136238 | 1.036975–1.159219 |
| `distinct-3` | construct | 1.129848 | 1.000927–1.164418 |
| `distinct-4` | consume | 1.226491 | 1.198300–1.243788 |
| `duplicate-valid-after-4` | consume | 1.107260 | 1.095358–1.119992 |
| `syntax-equals-value-after-4` | consume | 1.286817 | 1.280224–1.297670 |
| `syntax-flag-after-4` | consume | 1.286344 | 1.209093–1.299985 |

The four consume flags are outside the frozen 18-case protected set. The
protected rows were checked directly against both veto conditions (ratio above
1.05 and lower interval endpoint above 1); none satisfies both. The largest
protected median is `syntax-flag-after-0` at 1.029831, with interval
1.003291–1.060821, so its interval alone does not make it a veto. The other
protected syntax boundary, `syntax-equals-value-after-0`, is 1.022081 with
interval 0.967258–1.064163. The remaining protected duplicate and malformed
tail cases are below 1 in median ratio.

There are 41 process-p50 spread flags among 156 case/mode/leg groups: 33
construction and eight consumption. They remain retained diagnostics and do
not alter the frozen decision rule.

## Profile and mechanism audit

The profile lane contains 312 positive reports and 312 empty termination
dumps. All 312 owners qualify with one incoming call, all five guest counters
are conserved, and no owner qualification failed. The counters are guest
Callgrind diagnostics. The lexical and allocator symbol groups are name-based
censuses; they do not establish allocator API calls, native branch rates, or
phase fractions. Inclusive descendants overlap.

The report's instruction table is a per-positive-dump total from the parsed
self-summary, and both profile repeats agree. It should not be read as the
self-cost of `CheckedAttributes` alone. In particular, the candidate's
four-attribute consume cost is 2,372 guest instructions versus 1,702 for
production, while its 33-attribute consume cost is 27,487 versus 50,231.
Those totals support the mechanism investigation but do not assign native
latency to a unique function or instruction.

The candidate enters `OwnCheck::Linear(Box<LinearCheck>)` after its short
prefix, so the bounded array is behind one box. A unique 33rd item with a
remaining tail switches to a fresh ordered-map box and seeds that map from
the 32 stored borrowed names plus the current key. This avoids production's
reparsing of the first 32 values through a second unchecked iterator; it still
performs map insertion work at the transition. Therefore the 33-attribute
improvement needs a boundary distinction: an exactly-33-item tag whose tail is
only whitespace returns from the linear stage without creating the map, while
a longer tail performs the map seeding described above. The selected
`distinct-33` fixture therefore removes both parser replay and map creation;
the result does not isolate either effect or prove that any single instruction
caused the native result.

## Boundary and disposition

The native and profile receipt identities, semantic/checksum oracles, source
manifest, and binary identity are all bound in `analysis.json` and the root
audits. The comparison remains a hot direct-helper measurement with explicit
dispatch, opaque iterator, drop, and checksum work. It does not measure public
workflow latency, memory use, resource limits, or cross-format behavior.

The result authorizes only fresh workflow, resource, and cross-format trials.
It does not authorize production adoption, and the four unprotected consume
regressions remain explicit review triggers for those trials. No current
baseline may be replaced until those broader gates pass.
