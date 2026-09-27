# 0797 final results review

This review checks the finished packet against the retained main report,
`analysis.json`, `summary.md`, `root-native-audit.json`,
`root-cg-totals.json`, `decision.json`, and the failure custody audit. It is a
bounded evidence review; it does not rerun any build, test, capture, profiler,
or replay command.

## Evidence reconciliation

The packet is internally consistent. The analyzer records 792 native children
and 23,760 native samples, 264 Callgrind children, 264 positive dumps, and 264
empty termination dumps: 1,056 reports and 24,024 measured samples. The
independent native audit reconstructs all 33 literal inputs, checks each
accepted sequence and first error, recomputes every checksum, and produces 66
case/mode rows. The final validator confirms exact agreement between those 66
rows and the analyzer's paired ratios and confidence intervals.

The independent raw-counter audit contains 528 rows. All 264 positive dumps
are nonzero and all 264 termination dumps are zero, with self rows, summaries,
and totals conserved for `Ir`, `Bc`, `Bcm`, `Bi`, and `Bim`. The analyzer reports
264/264 qualified named owners, and no owner qualification failure is promoted
to a result. These counters remain guest diagnostics; they do not establish
native cycles, allocation API calls, or phase fractions.

The semantic oracle passes for every retained report, including the baseline
and candidate comparison, clone transitions, and terminal behavior. Both
iterator layouts measure 120 bytes. The initial candidate test-only compiler
failure is fully accounted for by `failure_audit.py`: the failed logs are in
`quality-failed-0/`, the exact failed mirror is in `test-src-failed-0/`, and the
one-line closure conversion is the only difference in the corrected archived
test source. It is not a discarded performance result.

## Performance result

The candidate clears the narrow one-attribute advancement condition. For
`distinct-1/consume`, the candidate/baseline native ratio is `0.642317`, with
the frozen bootstrap interval `0.621102–0.643662`, exceeding the required 3%
benefit and remaining below one at the upper endpoint.

The same evidence rejects the candidate's replay tradeoff for workflow
advancement. Twenty consumption rows meet the frozen diagnostic regression
rule. The clearest boundaries are:

| Case | Ratio | Bootstrap interval | Guest Ir before → after |
| --- | ---: | ---: | ---: |
| `distinct-2/consume` | 1.429871 | 1.426508–1.453885 | 910 → 1,312 |
| `distinct-3/consume` | 1.333872 | 1.310475–1.353962 | 1,286 → 1,711 |
| `duplicate-valid-after-1/consume` | 1.389820 | 1.356750–1.412975 | 754 → 1,133 |
| `duplicate-long-quoted-after-1/consume` | 1.366083 | 1.355620–1.391794 | 754 → 1,133 |
| `syntax-unique-tail-after-4/consume` | 1.196146 | 1.171540–1.210534 | 2,594 → 3,042 |

The analyzer also retains 22 process-p50 spread flags across 132 case/mode/leg
groups. These are stability diagnostics rather than an additional adoption
gate; the six paired block ratios and their frozen intervals remain the
decision statistic.

The long duplicate cases confirm the semantic goal: the candidate
does not scan the 4,096-byte duplicate value before returning the duplicate
error, yet replay and checked dispatch still make those multi-attribute inputs
about 1.37 times slower. Construction has no frozen regression flags, but its
opaque, forced-materialization timing is diagnostic only.

`decision.json` therefore records `advance_to_workflow_trials: false`, while
retaining `single_attribute_benefit: true` and the 20 consume regression rows.
The candidate is archived and the production baseline is retained. No public
workflow, resource, cross-format, or production-adoption claim follows from
this packet.

## Review conclusion

The main report's disposition is supported by independent semantic, numeric,
and raw-counter evidence: the one-attribute fast path is real, but the
first-attribute replay is too costly across common multi-attribute and early
error shapes for this preflight to advance. A future candidate should measure
real caller attribute-count and consumption distributions before specializing
this boundary, while retaining early duplicate ordering and the existing
bounded hostile-input handoff.
