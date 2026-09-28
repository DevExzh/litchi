# 0803 results review

This review covers the terminal 0803 packet: `summary.md`, the raw native
reports, `analysis.json`, `root-native-audit.json`, the Callgrind conservation
totals, quality receipts, and custody/cleanup records. I did not rerun a build,
test, native capture, or profiler.

The packet compares the rejected 0802 candidate source with the same helper
implementation after moving the exact-empty check from construction to
`next_first`. Its before leg is not current production. The terminal decision
keeps `advance_to_workflow_trials` and `production_adoption` false and records
that no production-baseline comparison is made. No historical ratio is
multiplied into these results.

## Independent raw paired p50 check

I read all 936 native JSON reports independently. Each report has 30 elapsed
samples; I sorted each sample vector and used element 14, the frozen nearest
rank p50. For the selected rows below, I paired the six block p50s by block,
computed each after/before ratio, and took their median. The values match the
recorded analysis and its deterministic 10,000-resample bootstrap (seed
`803080`, endpoints 250 and 9,749).

| case and mode | before block p50s | after block p50s | median ratio | 95% bootstrap interval |
| --- | --- | --- | ---: | --- |
| `distinct-0` construct | `[20060, 20020, 20080, 20010, 20000, 19081]` | `[5931, 7150, 6370, 5930, 6380, 5930]` | `0.3140057166` | `[0.2960074175, 0.3380714286]` |
| `distinct-0` consume (empty) | `[24550, 24570, 24520, 24510, 24550, 24550]` | `[13150, 12730, 12731, 12750, 12740, 12730]` | `0.5190748730` | `[0.5183225615, 0.5279186931]` |
| `distinct-1` consume | `[88680, 87591, 88421, 87561, 88460, 88481]` | `[67501, 67340, 67321, 67331, 67340, 67320]` | `0.7613084636` | `[0.7610081609, 0.7688807573]` |
| `distinct-2` consume | `[162321, 162151, 160321, 162090, 160571, 160791]` | `[140311, 139971, 140131, 140191, 140200, 140081]` | `0.8680476520` | `[0.8638092002, 0.8735995799]` |

These are process-level p50 diagnostics for named construction and consumption
owners, including their common dispatch/checksum work. They do not establish a
public workflow latency or a production speedup.

## Coverage, flags, and uncertainty

The native lane contains 39 cases × 2 modes × 6 paired blocks × 2 legs, with
30 samples per child: 936 reports and 28,080 elapsed samples. All 78 paired
case/mode rows have the recorded bootstrap policy. Recomputing the diagnostic
rule (median ratio above `1.05` and bootstrap lower endpoint above `1.0`)
produces zero regression flags, matching `root-native-audit.json` and
`analysis.json`.

The absence of a diagnostic flag does not make this a qualification result.
Process p50 spread exceeds 5% in 73 of 156 case/mode/leg groups. Across the
paired rows, construction medians range from `0.296307` to `0.372513`; consume
medians range from `0.519075` to `1.005577`. The three consume medians above
one remain below the diagnostic threshold and have bootstrap lower endpoints
below one. The spread and the named micro-input scope require these ratios to
remain diagnostic rather than general latency claims; no p95 or p99 claim is
made.

The iterator size witness reports 128 bytes for both baseline and candidate
in every native and profile report. This is a measured layout observation,
not a heap-allocation or memory-use result.

## Callgrind and quality bounds

The profile lane contains 312 positive owner dumps and 312 empty termination
dumps. All 312 owners qualify, no qualification failure is converted to a
success, every `Ir`, `Bc`, `Bcm`, `Bi`, and `Bim` total is conserved, and all
termination totals are zero. Selected guest instruction totals are `95 → 70`
for empty construction, `134 → 122` for empty consumption, and `35,997 →
35,978` for 32-name consumption. These counts are mechanism diagnostics from
Callgrind; they are not native timing, cycle attribution, allocator API
counts, or phase fractions. Inclusive descendant values overlap.

Both five-copy helper mirror legs pass their 20 tests per copy and warnings
denied Clippy checks. The retained `quality-failed-0/` attempt records the
earlier stale private-state assertion before the after test was amended for
the intentional `First`-then-`Done` transition. The final quality receipts
are successful, but they cover minimal helper crates rather than the full
production workspace.

The cleanup record retains the final binary identity, removes the owned target
directory and its 178,371,844 logical bytes, and records no failed-build
binary. The source and candidate custody records keep the before leg bound to
the rejected 0802 archive; no production source was changed.

## Disposition

Moving the exact-empty check improves the named construction and short consume
diagnostics relative to the rejected 0802 control, while consume rows at the
33-name boundary stay near parity. This packet provides no basis for workflow
advancement or production adoption. A separately controlled linear-equality
comparison is needed before any fresh comparison against production is
considered.
