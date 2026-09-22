# 0728 result review

The completed packet is a current-source baseline, not a before/after
benchmark. It contains 81 successful processes: six native processes for each
of the nine case/route cells and three allocation-instrumented processes for
each cell. The independent audit reproduced the matrix and statistics and
reports `analysis_match: true` in [`audit.json`](audit.json). The timing
summary is in [`analysis.md`](analysis.md); the full p95, p99, maximum and
allocation records are in [`analysis.json`](analysis.json) and
[`report-summary.json`](report-summary.json).

## What the timings establish

The table gives the range of whole-operation p50 values across the six native
processes in each cell. Stage and finish are per-process medians of their
within-sample fractions of the measured common-container whole window.

| Case | Route | Whole p50 (µs) | Stage (%) | Finish (%) |
| --- | --- | ---: | ---: | ---: |
| docfloat | public format | 984.82–1089.01 | — | — |
| docfloat | container Reuse | 243.05–247.64 | 53.83–54.10 | 30.13–30.67 |
| docfloat | container Rewrite | 166.80–177.67 | 53.80–54.22 | 23.02–24.94 |
| docnohf | public format | 108.01–110.83 | — | — |
| docnohf | container Reuse | 43.88–44.80 | 56.52–57.42 | 28.95–29.29 |
| docnohf | container Rewrite | 32.98–33.68 | 56.06–56.84 | 24.50–24.97 |
| ppt45543 | public format | 1128.81–1140.66 | — | — |
| ppt45543 | container Reuse | 271.53–278.34 | 43.58–43.77 | 17.96–18.78 |
| ppt45543 | container Rewrite | 210.36–211.63 | 41.66–41.83 | 9.61–9.74 |

For the DOC common-container Reuse controls, finish occupies about 29–31% of
that control's measured whole window. This is consistent with the traced
DOC route: the stage performs candidate render, reopen, recapture and
discovery, and the later finish materializes the changed package again. The
percentage cannot be assigned to the public DOC format save. The public-format
and common-container routes run as separate processes with different owners;
their medians are not nested measurements, so subtracting or dividing them
would not produce a public-save phase fraction.

The Rewrite control is faster in all three common-container cases. That is a
policy comparison under the current probe, not an optimization result for the
preservation path. The raw directory oracle records a physical normalization
difference for every Rewrite output while the source-to-expected model still
matches:

| Case | Reuse p50 (µs) | Rewrite p50 (µs) | Rewrite expected/output normalized difference (bytes) |
| --- | ---: | ---: | ---: |
| docfloat | 243.05–247.64 | 166.80–177.67 | 394 |
| docnohf | 43.88–44.80 | 32.98–33.68 | 40 |
| ppt45543 | 271.53–278.34 | 210.36–211.63 | 51 |

The corresponding raw reports mark `source_output_ok: false` under the
Rewrite report-only policy gate, while the semantic, stream, and expected
output checks pass. Reuse and public-format outputs have zero normalized raw
directory difference in these cells. Rewrite therefore remains a diagnostic
control until an explicit preservation policy permits its normalization; its
lower timing cannot be used as evidence that the current Reuse path is safe to
replace.

The PPT common-container cells are an alternative control for the shared OLE2
editor workload. They are not a decomposition of the public PPT slide-order
save. The slide-order owner uses its private embedded editor and its own final
write, full snapshot reopen, and semantic/persisted-record checks. The PPT
stage and finish percentages above must therefore not be presented as public
PPT save fractions or as evidence that a common-editor artifact can be handed
to that route.

## Oracle and allocation reading

The retained raw reports show stable output identities within each route, the
expected changed streams, zero source-to-expected normalized directory
difference, and the final semantic witnesses for DOC and PPT. Every named
negative control is rejected with a reason. The allocation lane confirms the
same route structure and reports per-region allocation counters, but its
`peak_live_bytes` and `retained_bytes` are region measurements. They are not
whole-process RSS and do not establish the memory cost of retaining an
additional rendered CFB artifact in a production editor.

The evidence is consequently sufficient to attribute a duplicate common-DOC
render/recapture window for the current route, subject to the phase boundary
above. It does not establish a safe cache lifetime, an output budget for a
retained `Vec`, or equivalence after another edit changes the candidate.

## Bounded next direction

The next attribution step should split the existing common-container stage
and finish windows into render, reopen, recapture, discovery, and final-write
subphases. That isolates the work a future handoff could actually remove and
keeps the current 0663 preservation and validation behavior as the baseline.

If that attribution justifies an optimization, evaluate a one-shot rendered
handoff at the DOC batched common-editor boundary, preferably a bounded
`put_streams_shared_with_rendered`-style API. The handoff should bind the
artifact to the exact source identity, candidate generation, target catalog,
policy and limits; invalidate it on every later stream, topology, metadata or
policy mutation; publish candidate state and rendered bytes atomically; charge
the retained output against an explicit budget; and consume it once. The
outer DOC public reopen and semantic validation remain required. A budget
refusal must fall back to the current writer, and no editor-wide unbounded
artifact cache is justified by these results.

The proof matrix for that future experiment should include exact no-ops,
same-length and length-changing edits, multiple edits after staging, failed
render/reopen/discovery, source-freshness changes, policy fallback, and budget
refusal. PPT slide-order should remain a separate owner-specific experiment.

This review uses the current 0663 implementation and the frozen packet's
oracle contract. It recommends no production retention or reuse change from
the baseline alone.
