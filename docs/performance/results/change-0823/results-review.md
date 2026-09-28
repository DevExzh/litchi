# 0823 results review

## Decision

**Reject the candidate for adoption.** The frozen benefit gate is not met. The
terminal validator accepts the evidence packet, but that status means the
packet is internally valid; it does not override the policy disposition.
Production is restored to the exact before source.

## Scope and custody

The trial is bound to base `b76786208d04310b4a033b39d70de439a64fcf69` and the
single production file `crates/litchi-pptx/src/shape/reader.rs`. The archived
candidate is exact: before SHA-256
`58e88e86ffc9d4053b694d7928377a05beacbfcb47368cefc9d2fc229b1af26e`, after
SHA-256 `a08a48f25c9e2e5f367bd67c579614f1c0c1311784ec81dd30d201b69c3db7f2`,
and patch SHA-256
`e4a97bd62ec14043922ba0babeb08bcf00a0e6ce110205d4ff67d7d14cd347c3`.

The admitted packet contains the frozen 38 qualification, 228 native, and 76
allocation reports: 342 reports and 7,106 samples. `analysis.json` and
`raw-audit.json` cover all 19 rows, with raw samples, paired cross-leg order,
bootstrap seed/endpoints, semantic oracles, and serial chronology checked.
Before/after quality receipts and the repaired probe-quality receipt pass. The
first native attempt is retained under `native-interrupted-0` and excluded
after the unexpected analysis-agent commit changed HEAD; its partial reports
were not pooled. The first analysis-reader failure and its source archive are
also retained; the corrected analysis and independent raw audit pass.

## Blocking result

The frozen eligible set is the twelve synthetic commit/lifecycle rows plus
`real/real/direct`; capture rows are negative controls. The best eligible row
is the real file, with paired p50 after/before ratio
`0.976061924` and bootstrap 95% interval `[0.974797930, 0.984335705]`.
That is a 2.394% median improvement, below the required `<= 0.97` ratio
(3% improvement), even though its interval is below 1. The full 19-row audit
therefore produces `benefits: []` and `benefit_satisfied: false`.

There are zero latency vetoes and zero allocation-resource violations. Those
passing guards do not compensate for the missing benefit gate, so the policy
result is rejection. The code-generation receipt confirms the intended
`Scene::read_with` symbol and removal of the old wrapper call in the candidate;
it does not establish a public speedup.

## RSS and spread limits

RSS is a diagnostic from process maximum resident set size, not an adoption
metric. The real row's paired RSS median ratio is `1.044440316` with interval
`[0.982268878, 1.054324958]`; no RSS review trigger is recorded, and this
packet cannot support an RSS-saving or RSS-increase claim. The native p99/p50
spread flag is set on 11 of 19 rows; those flags are diagnostic and do not
change the p50 benefit decision. No universal, cross-format, cold-cache,
tail-latency, historical-speedup, or other performance claim follows from
this trial.
