# Independent review and final checks

The tooling reviewer inspected the capture matrix, filename patterns and both
analysis outputs read-only. All 480 output/stderr pairs are unique; the 48
stage/case statistic maps agree exactly; source contrasts and restoration pass.
There are 48 native invocations, each with 100 measured owners (4,800 total),
plus 432 repeated-query processes.

The reviewer identified two assertions missing from the independent audit:
deterministic output/stderr filename binding and complete native metadata checks.
Root added filename/containment checks and assertions for probe, query/warmup/
sample counts, fresh-owner mode, all per-owner/per-query agreement flags and
sample ordinals. The initial placement of the per-owner agreement flag was
incorrect; its failing version and log remain in audit-schema-correction/.
The corrected independent audit passes and still reproduces identical statistics.

The primary analyzer, plan and measurement scripts were frozen before capture.
The independent audit/source guard were finalized separately and are covered by
the terminal artifact seal. They are not claimed to have been frozen with the
measurement tools. None of the frozen files or raw measurements changed.

After removing the two owned external roots, seven terminal commands pass:
source guard, primary analysis, independent audit, six verifier controls, report
classification, structural claim registry and non-iWork gate. Analysis JSON and
Markdown, audit statistics and verifier receipts replay byte-identically.
The library remains exactly at baseline; this diagnostic packet makes no
retention or support decision and does not supersede the 0723 rejection.

The final source/report reviewer reconciled all counts, tabulated percentages,
missing-target stage values, Simple q2 values and the maximum 3.05% mirrored
repeat p50 drift. Root clarified that the scan-timing contrast is layout →
selection, named the target-checkpoint probe to distinguish ordinary scan cursor
construction, explicitly preserved worksheet lookup/error order in the empty-slot
proposal, and labeled the allocator observation as prior 0723 evidence. No
remaining factual or causal-claim issue was identified.
