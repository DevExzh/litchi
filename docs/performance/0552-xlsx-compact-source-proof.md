# XLSX compact source-cell proof experiment

The change-0552 candidate is rejected. It reduced measured edit/save work but failed the frozen RSS, valid no-op planning, and cap-control admission gates. Baseline production has been restored; seven public regression tests and the complete experiment evidence are retained. All eleven final-source quality checks passed, including 1,313 XLSX tests. Decision, verification, and cleanup records are retained in the evidence bundle.

The candidate collected source-bound cell offsets during the existing source-backed XLSX `MultiSourceEdit` planning traversal, then reused them at commit instead of reconstructing the complete worksheet layout. It kept the full writer fallback, rewritten XML validation, and independent readback. The retained proof used 8-byte cell spans with a 2 MiB logical metadata cap. This cap did not bound all transient allocations or process RSS.

The comparison used two native and two allocation-instrumented repeats, four worksheet shapes, one-cell and one-percent edits, and both managed and unmanaged workflows. The order was baseline, candidate, candidate, retained baseline. All 32 one-percent workflow latency checks met the frozen improvement threshold, and all 32 one-cell latency guards and allocated-byte checks passed. That does not satisfy the combined optimization goal:

| Failed requirement | Recorded result |
| --- | --- |
| Process RSS at most 5% above matched baseline | Dense-sparse repeat two grew 6.97% for unmanaged one-cell edits and 8.39% for managed one-percent edits. |
| Valid no-op planning latency at most 5% above baseline | All eight p50/mean checks failed. |
| Cap-control latency at most 5% above baseline | Six of twenty p50/mean checks failed. |

The exact rows, sample vectors, individual adverse values, and repeat drift are retained in [the main analysis](results/change-0552/metrics-analysis.json) and [the guard/cap analysis](results/change-0552/guards-analysis.json). Analyzer status indicates valid evidence; it does not imply candidate admission. Allocation-instrumented elapsed time is excluded from native latency evidence.

The failed pilot makes the exact commit instruction profiles and prospective supplementary large-tag experiment unnecessary for this rejected candidate. Neither was captured, and no instruction, large-tag, hardware-counter, cold-cache, remote-provider, scaling, or single-sheet `SourceEdit` result is claimed.

The [experiment runbook](results/change-0552/README.md) documents source revisions, preserved failed checks, the metadata-only guard-analyzer serialization amendment, source restoration, and the serial reproduction workflow. OLE2 and OOXML remain the performance priority; ODF is deferred and iWork is excluded from this work.
