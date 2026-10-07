# 0955 — bounded cold XLSX compression qualification

The existing managed publication path admits two- and four-worker compression
waves for the medium one-percent workload without changing output bytes.
Dense-sparse and the one-edit control remain serial under the existing cold
threshold, exactly as the [0954 census](0954-xlsx-replacement-task-census.md)
predicted. This batch adds an opt-in standalone qualification example; it does
not change production Rust or the existing benchmark selectors and timers.
There is no latency, allocation, RSS or speedup claim.

The example uses the existing deterministic corpus builder, selected-sheet
edit/commit and source-backed publication API. Each case creates a fresh explicit
execution context and package. No caller-owned facility or ambient pool is used.
A sink samples the root budget's worker permits while publication is writing;
those permits are gone after the consuming publisher returns. Source review
confirms that the scheduled writer completes compression before its first
output write. These observations establish admitted width, not an independently
measured operating-system thread count or task overlap.

| Workload | Configured workers | Observed worker permits | Publication CPU tasks |
| --- | --- | --- | --- |
| Medium one-percent | 1 / 2 / 4 | 0 / 2 / 4 | 5 / 5 / 5 |
| Dense-sparse one-percent | 1 / 2 / 4 | 0 / 0 / 0 | 5 / 5 / 5 |
| Medium one-edit | 1 / 2 / 4 | 0 / 0 / 0 | 2 / 2 / 2 |

Zero worker permits identifies the inline serial path. CPU tasks are charged
for regenerated Deflate members on both serial and parallel paths; they are
not thread counts. The final probe samples the CPU-task delta immediately after
publication, before its inverse check.

Every context has finite root limits: 64 MiB input, 1,000,000 objects, depth 256,
1,000,000,000 work units, one concurrent input operation, the configured worker
count and exactly five or two CPU tasks. The output limit is twice the expected
archive size plus 64 KiB. The memory ceiling includes the selected source payload,
64 KiB planning allowance and four times the existing 655,360-byte worker-state
estimate. Holding the ceiling equal across widths permits comparison without
silently disabling a larger wave. It is a budget model, not measured memory use.
The respective ceilings are 2,956,761, 3,196,436 and 2,751,297 bytes for medium
one-percent, dense-sparse one-percent and medium one-edit.

All nine cases verify the frozen source and output hashes, equality with the
existing eager oracle, semantic reopen/readback, exact task and worker counts,
and zero live memory/worker/I/O/object usage after release. Accepted output bytes
remain cumulatively charged and exactly equal the sink length. Output hashes
match the retained 0954 fixtures, including every preserved byte in those ZIPs.

Each actual managed output is reopened as an owning package; the managed-origin
patch's inverse is applied with an explicit materialization ceiling and must
restore the original serialized ZIP exactly. This exercises the public bounded
owning handoff, not a parallel inverse-publication API. Both shapes additionally
pass an exact managed no-op, typed pre-publication cancellation with zero output
and zero tasks, and typed output-budget refusal with zero sink calls, accepted
bytes and retained live resources. The refusal uses the new four-worker context;
its five CPU tasks are admitted before output is refused. No claim of mid-wave
cancellation, stale/foreign-source testing or memory-driven width narrowing is
added by this probe.

The final nine-case report is byte-identical across two fresh diagnostic runs.
Scoped formatting, warning-denied Clippy for the example and the complete
crate-boundary checker pass. Independent source review approves the final
counter boundaries. The report directory refuses overwrite without changing its
manifest. No production correctness-suite rerun is claimed for this harness-only
addition. [The packet](results/change-0955/README.md) retains commands, source and
binary identities, limits, results and review.

Development history is retained: the first probe build was deliberately stopped
because root launched it from the working checkout instead of the isolated
source snapshot; no workload ran from that build. An initially qualified probe
sampled CPU tasks after its inverse check. A late proposed counter-placement
change was explicitly accepted as a new revision, then rebuilt, reviewed and
requalified with all affected checks rerun. Its final results are unchanged.
Earlier source revisions and receipts remain external; the final source binding
identifies exactly which version supports this report.

The next step is a frozen paired native experiment with the same explicit cold
contexts, identical sinks and phase boundaries, comparing widths one, two and
four. Medium is the candidate; dense-sparse and one-edit are required fallback
controls. Timing, tail, allocation and resource gates remain necessary before
adopting a performance recommendation. The previous one-edit tail review flag
remains binding. The overall Office performance goal remains active, and iWork
is outside this batch. Owned build caches and duplicate replay output are removed
at closure; the bounded records and external canonical evidence are retained.
