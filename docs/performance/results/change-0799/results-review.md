# 0799 final results review

This is a bounded offline review of the frozen 0799 packet. I reconciled the
main report with `analysis.json`, `decision.json`, `root-native-audit.json`,
`root-cg-totals.json`, the retained raw reports, and the self-check. I did not
run a build, test, capture, profiler, or replay command.

## Evidence reconciliation

The packet has 39 frozen cases and 1,248 report processes: 936 native children
and 312 Callgrind children. The native lane has 30 samples per child, so it
contains 28,080 native timing samples. Each Callgrind child has one measured
profile sample; including those 312 records gives the audit's 28,392 total
sample records. This is a combined report count, not a claim of 28,392 native
timing samples.

The six native block pairs, nearest-rank p50, paired median, and bootstrap
endpoints reproduce the 78 rows in `root-native-audit.json`. The selected table
in the main report is a presentation subset; all 78 case/mode rows, including
the non-protected tail and zero/one/33-attribute rows, remain in the summary
and analysis. The independent audit reports 1,248 reports, 28,392 sample
records, and `production_adoption: false`, and its rows agree with the full
analysis.

The semantic evidence is complete for the captured reports. All 39 self-check
cases pass baseline and candidate differential checks against quick-xml, with
clone checks at advances 0, 1, 2, 3, 4, 5, 32, and 33. Every native and profile
report carries the same semantic oracle parity, expected checksum/result, and
120-byte baseline and candidate iterator sizes. The independent native audit
recomputes the literal case bytes, expected error/checksum markers, every
sample result, and all 78 ratios. The retained setup failures and the rejected
40-case catalog are accounted for before capture; no captured case was retried
or omitted.

The raw Callgrind audit contains 624 dumps: 312 positive and 312 termination
dumps. For each dump it conserves `Ir`, `Bc`, `Bcm`, `Bi`, and `Bim` from parsed
self rows through the summary and totals records, with zero termination
counters. Exact owner qualification succeeds for all 312 positive profiles.
The selected checked-iterator, `IterState`, drop, allocator-name, and lexical
matching rows are a name-based census. Inclusive descendants overlap, so these
rows support mechanism diagnostics only; they do not count allocator API calls,
native cycles, or phase fractions.

## Decision and interpretation

The frozen decision is correctly fail-closed for workflow trials. The two
benefit rows pass: `distinct-1/consume` is 0.651018 (34.898% improvement,
CI high 0.655196) and `distinct-2/consume` is 0.832837 (16.716% improvement,
CI high 0.838709). Both satisfy the ratio-at-most-0.97 and CI-high-below-one
requirements.

Six protected consume rows veto advancement: the valid, quoted-long, and
unterminated-long duplicates after two items, plus the three syntax tails after
two items. Their ratios are 1.503305–1.559910 and their CI lows are
1.488662–1.529931, so each exceeds both protected thresholds. The decision
also retains all 19 diagnostic consume regressions and all 23 process-p50
spread groups over 5%; none is silently converted into a workflow result.
Construction has no diagnostic regression flag. Empty input remains a visible
non-benefit because its interval crosses one.

The main report's replay-boundary explanation is consistent with the raw
counter rows: the one-item long duplicate forms have identical candidate `Ir`,
while the first post-prefix duplicate and syntax cases pay the additional
checked-path work. That is evidence for a shifted replay cost, not proof of a
native cycle fraction or an allocator count. The proposed next design should
therefore be treated as a hypothesis to remove or avoid that state rebuild;
moving the replay later would preserve the semantic boundary but would not by
itself remove the measured work.

The direct helper captures use hot, repeatedly reused inputs, selected process
timing, and opaque construction/destruction. They provide no public-workflow,
producer/CRUD, cold-input, concurrency, cross-format, or end-to-end resource
result. The 0798 census motivates the two target classes but supplies no speed
model. Accordingly, `advance_to_workflow_trials` and production adoption both
remain false; the candidate is archived and production is unchanged. Any
replacement needs fresh semantic, direct, public-workflow, resource, and
cross-format gates.

Cleanup and custody are consistent with the final report: the owned target was
removed while the final and failed binary identities, raw captures, audits,
source/gate records, and failure archives remain retained. I found no results
blocker.
