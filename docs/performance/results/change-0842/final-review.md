# 0842 final review

Adopt the feature configuration for the measured larger generic OPC workloads.
The original p50/RSS and block-spread guards have zero flags. Both original tiny
p99 observations remain in analysis.json and decision.json. The separately
preregistered 1,000-sample confirmation has no p50/p95/p99/RSS guard flags; it
does not establish a universal tail guarantee. Its independent bootstrap units
remain six paired blocks. Read-tail uncertainty and the confirmation read RSS
point increase remain explicit in the report.

The independent reviewer verified the complete native-resave lock transition:
original to baseline-normalized adds same-file/winapi-util and changes exactly
four local records, with no removals or existing registry version/checksum
changes. Normalized baseline to candidate changes only flate2's crc32fast edge.
Perf-execution changes only that edge; the current five-file production diff
matches candidate.patch. All three locked checks/metadata receipts pass.
The original 506-file admission remains immutable.

Earlier review required canonical complete lock-package comparisons, three
locked tool checks bound at admission/closure, and the original manifest hashes
for the baseline rustdoc control. All were fixed and re-reviewed before the
native results were reduced. The stronger lock guard exposed the pre-existing
native-resave drift; both failed orchestration scripts and the successful
baseline control remain visible. No formal native or tail process was rerun.

The report deliberately separates the measured after-source from the final two
standalone lock repairs, and the original native result population from the
longer tail confirmation. Warning-denied whole-workspace rustdoc and all-target
checks are not claimed green: their failures reproduce with baseline dependency
declarations in user-excluded iWork. The five affected crates' applicable gates
and 6,931 tests pass.

The accepted ADR matrix is in selection.md. No public API, new unsafe code,
compression level/framing, preservation, durability or parallelism change is
introduced. Exact OPC/codec bytes and unchanged spans pass. Broader semantic
Office, other-architecture, cold-I/O, allocator and scaling claims remain open.

Final independent report/closure review passed: units, 168/4,344 original and
24/24,000 confirmation ledgers, all tail disclosures and decision fields agree
with the frozen analyses. close_v2 verification passed after cleanup. No
blocking issue remained before sealing.
