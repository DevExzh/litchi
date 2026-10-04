# 0905 result review

Verdict: pass with scope caveats. This is an independent offline review of the retained raw native, allocator, and micro evidence. I did not build, run workloads, or run profilers. The review is scoped to the frozen 0905 plan, the retained `4f39f5a88d` baseline, and the candidate manifest containing `crates/xml-minifier/src/audit.rs`; it makes no claim about a later merged HEAD.

The full source custody boundaries at qualification, native capture, and allocator capture each checked 10,773 files. All 8 qualification, 48 native, and 16 allocator child receipts exited zero with empty stderr. Per child, the ending receipt is a selected candidate-file check; the capture-level start and end receipts are the full-tree checks.

The native protocol has six alternating blocks, 100 measured samples and 10 warmups per child, for 4,800 samples across four cases. Even blocks run baseline then candidate in medium, dense-sparse, XLSX one-edit, and DOCX one-edit order. Odd blocks reverse both role and case order. All 48 operation sample orders are valid permutations, and the child intervals do not overlap. The allocator protocol has two repeats, three samples and two warmups per child, for 48 samples; each repeat runs all four baseline cases followed by all four candidate cases. Its 16 sample orders and child intervals also validate.

Semantic identity uses the validator's normalization: corpus, sink, output digest, and non-timing source constants are retained; keys ending in `_ns` and keys containing `allocation` are omitted, and constant vectors are collapsed. This gives zero identity differences across 4 qualification pairs, 24 native pairs, and 8 allocator pairs. It does not claim raw report byte identity. The XLSX cases retain exact corpus, output, source, cache, and sink counters. The DOCX control retains corpus and sink identity only because the validator intentionally does not fabricate XLSX source or output fields.

The primary native gate passes both cases. The gate requires a paired p50 median ratio at most 0.98 and a 95% bootstrap upper endpoint below 1.0, using seed 905009 and 10,000 draws.

| case | paired p50 estimate | 95% interval | gate |
| --- | ---: | ---: | --- |
| primary-percent-medium | 0.954388178045839 | [0.9508342443023583, 0.9584711178873901] | pass |
| primary-percent-dense-sparse | 0.947755376749029 | [0.9350215878157563, 0.9742145401706614] | pass |

The primary maximum estimate is 0.954388178045839 and the maximum upper endpoint is 0.9742145401706614. The control estimates are 0.9795903412367355 for XLSX and 0.9700651277259558 for DOCX; controls are diagnostic and are excluded from the primary adoption decision. Independent native review flags are empty.

The allocator evidence is from the separate `system_allocator_operation_scoped` binary using `CountingSystemAllocator(std::alloc::System)`, scope `operation_global_system_allocator`, and counter revision `serialized_region_peak_v3`. Across the three measured samples per child, allocation, deallocation, reallocation, failed-allocation, and byte counters are stable; failed allocation count is zero. Candidate minus baseline publication deltas are:

| case | allocation/deallocation calls | allocated/deallocated bytes |
| --- | ---: | ---: |
| primary-percent-medium | -18,832 | -1,205,248 |
| primary-percent-dense-sparse | -35,950 | -2,300,800 |
| control-xlsx-one-edit-medium | -2,354 | -150,656 |

Other phase counter deltas are zero. DOCX allocator metrics are not applicable. Live and peak fields show a three-byte snapshot difference whose cause is not established; they are excluded from stable counter deltas. These separate measurements are diagnostic and do not establish causation for the timing result.

RSS comes from `/usr/bin/time` `max_rss_kib`. It is whole-process lifetime high-water data affected by executable layout, so it is descriptive rather than an operation or phase peak. Native baseline/candidate medians are 75,576/68,836 KiB for medium, 93,044/85,992 KiB for dense-sparse, 75,820/69,278 KiB for XLSX control, and 26,760/19,804 KiB for DOCX control. The allocator medians are 72,568/67,856, 89,470/85,048, 73,258/70,626, and 24,444/19,694 KiB in the same order.

The raw micro review independently recomputes `micro.py` from all 12 TSV children (six role pairs): 36 jobs, 30 samples each, 1,000 complete slice audits per sample, and 12,960 samples. All 12 child TSVs contain 1,080 data rows with identical input/outcome comments and the expected even/odd schedules. The recomputation matches `micro-analysis.json` exactly. Eight cases have a p50 median above 1.0, none exceed 1.05, and no case meets the hard adverse rule (median above 1.05 and lower bootstrap endpoint above 1.0). Adoption remains allowed. The only review flags are p99 tails for `cardinality-2/compact` (paired p99 median ratio 1.0517844733869186) and `cardinality-2/authored` (paired p99 median ratio 1.0547573867938778); no spread flags occur. The micro timer excludes output formatting and the last-outcome check.

Qualification ran from 05:57:37.687096 to 05:57:43.899153 UTC, native children from 05:58:04.547981 to 06:01:52.158852, and allocator children from 06:01:53.328255 to 06:02:09.315730. The allocator supervisor/preflight started at 06:01:51.662140, before the final native child finished. No allocator workload child started before that native child finished, so workload child intervals did not overlap. This timing does not prove host exclusivity or a dedicated machine.

Evidence hashes and the detailed counters are recorded in [result-review.json](result-review.json). The review used the retained `analysis.json`, `micro-analysis.json`, `native/capture.json`, `allocator/capture.json`, and `diagnostics.json`; root separately owns replay and cleanup disposition.
