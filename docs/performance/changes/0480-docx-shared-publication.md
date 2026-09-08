# Change 0480: share DOCX target XML during publication

`performance_claim: none; scoped control/candidate ownership comparison`

`claim_authorized: false`

Publishing an existing-document paragraph-copy patch previously copied the
complete target XML into a fresh Vec and Arc for OPC. The publisher now clones
the existing immutable Arc handle and uses the existing shared overlay entry
point. That removes one payload copy and its separate Vec/Arc allocation while
retaining the same validation, source checks, limits, serializers, exact
untouched member preservation and inverse/partial-sink behavior. Production
changes are confined to two lines in the DOCX format owner.

The [evidence bundle](../results/change-0480/README.md) records fresh control
and candidate builds, source manifests, deterministic corpus identities,
A1/B1/B2/A2 process order and separate normal/allocator total/phase observations.
The scenario copies one source paragraph to the tail of an existing plain DOCX
with 64, 8,192 or 131,072 paragraphs. It has one operation per edit and refuses
section properties. The 48 reports contain 1,440 samples; pilots and separate
whole-process perf-stat runs are excluded from those totals.

## ADR and correctness

The shared bytes remain private immutable snapshot state under ADR 0003.
ADR 0005's explicit positional source, sequential sink and limits remain in
place. The existing OPC generic single-overlay path performs identical source,
no-op, signature, XML, limit and publication checks under ADR 0006. Ownership
stays with DOCX semantics and the OPC physical owner under ADRs 0010/0011/0024.
No production API, dependency, unsafe code or ambient provider is added.

The existing paragraph-copy suite covers byte-exact untouched publication,
canonical durable forward/inverse replay, stale source and source version,
exact no-op, limits/refusals, seven-byte partial writes, write-zero, and failure
after 128 accepted bytes. The same measured harness checks full source and
candidate semantics, independent XML, raw untouched members and exact inverses.
No test that merely mirrors the two-line implementation was added.

## Scope

This removes one temporary full-XML owner from a materialized transaction.
It does not turn snapshots/patches into a bounded-window append route. Complete
source/candidate XML and paragraph layouts remain; the full non-iWork goal,
including broader source/native/CRUD and scaling evidence, remains open.

## Matched normal results

Times are milliseconds for one complete lifecycle, with 30 samples per process
after three warmups. RSS is whole-process maximum KiB. Each row compares the
same size/repeat; no normal and allocator timings are pooled.

| Paragraphs | Repeat | Arm | Mean ms (95% t interval) | p50 ms | p95 ms | p99 ms | RSS KiB |
| ---: | ---: | --- | --- | ---: | ---: | ---: | ---: |
| 64 | 1 | control | 0.168799 (0.167171–0.170426) | 0.167000 | 0.177131 | 0.182921 | 4,744 |
| 64 | 1 | candidate | 0.168942 (0.167661–0.170223) | 0.167671 | 0.175420 | 0.180341 | 4,824 |
| 64 | 2 | control | 0.168659 (0.166589–0.170729) | 0.166301 | 0.179751 | 0.188360 | 4,744 |
| 64 | 2 | candidate | 0.170074 (0.168027–0.172121) | 0.168110 | 0.183391 | 0.191751 | 4,828 |
| 8,192 | 1 | control | 16.555023 (16.539328–16.570717) | 16.543697 | 16.632267 | 16.632917 | 11,184 |
| 8,192 | 1 | candidate | 16.570107 (16.550109–16.590106) | 16.564220 | 16.662310 | 16.698900 | 10,728 |
| 8,192 | 2 | control | 16.873467 (16.816410–16.930523) | 16.857759 | 17.187361 | 17.237431 | 10,924 |
| 8,192 | 2 | candidate | 16.903092 (16.860787–16.945397) | 16.889139 | 17.097400 | 17.153589 | 10,732 |
| 131,072 | 1 | control | 270.012341 (269.037428–270.987253) | 269.310218 | 274.485198 | 274.662879 | 113,388 |
| 131,072 | 1 | candidate | 269.373280 (268.965423–269.781138) | 269.179225 | 271.426324 | 273.144572 | 107,200 |
| 131,072 | 2 | control | 272.926225 (271.923455–273.928996) | 272.548574 | 277.595654 | 278.069437 | 113,648 |
| 131,072 | 2 | candidate | 265.883534 (264.968038–266.799031) | 265.787657 | 271.286070 | 271.389856 | 107,200 |

Normal total mean changes are +0.085%/+0.839% at 64 paragraphs,
+0.091%/+0.176% at 8,192 and −0.237%/−2.580% at 131,072. The evidence supports
retaining an ownership improvement; it does not establish a general latency
speedup. Every individual timing/RSS review flag is retained in the analyzer
and independent review, including short phase measurements.

## Allocator and I/O result

| Paragraphs | Control peak bytes | Candidate peak bytes | Bytes removed | Control allocation calls | Candidate allocation calls |
| ---: | ---: | ---: | ---: | ---: | ---: |
| 64 | 509,974 | 506,598 | 3,376 | 1,888 | 1,886 |
| 8,192 | 3,071,310 | 2,669,662 | 401,648 | 213,346 | 213,344 |
| 131,072 | 41,793,870 | 35,371,102 | 6,422,768 | 3,653,986 | 3,653,984 |

These are incremental operation peaks, after subtracting entry live bytes.
At each size, the reduction equals the exact candidate XML length plus 40 bytes
for the separately allocated Arc/Vec header on this target. Requested bytes
fall by that same amount, allocation calls fall by two, and reallocations are
unchanged. The target allocation stays owned by the returned publication; only
the duplicate overlay owner is removed. Both total allocator lifecycles return
to entry live bytes with no failed allocations.

The large peak reduction is about 15.37%. Normal process maximum RSS falls
from 113,388/113,648 to 107,200/107,200 KiB, approximately 5.46%/5.67%, within
the recorded broader process scope. That includes changed oracle/setup
allocation behavior and is not a direct substitute for the operation peak.

Logical source calls/bytes/request histograms, sink accepted bytes/write
histograms and output hashes remain identical between arms. The source is an
immutable caller-owned Arc adapter and the sink accepts at most 16 KiB per
write without retaining the archive. These are logical memory I/O observations;
no disk or remote-source speedup follows. The explicit removed XML clone is a
known copy site, while total physical memory-copy traffic remains unmeasured.
The unchanged large reallocation-request total is explained by 0479's parser
range-growth audit, not by physical-copy bandwidth.

## Whole-process hardware counters
These are separate normal processes with full corpus/oracle, warmup, lifecycle
and report work. They are contextual counter observations, with no isolated
operation PMU or general CPU speedup claim. All six event running fractions
are 100%. Generic cache misses are not an exact L1/LLC breakdown.
| Event | Control | Candidate |
| --- | ---: | ---: |
| cycles | 45,495,146,335 | 44,721,467,111 |
| instructions | 191,274,111,072 | 190,388,694,874 |
| branches | 38,186,718,211 | 38,035,131,100 |
| branch-misses | 116,505,098 | 107,606,006 |
| cache-misses | 12,551,761 | 8,196,329 |
| page-faults | 210,020 | 42,433 |

The prior 0479 sampled callstack profile remains the source of scanner
attribution. No new sampled CPU or syscall profile is claimed in 0480.

## Review and validation

The source and independent measurement reviews recommend keeping this change
for its exact allocation reduction. The production commit is `b3a9c499d`.
All positive pair flags above 5% occur in individual phase statistics; no
total-lifecycle latency or whole-process RSS pair has a positive flag above
5%. The independent review retains every pair and repeat flag, including the
small candidate normal total p99 repeat drift of +6.33%. Short phase ratios
and process-order drift prevent a phase speedup claim.

The scoped Rust checks pass: 926 DOCX library tests, 12 paragraph-copy tests,
six paragraph-removal tests, nine OPC shared-overlay tests and seven harness
tests; DOCX formatting, all-feature Clippy with warnings denied, rustdoc with
warnings denied, crate boundaries and the strict performance registry check.
Two earlier failed attempts remain in the evidence: workspace formatting found
an existing Keynote formatting difference, and default-feature Clippy found
six unfulfilled lint expectations around feature-gated modules. The scoped
formatting and all-feature Clippy gates pass without modifying those sources
or suppressing warnings.

Evidence review also corrected an inherited summary label: the absolute
operation-region allocator peak is distinct from the global allocator
high-water mark. The incremental peak comparison above is unchanged. The
initial generated summary and successful development evidence-test/analysis
receipts are retained; final gates cover the corrected portable verifier and
summary. Historic 0479 baseline timestamps and raw captures remain intact.

All 22 required final gates pass. The ledger retains 28 attempts: 26 successful
and the two disclosed development failures. The final 14 evidence tests,
data-only verification and full sealed verification pass. Final source and
embedded-input checks authenticated all four binaries before their removal;
shared Cargo caches and both user-owned untracked documents were preserved.

A fresh copied baseline also passes without the original runtime binaries.
All eight independently resealed corruption probes are rejected, including
summary/PMU arithmetic and validation-gate omissions/exclusions. The sealed
bundle retains the probe outputs and their input seal.

The materialized XML and paragraph index still grow with the document. The
next work is scanner/layout and patch ownership analysis followed by an
explicit bounded-window append capability; the full non-iWork goal remains
open.
