# 0483 CPU and syscall profile review

This review audits the retained normal binary profile at count 131072, with 30
measured samples after 3 warmups on CPU 2. It covers separate whole-process
runs for the materialized and bounded routes. The profile is diagnostic and
excluded from formal evidence. Each run includes corpus construction, untimed
independent output/member/semantic oracles, warmups, all 30 measured harness
lifecycles, report JSON serialization, and teardown.

The source of record is profiles/accepted/profile.json. Raw artifacts remain
under profiles/accepted/ and validation/, and the detailed machine-readable
audit is profile-review.json. The normal binary is
/home/zhuhe/.cache/litchi-goal-0483/accepted/normal/docx_bounded_tail_append_compare, 420,503,104 bytes, SHA-256
6f3a30ae7c826befb10c93787ece9d771e558848443ad0a866c5d9a12e0e629d. The source snapshot is 30d3ad60b0ae90911992b9855c5328913de2741f27090ba6c31b760c2ebcaeaf with 7,142 files.

## Validation and custody

All six workload reports passed analyze.validate_report with normal
instrumentation, count 131072, 30 samples, and 3 warmups. Each retained report
has 30 samples, so planned and actual totals are six workload runs and 180
samples.

All ten profile receipts have exit code 0, source_unchanged=true, and matching
source-before/source-after snapshots. Receipt and retained-output metadata
were recomputed against the profile manifest.

| Route | Workload report | Samples | Report SHA-256 |
| --- | --- | ---: | --- |
| materialized | profiles/accepted/materialized/perf-stat.report.json | 30 | 803737a5145ce7ee5d1f2597a20943140013d45aed5fd715618461251571433b |
| materialized | profiles/accepted/materialized/perf-record.report.json | 30 | 1e2803e73b0c49b554026f117a12e4b43928c2aa39b67ca313720c83037cc9c1 |
| materialized | profiles/accepted/materialized/strace.report.json | 30 | c194fdeba44fb959554685888844d7a5802650f21fbb374cc5309f581d1407bb |
| bounded | profiles/accepted/bounded/perf-stat.report.json | 30 | e662ad47c72ef19ae1e28c3e49e66b8d086f6c2d9f0dabcee70272a5d1cdc061 |
| bounded | profiles/accepted/bounded/perf-record.report.json | 30 | e8b2f2c203cc370f296cd5374fe485336f9c6578e42ffb8c79a4ca3d6fb63200 |
| bounded | profiles/accepted/bounded/strace.report.json | 30 | 2c950ad6e1b9b394c6b98957fa496bed2dea82a26f8fe0edaeff773caf3be1a8 |

Receipt SHA-256 values, plus stdout/stderr artifact hashes, are retained in
profile-review.json under validation.gate_receipts.

## Whole-process PMU counters

The perf stat counts are raw event counts from independent whole-process
invocations. IPC and rates are descriptive ratios computed from those counts.

| Route | Cycles | Instructions | IPC | Branches | Branch misses | Cache misses | Page faults |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| materialized | 43,831,570,742 | 186,115,375,309 | 4.246148886712 | 36,076,207,159 | 78,830,955 | 9,399,382 | 35,629 |
| bounded | 79,530,894,675 | 360,039,435,601 | 4.527038669341 | 70,103,099,689 | 123,220,522 | 4,041,293 | 35,625 |

The bounded/materialized descriptive ratios are 1.814465996282 cycles,
1.934495927611 instructions, 1.943194842519 branches, 1.563098176345 branch
misses, 0.429953054360 cache misses, and 0.999887731904 page faults. Branch
miss rates are 0.218512313815% and 0.175770433186% respectively. These are
whole-process route observations and do not identify operation-only costs.

## Perf record and weighted period reconciliation

perf record explicitly captured the cycles event with frame-pointer callchains.
The report used --no-children, and script periods were weighted by each sample
header.

| Route | Target blocks | Target cycles | perf-exec blocks | Helper cycles | All blocks | All cycles |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| materialized | 982 | 44,242,200,589 | 5 | 5,779 | 987 | 44,242,206,368 |
| bounded | 1,769 | 79,674,873,466 | 5 | 5,888 | 1,774 | 79,674,879,354 |

Each all-block period sum exactly equals the corresponding perf report event
count. Both reports say Total Lost Samples: 0. The perf stat cycles above came
from separate workload invocations, so their counts are not expected to equal
these sampling-run event totals.

The retained perf report top self/leaf rows are:

| Route | Symbol | Overhead | Raw line |
| --- | --- | ---: | ---: |
| materialized | litchi_docx::source_backed::paragraph_copy::scan_document | 15.77% | report stdout:12 |
| materialized | quick_xml NsReader process_event | 13.33% | report stdout:607 |
| materialized | xml_minifier::audit::verify_with_policy | 11.66% | report stdout:948 |
| materialized | __memcmp_evex_movbe | 11.29% | report stdout:1108 |
| materialized | quick_xml NamespaceResolver resolve_event | 9.86% | report stdout:1499 |
| bounded | xml_minifier::audit::verify_reader_with_policy<SpliceAuditReader> | 13.28% | report stdout:12 |
| bounded | __memmove_avx512_unaligned_erms | 5.71% | report stdout:544 |
| bounded | sha2::sha256::x86_sha::compress | 5.25% | report stdout:2314 |
| bounded | SpliceAuditReader::consume | 5.08% | report stdout:4917 |
| bounded | __memcmp_evex_movbe | 4.97% | report stdout:6501 |
| bounded | quick_xml NamespaceResolver resolve_event | 3.39% | report stdout:7204 |
| bounded | crc32fast::baseline::update_fast_16 | 3.16% | report stdout:7447 |

The rows are self/leaf percentages from --no-children; the ancestry shares
below are inclusive and overlap.

## Lifecycle intersections

Marker shares were recomputed from perf script blocks. A marker counts once
when a resolved frame contains it. The lifecycle intersection uses the same
sample block, separating repeated measured work from setup and oracle work.

| Route | Marker | Matching blocks | Target cycles | Target share | In lifecycle | Outside lifecycle |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| materialized | run_materialized_iteration | 807 | 36,321,107,271 | 82.096068431168% | 807 / 36,321,107,271 | 0 / 0 |
| materialized | scan_document | 691 | 31,107,116,011 | 70.310960116966% | 629 / 28,312,984,411 (63.995425259293%) | 62 / 2,794,131,600 (6.315534857673%) |
| materialized | publish_plain_paragraph | 497 | 22,365,458,059 | 50.552318287171% | 482 / 21,690,441,443 (49.026588086111%) | 15 / 675,016,616 (1.525730201060%) |
| materialized | verify_with_policy | 137 | 6,193,005,832 | 13.997960656459% | 132 / 5,938,469,449 (13.422635786513%) | 5 / 254,536,383 (0.575324869946%) |
| bounded | run_bounded_iteration | 1,593 | 71,713,643,182 | 90.007853244477% | 1,593 / 71,713,643,182 | 0 / 0 |
| bounded | scan_main_part | 611 | 27,507,931,480 | 34.525227695201% | 577 / 25,977,218,498 (32.604028557492%) | 34 / 1,530,712,982 (1.921199137710%) |
| bounded | run_replay | 633 | 28,502,624,073 | 35.773667196551% | 614 / 27,646,985,448 (34.699754446172%) | 19 / 855,638,625 (1.073912750379%) |
| bounded | verify_reader_with_policy | 1,052 | 47,356,856,029 | 59.437629416768% | 1,010 / 45,466,291,265 (57.064780007968%) | 42 / 1,890,564,764 (2.372849408800%) |
| bounded | audit_splice | 1,057 | 47,581,940,991 | 59.720133739911% | 1,014 / 45,646,336,782 (57.290755286206%) | 43 / 1,935,604,209 (2.429378453705%) |

The materialized scanner has 629 matching lifecycle blocks and 62 outside
them; bounded scanning and replay likewise have explicit outside blocks. This
is why these profiles cannot be read as transaction-only timing.

The first callchain annotations under the dominant rows include materialized
branch<quick_xml::events::Event,...> at 8.65% and with_xml at 5.09%
(report stdout:14 and :19), and bounded branch<(u64,...), OpcError> at 7.06%,
with_verified_decoded_reader at 6.89%, and measure_callback at 3.73%
(report stdout:21, :27, :32). These are ancestry annotations, not independent
costs.

## Syscalls

strace -c includes process startup, corpus/oracle work, measured lifecycles,
report writing and teardown. The largest rows are:

| Route | Syscall | Calls | Seconds | Share |
| --- | --- | ---: | ---: | ---: |
| materialized | write | 24,476 | 0.026357 | 91.07% |
| materialized | brk | 84 | 0.001078 | 3.72% |
| materialized | munmap | 8 | 0.000811 | 2.80% |
| materialized | read | 1,193 | 0.000462 | 1.60% |
| materialized | close | 203 | 0.000085 | 0.29% |
| bounded | write | 24,752 | 0.027014 | 90.84% |
| bounded | brk | 84 | 0.001768 | 5.95% |
| bounded | munmap | 8 | 0.000657 | 2.21% |
| bounded | read | 1,193 | 0.000180 | 0.61% |
| bounded | openat | 203 | 0.000050 | 0.17% |

Totals are 26,473 calls / 0.028943 seconds / 1 error for materialized and
26,749 calls / 0.029738 seconds / 1 error for bounded. The sole error row is
access (one call) in both reports. The 276-call delta is entirely in the
write row, but these whole-process writes do not establish a sink syscall
bottleneck.

## Selected next measurement

The measured next work is phase attribution for repeated XML scanning/layout
and bounded replay/validation paths. The strongest retained lifecycle
intersections are materialized scan_document (629 blocks,
28,312,984,411 cycles, 63.995425259293% of target), bounded
verify_reader_with_policy (1,010 blocks, 45,466,291,265 cycles,
57.064780007968%), and bounded run_replay (614 blocks,
27,646,985,448 cycles, 34.699754446172%). The next experiment should preserve
the accepted normal binary and source binding, route/count/sample shape, and
all XML/member/semantic and unchanged-member proofs while separating setup and
oracle ancestry with matching sample blocks.

These profiles select what to measure next. They do not approve an
optimization, establish a speedup, or provide operation-only PMU/syscall
counters.

## Limits

- perf stat and perf record are separate whole-process invocations.
- perf report --no-children rows are self/leaf overhead; inclusive marker shares
  overlap.
- Kernel address maps were restricted, so some kernel frames remain unknown;
  recorded samples were not lost.
- strace -c includes report JSON writing and cannot establish an in-memory sink
  syscall claim.
- The profile covers the normal binary at count 131072 and is excluded from
  formal timing and allocator evidence.
