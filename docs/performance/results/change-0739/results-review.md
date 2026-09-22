# 0739 PPTX lifecycle results review

status: bounded read-only review; accepted for the descriptive baseline

production_change: none

I compared the final report with the frozen raw-derived
[`analysis.json`](analysis.json), the independent
[`audit.json`](audit.json), and the timer/oracle boundaries in
[`source-review.md`](source-review.md). I found no numerical or interpretive
blocker. This review does not turn the baseline into a candidate decision or
an Office-compatibility result.

The reviewed artifact identities are:

| artifact | SHA-256 |
| --- | --- |
| [`0739-pptx-cross-copy-current-baseline.md`](../../0739-pptx-cross-copy-current-baseline.md) | `c3e7a819e29692aac9b078313a86c3c15d605b3294996b84fb5c2d0489de240c` |
| [`analysis.json`](analysis.json) | `0f78a6e966f7fc4e007f64f28242edf3ff84c763c10534e469f048c743b9ba7c` |
| [`audit.json`](audit.json) | `b58315d675df2f834be067dccf31e1a252774ab3c618784feb03371b86fad22a` |
| [`source-review.md`](source-review.md) | `ae5fc593bb0d3ed37ed3cb3c952e76f16d20724375029fbdcd27cbd9755f97cc` |

## Arithmetic and interpretation check

The packet has 18 native processes and 540 retained native samples, plus six
one-sample allocator processes. The report's central values and bootstrap
intervals round from the analysis groups as follows:

| case | lifecycle p50 | plan p50 | commit p50 | publication p50 | unassigned p50 |
| --- | ---: | ---: | ---: | ---: | ---: |
| plain | 8.3342 ms | 3.2895 ms | 3.8020 ms | 0.0008 ms | 1.2616 ms |
| media-rich | 403.9453 ms | 286.9387 ms | 88.8279 ms | 5.9763 ms | 22.0252 ms |

The displayed phase shares match the raw-derived group summaries: plain
plan/commit are 39.44%/45.44%, and media-rich plan/commit are 71.06%/22.00%.
The publication p50 process range is approximately 1.50–6.29 ms for the
media-rich case, which agrees with the report's corrected range. The report's
p95 and p99/max values also match the analysis after millisecond rounding;
with 30 samples per process, p99 is the nearest-rank maximum.

The 14 listed spread flags are the complete `analysis.json` flag set: seven
plain flags and seven media-rich flags. They include the large publication
spread in the media-rich case and the smaller plain plan/publication flags.
`order_flags` is empty for both cases. These are correctly presented as
diagnostic stability observations. They do not support a mechanism claim,
timing exclusion, or candidate decision.

The allocation table matches all six raw allocation observations exactly. The
plain process values are identical across its three runs. Media-rich allocation/deallocation counts and bytes vary slightly across
runs; live and region-peak values are stable.
The table correctly labels the values as operation-scoped global-system-
allocator evidence and does not use them as native latency.

The report correctly treats phase shares as medians of within-sample ratios
followed by process summaries. They are descriptive fractions of the chosen
clock, not CPU fractions or removable-cost estimates. `unassigned` is the
outer lifecycle residual and remains a mixture of ingress, opened snapshots,
and overhead; it is not an independently measured open phase.

## Validation and independence

`analysis.json` and `audit.json` both report `passed`; the audit records 18
native processes, 540 native samples, six allocation processes, 10,000
bootstrap resamples, and seed 7339. The independent replay agrees with the
analysis within its declared floating-point tolerance. The audit shares the
packet's invocation/report validation, so it independently recomputes raw
statistics and group arithmetic but is not a second native execution or a
second PPTX implementation.

The report's quality language agrees with the source review. Corpus creation,
the first plan, expected-output generation, and refusal controls are outside
the measured loop. Each measured loop repeats the selected logical output and
source-immutability checks after the clock. Production graph validation,
candidate construction, revision checks, plan application, and final stream
publication remain inside the lifecycle clock. Reopen and post-operation
oracle work remain outside it. The owned oracle uses the same Litchi stack and
does not establish native Office, real-producer, rendering, or raw ZIP
metadata compatibility.

## Remaining risks and disposition

The media-rich publication phase has a large across-process spread while the
outer lifecycle is comparatively stable. No mechanism is measured for that
spread, so it should not guide an internal optimization. The allocator lane
has one sample and zero warmups per process, no phase allocation split, and a
separate instrumented binary; it is suitable for retained descriptive counts,
not latency comparison or causal attribution. Region peaks are not RSS or a
full-process heap measure.

The source review identifies the next useful seam: split plan work into
closure/topology proof, candidate graph construction, serialization/deflate,
reopen/capture, and patch capture; split apply work into fingerprints,
snapshot capture, fresh prepare or retained-archive reuse, validation, and
assignment. That is a future diagnostic measurement requiring requalification,
not a production change implied by this packet.

The current packet remains limited to the generated plain and media-rich
corpora on one serial host and warm process lifecycle. Physical cold-cache,
remote/range, concurrent scaling, producer variation, retention lifetime,
RSS, and comprehensive CRUD completion remain unproved. No historical
cross-build timing comparison or candidate retention decision is supported.
