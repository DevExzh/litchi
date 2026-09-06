# 0430: PPTX publication CPU attribution

The previous media-rich provider profile left most Deflate samples without
their lifecycle callers. New frame-pointer captures use the exact same release
executable and recover those callers, separating measured publication from
corpus construction and untimed checks. The
[portable bundle](../results/change-0430/README.md) retains two full-process
recordings, each with three warmups and 100 retained lifecycle samples.

| Unweighted cycle samples | Bytes | Warm file |
|---|---:|---:|
| Full process | 16,674 | 16,673 |
| Explicit lifecycle iteration ancestry | 14,050 | 14,046 |
| Iteration with publication ancestry | 12,520 | 12,541 |
| Iteration with `deflate_medium` and CountingSink ancestry | 11,692 | 11,712 |
| Explicit corpus-builder ancestry outside iteration | 2,386 | 2,370 |
| Other or unresolved | 238 | 257 |

The Deflate intersection is 83.22% / 83.38% of iteration samples. Iterations
include provider construction, observers, checks, drops and warmups; these are
sampled CPU fractions, not API wall-clock fractions. Unknown samples and
decoder diagnostics are retained. The two providers are descriptive checks of
the same path, not a controlled provider-cost comparison.

Decoded event totals equal the counts printed by `perf record`. The script
logs retain six/five addr2line diagnostic lines; five/four sample blocks have
no decoded frames. Neither count is silently discarded or interpreted as an
API sample. The exact compressor marker avoids classifying Deflate decoding
as compression.

Source review traces copied image/chart payloads through logical shared buffers
into OPC topology additions, then ZIP's Deflate generator during publication.
Untouched destination entries already use raw preservation. This makes safe
compressed source-part transfer the first optimization candidate for this
measured workload. ZIP must retain physical framing ownership and OPC must
issue source-bound transfer authority; existing validation, source checks,
limits and cancellation must survive. The
[design audit](../results/change-0430/transfer-design.md) records the proposed
implementation boundary and tests.

No production or harness Rust code changes in this batch. No speedup, memory
reduction, physical-copy reduction, native support, or scaling result is
claimed. Captures remain synthetic and files warm. A matched before/after
experiment is still required to measure the candidate, including any staging
memory or range-read cost. The global non-iWork goal remains incomplete.
