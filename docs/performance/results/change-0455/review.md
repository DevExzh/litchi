# Review of the 64 KiB transfer experiment

The production experiment changes only the private copy buffer size in the ZIP
preservation owner. The baseline and candidate share the new short-read/partial
sink regression and the sink-summary journal field. A separate ZIP64 test-only
upper bound changes from 32 KiB to 64 KiB alongside the candidate.

The owner prepares the plan and validates layout before output. Its exact-read
helper loops over short reads; `write_all_counted` retries short writes and
charges raw bytes only when accepted. The focused regression uses a stored
member larger than 192 KiB, 4 KiB short reads, and a sink failure three bytes
past one full copy chunk. It asserts exact complete output and exact partial
accepted-byte accounting. Existing ZIP tests cover Deflate, descriptors,
ZIP64, unknown metadata and no-op preservation; OPC/PPTX tests cover source
identity, cancellation and bounded output.

The stack buffer is fixed and adds 32 KiB per active publication. It is not
heap storage and is not represented by allocator counters. There is no new
allocation, worker, executor, provider or public API. Concurrent callers
multiply that stack footprint. The ODF writer and its memory formula are
unchanged.

Cancellation/source checks occur at existing owner boundaries, with larger
transfer chunks between checks. The output-budget wrapper admits a requested
write before calling the sink: insufficient remaining budget may therefore
refuse a 64 KiB request earlier than the 32 KiB version. The operation still
returns a typed failure and reports accepted partial output; it does not
promise an identical failure byte offset across buffer sizes.

The PPTX journal exposes the existing bounded sink summary only after the
publication timer and allocation region close. It does not expose ZIP raw-copy
accounting. Direct ZIP tests validate that counter; sink accepted bytes and
provider read bytes are separately named observations, not substitutes for it.
Retain the candidate after 720 formal samples, 240 ordinary confirmation samples,
and 240 fixed allocator-policy diagnostic samples. The range request and write
reductions repeat with unchanged bytes, hashes, managed budgets and operation
allocation counts. Range API medians improve 4.775%/4.449%.

The original perf stat result remains adverse (+9.25% instructions, +12.00%
cycles). Slow/high-page-fault processes also occur in the baseline. Fixed mmap
thresholds reproduce slow and fast regimes in both builds without a material
candidate penalty; local instruction profiles also have similar totals.
This supports allocator paging as a major source of process variability, but
exact historical allocation decisions were not traced. No CPU improvement or
large bytes-only speedup is claimed. All original flags, the failed remote-symbol
profile attempt and the separate successful local-symbol retry remain visible.
See the [diagnostic summary](formal/diagnostic-summary.md) and
[full change review](../../changes/0455-zip-preservation-transfer-chunks.md).

Release checks pass 2,188 tests, with 6 ignored, and 1,000 sanitizer fuzz runs.
The pinned native QA self-pair keeps its prior exact output through both
providers. Broader native pairs, application roundtrips and scaling remain open.

An independent read-only review of the preservation path and diagnostic evidence
agrees with retaining the range-request change under these limits. The review
notes that the fast local profiles do not attribute the slow regime, and that
other allocators and physical/cold I/O remain unmeasured.
