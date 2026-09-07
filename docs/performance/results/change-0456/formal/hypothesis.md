# Direct shared payload framing experiment

The ZIP preservation writer currently prepares each regenerated member as a
complete one-entry archive Vec, even when its compressed bytes are already
verified and shared. The measured media publication has 21,057,867 allocated
bytes and 17,406,388 regional peak above entry (0455). The 16 MiB payload-sized
preparation buffer is a concrete candidate for removing another proportional
allocation and copy before sequential output.

Compare current code with a small framing/payload separation for verified
precompressed entries and, only if the same proved grammar permits it, shared
stored entries. Keep full layout preflight, exact output, raw preservation,
ZIP64, partial-write accounting, source limits and caller budgets. Do not lower
managed admission until its complete model is proved. No new executor or I/O.

Use the same 24-lane ABBA ordinary/allocator protocol as 0455. Captures use
CPU 2, one worker, 30 samples after 3 warmups. Preserve all >5% timing/RSS flags;
ordinary and instrumented populations remain separate. The allocator paging
variability observed in 0455 is a known limitation, not permission to discard
adverse observations. The intended improvement is proportional publication
memory/copy removal, with unchanged byte/hash/work output, not a predeclared CPU
or latency gain. Confirm exact ZIP/PPTX output and run relevant release/fuzz gates.
