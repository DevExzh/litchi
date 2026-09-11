# Fresh control profile scope

Control binary is byte-identical to the 0508 final binary. Heaptrack runs the
whole child with 20 large ODT exports and no warmups; generation, preflight,
opening and JSON output are included in whole-child totals. The stack through
`append_sink_precharged` accounts for exactly 200,000 allocation calls, or
10,000 per export. This is the admission evidence for bounded buffer reuse.

Callgrind collects only within matched `*write_text_blocks_to_writer*` calls,
including their callees, for five large exports and no warmups. Collected
instruction references total 304,455,496; all are inclusive under the parser.
SHA-256 software compression in the checking sink accounts for 43.26% exclusive,
and `normalized_xml10_decoded_len` 9.31%. The latter is a separate future
candidate; this batch does not change decoding or sink write boundaries.
Callgrind references are simulated executed instructions, not hardware cycles;
its emulated CPU dispatch and instrumented timings do not describe native SHA
cost or native elapsed-time shares. Raw profiles and annotations are retained.

The fresh perf hardware probe failed with permission denial. No hardware
counter, cache, branch, memory bandwidth, or operation-local RSS claim follows.
Heaptrack RSS includes profiler overhead; uninstrumented child RSS is measured
separately and includes setup/open/oracle activity. Heaptrack's 544 leaked bytes
and one allocation are whole-child reports, not an attribution to ODT.
