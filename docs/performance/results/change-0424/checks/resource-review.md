# 0424 resource-review disposition

Static review of `resource-review.md` against the frozen
`matched/summary.json`, `matched/result-table.md`, and `profile-notes.md`.
No Cargo, workload, or profiler command was run for this review.

The disposition is supported by the retained evidence. The summary has all 16
passing rows with identical corpus/output/sink identities. The allocator lane
has 30 samples for each role/repeat. For plain, requested bytes and allocation
calls are 8,916,495 and 10,993 for both roles; mean region peaks are 1,059,128
bytes for control and 1,059,132 for candidate. For media-rich, the corresponding
values are 117,608,609 versus 100,831,137 requested bytes, 13,314 versus
13,306 calls, and 229,300,275 versus 212,522,935 mean region peak bytes. These
are the stated −14.266% and −7.317% deltas; the exact differences are
16,777,472 and 16,777,340 bytes.

The timing and RSS statements match the result table and summary review
triggers. Plain normal p50 deltas are +0.805% and +0.755%; media normal p50
deltas are −0.013% and −3.825%. All repeat-drift fields remain within the
declared 5%/5%/10%/15% ceilings, while allocator elapsed comparisons remain
withheld. The largest adverse whole-process RSS delta is plain allocator R1 at
approximately +0.107%; the media values are approximately +0.015% and
+0.033%. The −100% media R1 `rss_delta_bytes` entry has an endpoint zero and is
properly called out as undefined process-memory evidence.

The read counters are equal across every role and repeat: plain source and
destination are 123/14,105 bytes and 423/58,432 bytes; media source and
destination are 1,794/50,368,324 bytes and 975/16,845,847 bytes. The profile
manifests independently retain the stated media clone ancestry: control has
184,610,503 requested bytes over 99 `clone_bytes_checked` events, including
33,565,546 lifecycle bytes; candidate has 117,501,639 over 67 events,
including 16,788,330 lifecycle bytes. The profile notes correctly keep these
whole-command Heaptrack totals separate from the operation allocator vectors.

No material overclaim or scope mismatch was found. The recommendation to
retain the candidate resource change is bounded to the media source-backed
allocator diagnostics. The document correctly withholds latency, physical-I/O,
managed-budget, post-drop retention, concurrency, and native-producer claims;
the small plain timing adverse values remain visible.
