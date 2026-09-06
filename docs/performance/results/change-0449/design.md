# 0449: Attribute retained PPTX CPU and read evidence

Previous goal turn: progress, committed 0448. Current worktree initially contains
only user-owned GOAL.md; no owned workload/build/profile process is live. Accepted
ADR tree is unchanged from the earlier complete read. No production or harness
Rust is edited in this batch.

Reanalyse both retained 0448 media-rich profiles and every formal source/destination
phase counter. Partition SHA-256 leaf samples by explicit caller frames, retain
unclassified and missing callchains, and conserve sample count/period. Do not
confuse lifecycle-frame samples with timer-only samples. Keep reports, stacks,
capture identities and exact source copies independently replayable.

Review the hypothesis that publication byte totals imply repeated cold source
materialization. Distinguish cached logical reads, fresh source compressed capture
and destination passthrough. Map static callsites separately from counter proof;
the existing reports contain no per-range offsets or per-member read journal.
This batch makes no new latency, allocation, native, cold-I/O or scaling claim.
