# 0446: Consume owned content-type Part names

Previous turn: progress, committed 0445. Matched plain-source profiling places
ContentTypeMap::from_xml in 29.291% of inclusive run-frame sampled period. Source
inspection finds an owned Part-name String passed by reference to PackURI::new,
whose Into<String> parameter clones it. The original String is not used afterward.

Hypothesis: move the String into PackURI without changing parsing, validation,
error priority, limits, case folding, map ownership, freshness or publication
readback. Three content-type parses per ordinary Part-addition lifecycle suggest
about 3*N+1 fewer allocation calls. Do not cache or skip a required parse.

Freeze before editing production code or retained measurements. Capture the same
plain-source selector and fixtures using before/after normal and allocator builds:
A1 before R1, B1 after R1, B2 after R2, A2 before R2, three sizes 64/1024/4096,
30 samples/three warmups, CPU 2, one worker. R1 runs normal then allocator, tiny to
large; R2 reverses modes and sizes. Total: 24 reports/720 samples, plus before/after
stat and record profiles. All attempts are retained; root CPU jobs are serialized.
Restore only the owned candidate file for before replay, after the prior CPU handle
is terminal, then restore the candidate before final checks/commit.

Practical acceptance: at least 5% fewer operation allocation calls at medium and
large in both repeats, complete correctness/preservation gates, and explicit review
of every absolute 5% repeat or paired regression trigger (normal latency, peak/RSS,
and instrumented tails). A normal latency claim additionally requires at least 3%
lower medium/large p50 in both repeats; memory/call claims remain separate. Revert
if the allocation gate fails or a material regression cannot be justified.

Use the unchanged 0445 independent ZIP/payload/XML/raw-record and report oracle
(--role plain for both builds). Fixture archive identities are unchanged. Run
full OPC and harness tests, owner strict lint/baseline-debt comparison if needed,
warning-denied owner/harness docs, workspace/feature checks, formatting and boundaries.
No new test merely mirrors the one-line ownership handoff; existing malformed,
URI/content-type, resource, signature, ZIP64 and sink gates cover its semantics.

The accepted ADR tree c950b6c8be822561b498d7bbe87c460873dcbf49 is unchanged from the
prior complete read. No public API, dependency, unsafe code, executor or ambient I/O
change. The full non-iWork goal remains active; this batch does not close native,
semantic CRUD breadth, repackaging, cold/range or scaling gaps.
