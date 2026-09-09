# 0489: reuse a prepared OPC candidate XML audit

OPC splice preparation already validates the complete source and candidate XML.
Later ZIP measurement, emission and expected-artifact preview passes now reuse
a private capability bound to the immutable proof, target and frozen limits.
They retain complete byte/EOF/hash authentication, freshness, cancellation,
Work, conservative XML workspace admission and exact accepted-output counts.
Final DOCX semantic reopen remains required.

The [144-process comparison](../results/change-0489/results-review.md) shows
about 20–21% lower source-heavy owned/file median latency, 14–15% lower
authored-heavy file medians, and about 19.89% lower deterministic operation
heap. Candidate archive bytes are unchanged. A small file-store repeat has
slower tails, and four whole-child RSS observations increase above 5%; the
review records every adverse row and leaves their cause open.

OPC/DOCX feature-matrix tests, format, Clippy, rustdoc, crate boundaries,
serial allocator harness tests, Python helper tests, and two 10,000-run ASan
campaigns pass. The evidence retains a development Clippy failure for test
literal formatting and its successful final retry. The change adds five
adapter cases and a public 36-combination late replay failure test. It adds
no public API, dependency or unsafe code. The broader non-iWork goal remains open.

The [cleanup receipt](../results/change-0489/cleanup.json) records removal of
all scoped Cargo output and unused batch scratch while retaining authenticated
benchmark/fuzz executables and reproducible evidence. The spec-gap worktree
and other sessions' active files are excluded.
