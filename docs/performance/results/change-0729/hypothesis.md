# 0729 public DOC phase attribution and observer controls

0728 identifies a duplicate rendering in the DOC route and measures common
Reuse finish at about 29–31% of a common-container control. That control is
not the public format owner. Measure the actual public lifecycle before making
an end-to-end optimization prediction or retaining a render cache.

Use unchanged production source and the two 0728 DOC fixtures/paragraph-zero
replacements. Compare four routes within the same diagnostic-feature build:
ordinary opaque lifecycle, ordinary lifecycle with external phase clocks,
existing profiled open/commit with empty observers, and profiled open/commit
with bounded timestamp observers. This distinguishes outer timer overhead,
profiled implementation overhead, and timestamp callback overhead. It does not
measure the compile-time effect of enabling diagnostics versus a default build.

The outer lifecycle preserves open, Edit construction, replacement, commit,
output extraction and local destruction. Measured outputs remain alive for
untimed oracle checks. Inner semantic events must be exactly balanced and
successful with their expected sequence; nested fractions must use the same
measured owner, not medians from independent processes. Residual time includes
unattributed scope/drop/timer costs and is not silently assigned to a phase.

All source, corpus, probe, semantic oracle, build, command and analysis identities
are frozen after qualification. Preserve every sample and process; report
per-process p50/mean/p95/p99/maximum and paired process differences. Flag both
mean and p50 observer/control changes above 5% in either direction; these are
measurement-interpretation flags, not production-retention gates. No small-time
exception, trimming, outlier exclusion or adaptive repeat is authorized.

Three fixed cycles, three rounds per cycle and both cases yield nine processes
per route/case. Route order rotates by round and reverses in alternate cycles;
case order alternates as well. Each process performs three warmups and fifty
measured lifecycles. This is 72 processes and 3,600 measured lifecycles. The
0728 allocation record remains separate; no new allocator, RSS, cold-cache,
remote-source, concurrency or production speedup claim follows.

Next implementation work remains conditional on current public-owner evidence.
Any validated-render handoff must retain mandatory validation, atomic publication,
no-op semantics, source/policy freshness, successive-edit behavior and bounded
retained memory. Existing 0652 placement and 0663 Reuse decisions remain binding.

Source inspection identifies a profiled-path lifetime difference: ordinary open
holds its strict editor through public-reader validation and source retention,
whereas the profiled strict-owner closure drops that editor before the later
phases. Ordinary-split versus profiled-empty therefore measures the existing
profiled implementation, including lifetime/code-path differences, not just
callback dispatch. Do not alter production to hide this distinction.
