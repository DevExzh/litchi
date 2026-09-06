# Validation history

The unchanged source at `dd0681cf8` rebuilt the normal and allocator harness
binaries and passed 349 ODP tests. Full strict Clippy stops at the existing
`litchi-odf-common::ArchiveReaderKind` large-enum-variant diagnostic. Strict
ODP Clippy with `--no-deps` passes; the dependency failure remains visible.

The first pilot adapter addressed binary identity through `tool.binary`, which
is a string label. Its failed receipt, report and original script are retained.
The corrected adapter reads the top-level `binary_identity` object. All six
before pilots then passed the unchanged 0439 oracle and actual binary SHA,
size, path and revision checks. No formal measurement used the initial adapter.

The root owns builds, tests, scripts, pilots and profiling. Agents perform
source-only reviews and write external drafts; the root applies
production changes only between terminal CPU jobs. The user-owned GOAL.md and
the sealed 0439 bundle remain unchanged.

The delegated agents stopped at a usage limit before a final implementation
handoff. The root retained their drafts, completed the driver review, and
removed an unintended one-shot lookup rewrite from the parser draft. Only the
borrowed namespace cache change was applied. The first focused compile exposed
an extra dereference in the new independent test oracle; correcting it made all
18 focused tests pass. The full candidate ODP suite passes 352 tests and strict
owner Clippy (`--no-deps`) passes. Both measurement binaries rebuilt, and all
six candidate pilots passed the unchanged independent 0439 oracle.

The root fixed draft evidence adapters before using them: capture excludes its
own artifact paths from worktree custody, fails on source/status changes, and
enforces the frozen oracle verifier and protocol hashes. Profile capture checks
binary/report identity and the same custody. Summary derivation uses the actual
semantic_shapes and binary_identity schema, retains chronological allocation
vectors and balance checks, and derives peak above entry separately from the
absolute process-global allocator peak. Original drafts remain in draft-history.

The formal A1/B1/B2/A2 matrix passed all 24 reports and 720 samples. One large
normal R2 comparison triggered p95/p99 review. A separately frozen confirmation
plan declares exactly four additional large normal reports in ABBA order.
The root then mistakenly launched its first C1 while the baseline record
profile's text conversion was still running. Both overlapping attempts are
excluded by overlap-exclusion.json, with their original passing command/oracle
receipts and interval evidence retained. Replacement capture uses separately
named output paths after both handles were verified terminal. The main matrix
is unaffected. No latency claim is accepted automatically from an allocation
gate; all adverse tail observations require explicit review.

Both pairs in the fixed 120-sample confirmation had lower candidate p95/p99;
the main tail flags did not recur. The root keeps the change for allocation
calls and cumulative requested bytes only. The profile counter changes remain
under 1% for cycles/instructions, and peak above entry and retained live delta
are unchanged. The selected baseline record retains 15 addr2line warnings in
each text conversion; the candidate record retains 13 in each. Neither profile
is presented as complete symbolization or isolated operation-only attribution.

The final harness library suite passes 330 tests with one ignored; the binary
targets pass a further 34 tests, and the filesystem integration target passes
four tests, completing the prior 368-test harness scope without rerunning the
library suite. The independent ODP suite remains 352 passing tests. Rustdoc
with warnings denied, pinned formatting and crate boundaries pass.

All eight repeat flags remain retained: baseline large normal RSS +9.665%,
baseline large allocator RSS −5.162%, four baseline tiny instrumented timing
statistics about −17.4% to −17.6%, and candidate large normal p95/p99
+9.050%/+8.940%. Allocation counters are identical across repeats. These
variations support neither an instrumented latency nor an RSS benefit claim.

Portable copied controls passed before and after cleanup; all 12 precleanup
and 13 postcleanup corruptions were rejected after refreshing inventories.
Cleanup removed four owned temporary directories totaling 1,830,361,569
regular-file bytes. Both build-cache directory identities and the user-owned
GOAL digest were preserved. The current parser matches the measured candidate;
the accepted ADR tree and sealed 0439 bundle remain unchanged.
