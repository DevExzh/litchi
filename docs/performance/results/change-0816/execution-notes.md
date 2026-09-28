# 0816 execution notes

The previous goal turn was progress: commit `c8ff2b9f65` retained the fresh
0815 rejection and restored the exact production baseline. Root rechecked the
commit history and workspace. The same three unrelated files remain excluded.
All 35 previously read normative inputs retain their recorded hashes.

Two independent bounded audits considered provider/scaling and native-producer
gaps. The selected gap is the current finite-budget/cache-state cross-product
for delayed CFB and OPC Part reads, not first-ever delayed-provider evidence:
0498 already measured delayed short-read Parts. No production change is planned.
Root owns Cargo, rustfmt, workloads, binary tools, source disposition, cleanup,
and commit. Agents implement/review the harness and prepare offline readers.

## Pre-build review and formatting

Root initial formatting check found formatting-only differences. The original
handoff snapshot and the successful formatting receipt are retained under
`pre-format/`; root then reran the formatting check successfully. Static review
found a test that incorrectly expected an error for `u64::MAX` on 64-bit hosts.
The test was made pointer-width aware without changing source behavior. Root
also resolved a reviewer/reader schema mismatch before freeze: local zero/zero
settings use baseline-v1; either nonzero setting uses range-baseline-v1.
Source and protocol reviews now pass. All drivers and required inputs were
frozen before quality gate one; no benchmark has run yet.

## Quality completion

The original quality process completed all six gates successfully: formatting,
all-feature/all-target compile checking, all seven harness tests, warning-denied
Clippy, warning-denied rustdoc, and the full repository crate-boundary checker.
No quality gate was retried. Production, harness, lockfiles, normative inputs,
and unrelated files remained stable under the frozen checks.

## Qualification acceptance

Both release binaries built successfully. All 72 qualification processes
completed. Root independently checked corpus and full byte/order verification
objects against sealed 0786 large/mixed reports, plus source-counter totals and
resource bounds/release. Its new offline reader initially needed two interface
corrections: 0786 seal keys are packet-relative, and reusable sessions may hold
worker permits after the operation until teardown. The final check enforces
bounds at each snapshot and zero worker/I/O permits after drop. No raw report
or workload was altered or retried. The accepted immutable qualification audit
was written before native capture began. Capped CFB calls are 128/125 for
large/mixed payloads; fresh Parts remain 64 calls, and primed Parts have zero.

## Capture completion

The native and observer drivers completed on their original root process
handles. The frozen schedule retains all 432 native reports and 144 observer
reports, plus the 72 qualification reports: 648 reports / 13,320 samples.
No measured child failed, was retried, or was removed. Only after the observer
process terminated did root authorize heavy offline analysis and independent
result review. Root separately recomputed paired width-one/width-eight p50
ratios directly from raw reports; this is a cross-check, not another capture.

## Offline reader recovery

Root review found two errors in the first offline analysis: nearest-rank
bootstrap notation selected lower endpoint 249 instead of frozen index 250,
and finite Amdahl fits outside [0, 1] were labeled valid. The first numerical
outputs and reader snapshots were preserved under `analysis-attempt-0/` before
correction. The corrected endpoint changes no numerical result; admissibility
flags change in eight of 18 families, repeated across 32 of 72 width rows.
Unconstrained fractions, clamped fits and residuals remain available. No raw
report or frozen driver changed; no workload was rerun.

Root replayed qualification acceptance successfully, independently checked
all 18 report-table families (four widths and paired speedup intervals), and
ran the aggregate validator successfully before cleanup: 648 reports, 13,320
samples, 72 scaling rows, 72 independent audit rows and six quality gates.

## Cleanup

Root removed the owned build directory after verifying both captured binaries: 2,113 files / 808,118,216 logical bytes. Production and the three unrelated files remained unchanged. The pre-existing tool lockfile was preserved.

The first post-cleanup validator invocation exposed a reader interface error:
`load_cleanup` looked for an older `executables_verified_before_removal` key,
while the root cleanup driver and final validator use
`binaries_verified_before_removal`. Root corrected the unfrozen reader to
require the exact 0816 cleanup schema, removal marker and binary-verification
marker. The cleanup witness, binaries' recorded identities, frozen drivers and
raw measurements were not changed.

The next invocation caught a related comparison error: the aggregate reader
compared the raw binary descriptor against its normalized form, which adds
`custody_verified`. It now compares the exact path/size/SHA-256 descriptor;
the derived custody marker is checked by the binary reader. This is also
an unfrozen reader correction, with no evidence or numerical changes.

The corrected post-cleanup validator passes all retained checks with
`cleanup_checked: true`; deterministic numerical outputs remain identical.
