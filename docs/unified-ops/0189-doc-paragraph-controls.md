# 0189 — DOC paragraph diagnostic controls

## Change

Added opt-in native policy and retained-reader timing controls to the existing
DOC paragraph harness. The primary native and unified timing boundaries and
route names remain unchanged. A retained-reader sample opens a fresh native
Package and Document before timing and keeps them alive through the ordinary
Snapshot open/edit/commit. No Snapshot preparation moves outside the clock.

This is progress on the performance investigation following record 0188.
It does not change the library or complete the DOC owner migration.

## Evidence

Release build, warning-denied release Clippy, rustfmt check, default-mode smoke,
and diff checks pass. Two nine-round diagnostic matrices cover 1, 10, 100,
and 1,000 paragraphs with empty failure matrices and exact native/unified
warm-up output equivalence:

- [Controls](results/0189-doc-paragraph-controls.json)
- [Repeat](results/0189-doc-paragraph-repeat.json)

Each receipt records source/lockfile and binary digests, build command, corpus
digests, timing boundaries, and raw samples. The executable is not committed.
The external lockfile starts from the root lockfile with cc resolved to the
locally cached 1.6.0. Runs were pinned to CPU 0 without concurrent builds/tests.

| Paragraphs | Facade/native, first / repeat | Retained reader/plain native, first / repeat |
|---|---|---|
| 1 | 1.059 / 1.058 | 1.023 / 1.039 |
| 10 | 1.078 / 1.051 | 1.010 / 1.011 |
| 100 | 1.174 / 1.147 | 1.144 / 1.145 |
| 1,000 | 1.177 / 1.188 | 1.179 / 1.187 |

The policy-control ratios range from 0.978 to 1.038. Retaining a native
reader reproduces an overhead of similar magnitude at 100 and 1,000
paragraphs. This supports investigating live-owner effects; it does not prove
an allocator/cache mechanism or justify removing validation. Diagnostics
interleave extra work, so these receipts do not replace the frozen acceptance
runs in record 0188. The 10% single-write budget remains unmet.

## Remaining scope

Continue the immutable native DOC owner and canonical options migration,
then repeat ordinary paired acceptance measurements. Full DOC CRUD, XLS/PPT
write loops, tier 3, Python/CLI and all other goal acceptance remain open.
No OfficeCLI behavior or capability changed in this diagnostic-only slice.
The branch already contains the current base history (zero base-only commits).
