# 0773 — integrate explicit filesystem-save durability

Status: integration validated; explicit policy retained. `performance_claim: none`.

This integrates the owner-authorized 0761 implementation onto production source
`6074e10e57`. The four integration commits end at `339572acbf`; main subsequently
advanced to `d5aaecabed` through documentation and evidence only. The original
0761 worktree and its unfinished evidence remain untouched. The authoritative
current validation is in [this packet](results/change-0773/README.md).

Ordinary saves retain `Durability::Full`. A caller can explicitly select
`FileOnly` or `NoSync` through the new `*_with_durability` methods. Full syncs
the staged file, replaces the destination, then syncs the parent directory;
FileOnly skips the last synchronization, and NoSync skips both. Every level
still stages and flushes the complete artifact before one replacement, retains
route validation, and distinguishes failures before replacement from a typed
`Committed` error after replacement. Weaker levels deliberately give up crash
persistence guarantees; they are never selected implicitly.

The implementation covers OPC, DOCX, XLSX package/workbook, plaintext PPTX,
XLSB, four CFB publication routes, DOC/XLS/PPT writers and their listed
source-backed overlays. Encrypted filesystem saves and DOCX tail append retain
their Full-only behavior. ODF and iWork are outside this change.
The ordinary-save harness accepts `--save-durability` only for supported
lifecycle and atomic-publication selectors, and records the selected level.

## Architecture and integration

| Constraint | Evidence |
| --- | --- |
| ADR 0005, explicit I/O policy | Owner decision 5 in change 0758 authorizes the per-call policy; the 0761 amendment is retained. |
| ADR 0002/0024, ownership | The public enum lives in core; OPC and CFB own publication; formats forward the policy. Crate-boundary gate passes. |
| ADR 0003/0006, preservation and failure | Output equality, exact-source authorization, validation, cleanup and typed post-replacement failures are exercised by route tests. |
| Existing resource/cancellation rules | The switch controls synchronizations only; core/OPC/CFB cancellation and source-check tests pass. |

The production commits cherry-picked without source conflicts. The ADR append
conflict was resolved by retaining both the current budget-lease amendment and
the complete durability amendment. `integration.json` records original and
integrated commits and hashes all changed files. Independent source review
found no integration blocker; its evidence concerns are retained in `review.md`.

## Paired syscall evidence

The fresh debug probes use one dependency lock and identical input bytes.
The baseline emitted 24 marked save windows; the candidate emitted 96 across
12 routes, existing/absent destinations, and default/Full/FileOnly/NoSync.
Both processes exited successfully. Every published output hash matches its
baseline and other policy levels. Default-save sequences match the baseline,
and explicit Full matches the candidate default after address normalization.

| Policy | File sync | Parent-directory sync | Replacement |
| --- | ---: | ---: | ---: |
| Default / Full | 1 | 1 | 1 |
| FileOnly | 1 | 0 | 1 |
| NoSync | 0 | 0 | 1 |

This establishes removal of requested synchronization work on the captured
Linux routes, not a latency ratio. FileOnly removes the directory descriptor
lifecycle along with its sync; NoSync additionally removes the file sync.
The first verifier rejected the directory lifecycle difference because it
retained debug Rust's `fcntl(F_GETFD)` check immediately before closing that
directory descriptor. The initial failure is archived; the corrected verifier
replays the same raw traces without rebuilding or rerunning native work.

## Current verification

All commands run serially on Linux with Rust 1.95.0, offline/locked dependencies,
two build jobs, incremental compilation disabled, and dev debug information
disabled. These are correctness and syscall checks, not latency measurements.

| Gate | Result |
| --- | --- |
| Ten owners: format, all-features/all-targets check | Pass |
| Core/OPC/CFB full tests | 1,738 passed, 2 ignored |
| Seven format durability suites | 8 passed |
| Ten owners: warning-denied library Clippy and rustdoc | Pass |
| Crate boundaries | Pass |
| Seven format owners, full tests | 10,466 passed, 73 ignored |
| Five dependent owners, full tests | 823 passed, 4 ignored |
| Facade with stated Office/ODT features | 382 passed, 7 ignored |
| Full performance harness library suite | 555 passed, 1 ignored |
| Standalone harness formatting | Pass |

Counts include the suites and doctests emitted by each recorded command; the
focused durability tests overlap the full format run and must not be added as
unique coverage. Exact commands, source census, logs and exits are retained.
Two initial quality attempts failed because the isolated worktree lacked an
ignored LibreOffice fixture checkout. Both failures are retained. Linking the
existing reference corpora enabled the successful third attempt; fixture hashes
and reference locations are recorded in `reference-fixtures.json`; its recorded
HEADs belong to the enclosing repository, not upstream corpus revisions.

The successful core/OPC/CFB run includes injected file-sync failures, skipped-sync
checks, typed parent-sync failures, source fingerprint changes immediately
before replacement, substituted-temporary rejection and identity-aware cleanup.
These assertions complement the successful syscall windows: traces alone do
not establish failure atomicity or cancellation behavior.

The old 0761 scratch measurements and incomplete gate record are not promoted
as evidence. This integration makes no speedup, allocation, RSS, cold-cache,
concurrency, Windows/macOS, power-loss, or native Office application claim.

## Custody and cleanup

`trace-0/` retains both raw traces, all 120 published outputs, the initial
verifier and rejected report, the corrected offline replay, probe source/locks,
build logs, binary identities and pre/post source custody. The corrected
verifier passes every comparison; even the separately reported exact
save-versus-Full sequences match in this capture. No raw trace was rewritten.

Owned build targets and probe binary copies were removed only after all native
processes exited and executable hashes were rechecked. `cleanup.json` records
paths, file counts, logical byte totals and binary identities. The final seal
covers every retained packet file. Unrelated main edits, reference corpora and
external worktrees remain untouched.

This closes durability integration, not the non-iWork performance goal. The
unintegrated XLS writer length-field branch remains a separate next step;
physical cold-cache, remote/range and concurrency evidence remain open.
