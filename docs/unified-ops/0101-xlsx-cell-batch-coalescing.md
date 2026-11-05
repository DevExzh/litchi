# 0101 — Coalesce distinct XLSX cell payload commands

This checkpoint advances the M2 transaction kernel toward the required native
batch cost. Consecutive direct-cell value, formula and clear assignments share
one native workbook edit. Native staged cell state supplies per-item change
facts; the facade does not maintain a second cell model.

Repeated cell targets, font assignments, selectors, reads, metadata and other
command families close the pending group. This preserves sequential visibility
and prevents a later write from masking an earlier commit-time refusal.
Best-effort commands retain isolated per-item commits. Native commit failures
retain the originating command index; failure publication remains atomic.

The coalescer admits the existing per-call cell payload/action charge before
staging and closes a run before its cumulative charge exceeds that same bound.
Snapshot retention is charged at native commit. Neither charge measures total
allocator use. Huge syntactic ranges reach the existing preflight before any
grid resolution. Source revisions are fenced during planning, staging, commit
and error-path replay.

The full goal remains open, including broader coalescing, structural XLSX
operations, other format write families, Python/CLI and acceptance coverage.

Validation:

- 870 combined DOCX/XLSX/encryption facade tests/doctests passed, including
  11 coalescing regressions: native commit counts, item change flags, replay,
  inverse, best effort, barriers, resource bounds, source changes and poison.
- 419 minimal-feature tests/doctests passed.
- Three native staged-state tests and all nine native doctests passed.
- Capability snapshots and 14 generated-delta tests passed.
- Warning-denied Clippy passed for both XLSX and ops, all targets, with
  DOCX/XLSX/encryption enabled. Cargo ran offline.

The Clippy all-targets check uses a temporary copy of the existing base
LibreOffice fixture tdf169496_hidden_graphic.xlsx. It is removed during cleanup.

Paired measurements use the new isolated
[unified-xlsx-batch harness](../../tools/perf-baseline/unified-xlsx-batch/README.md).
Two nine-round ABBA invocations per build ran pinned to CPU 0, with no other
validation build running. The corpus is a deterministic blank 2,077-byte XLSX;
opening, promotion, command construction, serialization and result destruction
are outside timing. Every count verifies exact native/unified output bytes,
per-item change results and one native commit. Both timed outputs retain their
patches until after the timer stops.

The first control measured source-backed 1,000-cell batches at 6.60–6.67 times
the native loop. The coalescer was resolving each cell through the retained
source even after write promotion. It now resolves addresses against that
same promoted immutable workbook, while retaining the original source fence
before and after planning, staging and commit. It introduces no shadow cell
model and does not serialize/reopen.

Candidate unified/native median ratios (four invocations: two before and two
after a rustfmt-only rebuild; all samples retained):

| Cells | Eager | Source-backed |
|---|---|---|
| 1 | 1.080–1.107 | 1.099–1.107 |
| 10 | 1.304–1.337 | 1.329–1.341 |
| 100 | 1.746–1.752 | 1.744–1.781 |
| 1,000 | 1.854–1.884 | 1.864–1.894 |
| 10,000 | 1.878–1.928 | 1.857–1.959 |

The 1,000-command 2× review budget is met for this prepared-owner workload.
The one-command batch ratios approach or slightly exceed 10%; this does not
measure the separate single-operation File::set route. It is not acceptance
evidence for large existing worksheets, filesystem-source checks, cold
promotion, repeated targets, formatting, ranges, structural edits or saving.
Those scenarios and the complete performance bar remain open. The control
already coalesces native edits: these measurements isolate target resolution,
not the improvement from replacing per-command commits.

[Evidence and provenance](results/0101/provenance.json) identify both binaries,
the lockfile, compiler, source manifest, corpus, control patch and reproduction
commands. Six small JSON files retain all ABBA samples, including the pre-format candidate. The source manifest
hashes changed crate/harness files against HEAD; unchanged files are identified
by that HEAD. No binary is retained in the repository.



The final promoted-owner implementation passed all 870 combined tests/doctests
and warning-denied Clippy again; 30 targeted cell/coalescing tests also passed.
Scoped rustfmt, whitespace and repository artifact audits passed. The rewritten
base remains an ancestor; no rebase replay is needed for the current base.

Commit staging was attempted, but shared Git worktree metadata remains
read-only (index.lock). No commit was created. Temporary logs, the borrowed
fixture, root Cargo.lock, Python cache and both build target directories were
cleaned. No temporary worktree was created. The complete goal remains active.
