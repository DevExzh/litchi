# 0817 execution notes

The previous goal turn made progress: 0816 was committed at `953866d382`,
retaining delayed-source finite-budget evidence and cleaning its owned target.
The current batch follows its queued real-file ordinary-save corpus direction.
The three unrelated workspace files retain their recorded hashes.

Root owns all Cargo, workload, binary, and profiling execution. Agents prepare
and statically review drivers and offline readers; no capture is delegated.
The current 35 normative documents were previously read and revalidated by
exact hash before work. Production and the tracked performance harness remain
unchanged. The existing harness lock differs from the root lock; this is
recorded explicitly without changing dependencies or pooling old timings.

## Before freeze

Static review corrected a draft native build that inadvertently enabled
process metrics and a draft profile that declared settings without exporting
them to Cargo. Native is now feature-off and all release profile settings are
explicit. Review also corrected qualification's observer identity branch,
exporter policy cardinality/nested descriptors, native nullable allocation
metrics, and artifact/qualification acceptance gates. No workload or build ran
with those drafts. Final driver and protocol reviews completed before quality
gate one; all execution inputs are frozen from that point.

The independent audit is intentionally stricter than deterministic hashes:
member preservation and paragraph/cell/shape semantics must be admitted before
timing. Failed checks remain evidence, and no input is silently substituted.

## Quality attempt 0: input guard stopped the run

Formatting and all-feature/all-target checking both exited zero. The following
input-stability guard aborted because the audit script, included in the frozen
packet, was still being finalized. This was a root/agent freeze-coordination
error, not a compiler or test failure. The original driver held its packet hash
dictionary in memory until completion, so the initial dictionary is unavailable;
we do not reconstruct or invent it. Logs, source census, successful commands,
and an explicit guard-failure record are retained under `quality-0/`. Recovery
snapshots are labeled as recovery-time files, not initial packet snapshots.

Root revised the quality driver to serialize all frozen input hashes before
gate one. The next numbered attempt will start only after the auditor handoff.
No release build or benchmark ran in the interrupted attempt.

## Quality attempt 1: stale feature-matrix assertion

The settled-input attempt passed formatting and all-feature/all-target checking.
The full all-feature test command recorded 640 passes and one ignored test in
26 successful suites, then failed the final `xlsx_planning_allocations`
integration test. Its normal-binary assertion expected `none`, but the enabled
`ordinary-save-process-metrics` feature correctly reports
`ordinary_save_procfs_operation_scoped`. The allocator identity has the same
feature-dependent contract. This is a test expectation defect; neither the
runtime harness nor production source changes.

The bounded fix makes both exact expected labels conditional on the existing
feature. All allocation, phase-sum, alignment, corpus, source, semantic, and
output assertions remain unchanged. Root ran Cargo formatting successfully.
Both failed child reports were copied with exact hashes from their owned test
temporary directory before removing that directory. All old frozen packet
inputs and the original test are archived under `quality-1/input-snapshot/`.

The next attempt uses a checked recovery gate: prove all production and tool
files except this one test are identical to the previous source census, verify
all prior successful suite totals and the exact final failure, execute the
corrected integration under all features and allocator-only, and run the
not-yet-executed all-feature doctests. Other quality gates run freshly. This
reuses passing unchanged suites without claiming the failed full command was
successful. The existing owned quality target is retained as a Cargo cache;
no historical performance result is reused or pooled.

## Quality completion and build-driver recovery

Quality attempt 2 passed all six gates. The corrected integration passed once
under all features and once under allocator-only; the doctest command succeeded
with zero doctests. The previous 640 passing tests and one ignored test remain
explicitly inherited through the checked recovery described above.

Build attempt 0 stopped before invoking Cargo because the plan reader omitted
the actual `cargo_bin` field. Its original script and an explicit failure receipt
are retained in `build-0/`. Root corrected that one field lookup. The quality
receipt stays immutable; the next build freezes the corrected driver while
retaining the same plan, Rust source, features, release profile, and lockfiles.

## Artifact admission failed; no timing capture

Build attempt 1 completed all three release binaries. The artifact exporter
completed six corpora and five outputs per corpus. Admission attempt 0 then
exited one and retained 22 audit errors: fifteen generated-metadata accounting
assumptions, three DOCX relationship-order observations, and four XLSX workbook
calculation-property observations across real/generated cases. No successful
admission file was written, and no qualification/native/observer capture ran.

Root compared the actual XML bytes independently: DOCX moves fontTable rId4
without changing the relationship edge set; XLSX inserts/updates calcPr to
request full recalculation. The frozen auditor and plan remain unchanged.
Offline diagnosis distinguishes reader assumptions and expected edit closure
from preservation-order concerns. The batch closes as admission evidence with
zero measured timing samples, not a successful performance baseline.

Root independently checked all six source descriptors and thirty output hashes;
each five-policy output set is byte-identical. The frozen auditor's `--check`
exactly replays admission-0 and returns zero for replay equality, while the
retained audit itself remains `ok: false` with 22 errors. Replay success is not
admission success. No qualification/native/observer directory or successful
admission receipt exists.

## Terminal replay and cleanup

The failure-aware offline reader passes before and after cleanup. Its two
failed passes and bounded corrections are retained in `reader-0/`, `reader-1/`,
and `reader-recovery.md`. Root independently confirmed the diagnosis helper's
DOCX order/edge values and XLSX calculation flags; the helper output is retained.
Root verified all three built binaries before removing the owned target and
scratch directories (9,063 files / 8,591,960,251 logical bytes). Final replay requires
zero timing lanes and failed admission; it does not convert rejection to success.
