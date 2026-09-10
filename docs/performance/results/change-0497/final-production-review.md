# 0497 final production review

This is a bounded, read-only review of the current DOCX atomic-publication
candidate and its retained evidence. No source was edited and no Cargo command,
build, capture, or expensive rerun was performed. The protected
`/home/zhuhe/code/litchi-spec-gaps` worktree was not accessed.

## Production disposition

There is no production correctness blocker in the reviewed DOCX route.
`ParagraphStreamPlan::write_to_path` and
`ParagraphStreamCommit::write_to_path` are consuming methods over the existing
plan/commit products and delegate to the existing OPC sibling-temporary atomic
replacement helper. The callback retains the existing source, authored-replay,
budget, cancellation, and late-output checks before replacement. The focused
route tests cover exact output and reopening, source and destination aliases,
failure preservation, cancellation, limits, permissions, temporary cleanup,
and managed-workspace release. The existing OPC test supplies the injected
parent-sync `Committed` case. Windows same-path/hardlink behavior and crash
durability remain outside the tested contract.

The counting and atomic routes should remain as necessary after-only
capabilities. Counting exercises a bounded nonretaining sequential sink while
using the production publication proof. Atomic exercises the ADR 0005
filesystem publication path, which has no before-baseline equivalent. Neither
route is evidence of a speedup, and neither should be folded into the default
before/after comparison. No source change or benchmark rerun is required for
this limited capability disposition.

## Source and gate custody

The six files in `candidate-source.json` match the current tested candidate
byte-for-byte and by SHA-256. Comparing the full after-build source manifest
against the worktree leaves only the nine explicitly concurrent ODG/iWork
paths; no 0497 DOCX or harness source mismatch was found. The retained final
gates contain 16 passing named gates, and formal1 verification reports 288
passing children and 8,640 samples. Concurrent ODG/iWork changes remain out of
scope for the 0497 commit.

## Formal-result disposition

All 100 adverse analysis flags are from the default `hashing_sink` route. The
`file_store-owned-s64-a64` allocator latency p99 is +73.58% and +75.90% in the
two repeats; its p50 is +7.63% and -2.26%. Both p99 regressions remain
unresolved observations. The retained analysis makes no causal claim and no
atomic speedup claim; these flags must not be dismissed as effects of the
after-only routes.

The 72 allocation live-byte endpoint flags require a separate harness-layout
qualification. The current `Sample` adds an inline
`Option<PublicationRecord>`, whose publication record embeds the atomic output
record, and `samples` is preallocated before each allocation region begins.
The formal live-byte deltas are fixed at roughly 12.5 KiB across the affected
arms and are consistent with that enlarged retained sample layout. This review
does not claim an exact byte-for-byte causal decomposition without a type-size
measurement. Mark those live endpoints harness-layout confounded and make no
production live-memory claim from them. Allocation peak increments had no
adverse flags. Retain the raw rows and their unresolved status.

## Remaining custody

Final cleanup and the evidence seal remain pending. Before sealing, audit
failed-child atomic destinations and private temporary directories, clean the
owned `/home/zhuhe/.cache/litchi-goal-0497` material under the cleanup contract,
and keep the concurrent ODG/iWork paths out of the change-0497 commit.
