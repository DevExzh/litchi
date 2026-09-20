# 0723 worksheet target checkpoint design review

## Review identity

This review is read-only with respect to production code. The immutable
baseline is commit `45cb480eaa` (`perf(docx): retain writer-local structural
scan fusion`). The baseline XLS/CFB implementation reviewed here is therefore
the source at that commit, with the existing retained 0684 occurrence index,
0685 worksheet-start checkpoint, 0688 checked CFB chain path, 0689 SST-chain
checkpoint, and 0690 borrowed directory lookup.

The current candidate remains an uncommitted worktree layered on that baseline.
The reviewed source hashes are `query_cache.rs`
`97d78292d4aff22c8d4dc1aea43aefcfd04b6fea8f83f4cccc77c3aed49f649a` and
`source.rs`
`30b466bfbcf68e95e424f3242efc903786033401ef46cee5c467d178d60753ce`; the
focused integration test is the uncommitted
`crates/litchi-xls/tests/xls_query_chain_checkpoint.rs`.
The current production references are:

- `crates/litchi-xls/src/workbook/source.rs`: worksheet scan and indexed replay;
- `crates/litchi-xls/src/workbook/query_cache.rs`: weighted index state and
  checkpoint ownership;
- `crates/litchi-cfb/src/shared.rs`: opaque CFB checkpoint restoration and
  hinted cursor creation.

No Cargo, native, or profiling command was run for this review.

## Decision

Retain the existing worksheet-start checkpoint and add at most one target-local
checkpoint per admitted worksheet index. Replacing the worksheet-start point
would preserve value semantics through CFB's backward-position fallback, but it
would make a later query whose first occurrence precedes the original selected
target restart from the Workbook stream's first sector. That would discard the
0685 worksheet-prefix saving. The index is shared across arbitrary coordinates,
so replacement is not a safe performance design.

The bounded representation is:

```text
target_chain_checkpoint: Option<(frame_offset, StreamChainCheckpoint)>
```

The offset is needed because the opaque CFB checkpoint intentionally exposes no
physical ordinal. Replay selects the target checkpoint only when the first
matching slot is at or after `frame_offset`; otherwise it uses the original
worksheet-start checkpoint. `slots_for` already returns matching slots in
stream-offset order after publication.

## Why capture must be a post-scan metadata probe

`scan_worksheet` creates the worksheet cursor at the validated sheet start and
calls `CellSink::started` immediately afterward. The scanner then uses a
read-ahead window. Its CFB cursor commonly sits at the end of the filled window
while `WorksheetScan::position` still identifies the current BIFF frame. The
`StreamChainHint` passed into the initial cursor creation is not advanced by
subsequent cursor reads; it records the position selected by
`stream_cursor_at_hinted`.

Consequently, taking that hint after the selected frame has been consumed does
not produce a target-frame checkpoint. A cursor-level API would have to expose
or reconstruct a chain state at the exact frame boundary despite read-ahead.
That is a wider CFB change and risks coupling the worksheet window to physical
cursor state.

The small alternative is safe: after the scan's existing final source and
execution fences succeed, locate the first matching target slot, restore the
worksheet-start hint, and call `SharedOleFile::stream_cursor_at_hinted` at the
slot's exact `stream_offset`. The call performs directory lookup and validated
FAT/MiniFAT metadata traversal, but it does not call source `read_at`; source
freshness remains governed by the existing scanner and replay reads. Capture the
checkpoint from that temporary hint and discard the temporary cursor before
publication.

If this optional metadata probe fails, omit the target checkpoint and publish
the ordinary index. A successful selected query must never become an error only
because optional replay state could not be prepared.

## Semantic and resource boundaries

The target checkpoint must be prepared only when the candidate is still active
and after `scan_worksheet` has completed. Abandoned, cancelled, stale, and
failed candidates must release it with the existing worksheet and SST
checkpoints. The added tuple needs a fixed logical charge before slot
collection; the current 224-byte `INDEX_OVERHEAD` should become 264 bytes if
the measured layout is 40 bytes larger.

The checkpoint retains no source bytes, decoded values, errors, SST state, or
per-slot cursor. Its weak CFB identity must retain the current foreign-reader,
expired-reader, stream-identity, and backward-offset fallback rules. Replay
must continue to parse every selected frame and decode XF, SST, formula STRING
and CONTINUE grammar at the existing point, preserving duplicate order and
typed-error precedence.

## Evidence that makes the hypothesis testable

The retained 0689 route traces show worksheet-prefix links from the worksheet
start to the selected frame of 14 for 54016 first, 1,044 for 54016 late, 6 for
WithCustomViews first, 72 for WithCustomViews late, 2 for 45365-2 first, and
258 for 45365-2 late. These values are recorded in
`docs/performance/results/change-0689/route-comparison.json`.

The prospective trace must build the index from a late target, then query an
earlier coordinate, the original target, and a later coordinate. It should show
the original worksheet-prefix count for the earlier coordinate, zero skipped
prefix links for the original target, and only the forward suffix links from
the retained target position for a later coordinate. A fixed-target-only probe
is insufficient because it cannot prove the backward fallback rule.

The native matrix should separate the one-time candidate-build cost from warm
replay behavior and include 54016 first/late, WithCustomViews first/late,
45365-2 first/late, Simple, missing, numeric, formula-refusal, and zero-budget
controls. Counted source ranges, bytes, and freshness observations must remain
identical for the candidate build; the metadata probe must add no source I/O.
Allocation evidence should distinguish the 40-byte logical retained charge
from allocator and RSS observations. No broad CRUD, cold-device, remote,
concurrent, native-Office, cross-platform, or iWork claim follows from this
warm selected-query experiment.

## Required correctness checks

The existing XLS differential suite should be extended or explicitly audited
for these cases:

1. Build from a late coordinate, then query earlier, target, later, missing,
   duplicate, SST, and formula-STRING coordinates in alternating order.
2. Verify clone sharing and reopened-owner isolation; the weak checkpoint must
   remain local to the parsed CFB index identity.
3. Verify optional candidate abandonment under allocation refusal, cancellation,
   source change, malformed tails, and intrinsic admission refusal.
4. Verify exact values, errors, duplicate order, source observations, and source
   ranges against an index-disabled owner.
5. Verify cache accounting at the old and new overhead boundaries and ensure
   zero query-index budget never constructs target state.

The source-level proof should retain the existing CFB identity tests and add
only the XLS-specific target-selection and backward-after-late evidence. No
new public API or physical CFB identifier may be exposed.

## Disposition before implementation review

The post-scan metadata-only seek is a justified, bounded candidate for
measurement. The replacement-only design is rejected because it loses the
worksheet-start optimization for earlier coordinates. Retention still depends
on the frozen primary latency gates, explicit build-cost and tail reporting,
unchanged I/O/freshness vectors, exact accounting, and the targeted
backward-after-late trace.

## Independent implementation review

Reviewed against the candidate worktree above, without changing production
sources or running Cargo/native/profiling commands. The implementation matches
the bounded design at these points:

- `crates/litchi-xls/src/workbook/query_cache.rs:17,40-53,215-229,329-335`
  charges the fixed overhead before collection and carries one optional target
  tuple through publication. `:646-673` clears it on abandonment and drop.
  The added lifecycle test at `:1333-1383` checks abandoned, dropped, and
  published candidates against child and parent memory reservations.
- `crates/litchi-xls/src/workbook/source.rs:3508-3554` performs the optional
  probe only after a successful full scan. It restores the worksheet-start
  checkpoint, calls the existing metadata-only CFB constructor, and silently
  falls back if allocation or cursor construction refuses. No source read or
  freshness fence is introduced by this path; `shared.rs:800-864` confirms the
  constructor records the hint only after validated allocation metadata
  traversal, while source reads remain in `shared.rs:2896-2925`.
- `crates/litchi-xls/src/workbook/source.rs:3587-3628` chooses the target
  checkpoint only when the first matching slot is at or after its saved frame
  offset. Earlier coordinates therefore retain the worksheet-start route;
  equal-offset cells in one BIFF frame and later offsets may use the target
  route. The SST hint remains independent at `:3581-3585`, and selected frames
  are still parsed and decoded at `:3610-3622`.
- `crates/litchi-cfb/src/shared.rs:188-235,888-910` keeps the existing stream,
  reader-identity, and backward-position fallback rules. The XLS change adds no
  public CFB API or physical sector identity.

The focused test file adds useful value/refusal boundaries at
`crates/litchi-xls/tests/xls_query_chain_checkpoint.rs:235-349`: a late build
followed by indexed queries, a cross-sector exact-frame read with an armed
later-sector fault, and a missing-coordinate build. The cross-sector test is a
good safety check for the no-read-ahead claim. The test named
`target_checkpoint_keeps_forward_and_backward_replays_identical` currently
builds at row 4000 and then queries rows 1, 4500, and 4000
(`:242-266`). It now exercises both backward fallback and a genuinely later
coordinate, then revisits the origin; the disabled-index oracle covers the same
three values (`:268-285`). The cross-sector test also has a positive cold-scan
fault control (`:316-324`), so its marker distinguishes exact replay from a
read-ahead scan.

One boundedness observation remains visible in the source: the metadata walk in
`source.rs:3536-3549` has no `ExecutionContext` check inside it. The walk is
bounded by the already admitted first target slot and the validated CFB chain;
`publish` checks cancellation afterward, and `scan_target_cell` still returns
the already-found value when publication refuses (`source.rs:3509-3515`). This
matches the prior post-scan cancellation interval: cancellation during the
walk drops the optional state and returns the ordinary successful result. It is
therefore an explicit measurable interval cost, not a skipped freshness fence
or an identified contract violation. The temporary path vector is also outside
the logical reservation (`source.rs:3536-3540`); allocator evidence should show
that its bounded scratch cost is the stated allowance.

The requested targeted trace is still pending in this worktree. The current
driver at `docs/performance/results/change-0723/trace.py:315-375` runs each
case's one fixed coordinate repeatedly, and `trace-analyze.py:187-245`
explicitly expects the candidate replay vector to be all zero. There are no
`trace/baseline` or `trace/candidate` outputs to audit. Consequently this
review confirms the source-level fallback rule but does not claim physical-link
evidence for the sequence late-build → earlier coordinate → original target →
later coordinate. The implementation is conditionally acceptable for the
bounded candidate; final retention remains pending that trace plus the root
lane's frozen build/evidence gates.
