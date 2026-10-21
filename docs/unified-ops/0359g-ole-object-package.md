# 0359g — OLE object package owner and consumer migration

- Date: 2026-10-21.
- Base: `45a5f9e9d1` plus the 0359f raw-bounds commits.
- Status: implementation and consumer migration are in progress on
  `uo2/w5-ole`; validation and coordinated performance measurement remain
  pending.

The shared object owner now follows the common package contract. `Package` is
an immutable read owner; `Package::open(bytes)` uses the finite default policy
and `Package::open_with_options(bytes, OpenOptions)` accepts caller-selected
targets and bounds. `Package::edit()` and `edit_with_options(EditOptions)`
start an isolated `Edit`, and `Edit::commit()` returns the shared
`Commit<Package, Patch, Diagnostics>` triad. `Package::apply(&Patch)` and
`Patch::apply(&Package)` use the recorded source content and return the same
triad. Dropping an edit discards its staged state.

The owner retains admitted source bytes and captured stream allocations in
`Arc` values. Read projections and output methods use immutable receivers;
`write_to(&self, ...)` streams an unchanged source directly, while changed
output is rendered only at publication. `source_shared()` exposes the retained
source allocation to deferred format owners, and `OpenOptions`/`EditOptions`
provide the policy getters with the required `_value` names. DOC, XLS, PPT
tests, and the OLE performance harness were migrated to package ingress and
edit/commit publication in the same slice.

The migration removes the former public object-editor open constructors and
staged-snapshot publication path. The raw OLE directory, OLEDS, link, and
protection bounds and typed errors from 0359f remain unchanged. Exact no-op
bytes, source policy fencing, atomic edits, reversible replay, and opaque
stream fidelity are covered by the existing object and format regression
suites; their post-migration run and ABBA receipt will be added before this
slice is accepted.
