# 0362ugd-a — DOC lazy-entry consumers and profiler state

- **Status:** source follow-up on `uo2/doc-lazy-entry`; pending coordinator
  integration and the scoped Cargo rerun.
- **Date:** 2026-10-23.
- **Input:** [0362ugd](0362ugd-doc-lazy-entry.md), implementation
  `663c1d905388964cfd6a2b516eefcf5efffb273d`.
- **Scope:** consumers and diagnostics affected by the fallible lazy DOC
  primary edit entry. No new editing family or capability was added.

## Consumer migration

The DOC facade now propagates the fallible `properties_mut()` getter and the
fallible `source_bytes()` getter. The coalesced body-run snapshot helpers take
`&mut self` where the getter can materialize the native body owner, and their
errors remain DOC transaction errors rather than being discarded. Ordinary
native edit storage and all direct DOC facade/tool signatures use
`litchi_doc::Edit<'static>`; PPT and XLS metadata consumers remain with their
format owners.

## Observer-state correction

The ordinary `Edit` carries a zero-sized default deferred-open state. A
profiled edit carries its borrowed observer state as a separate generic state,
so enabling `performance-diagnostics` does not make an ordinary edit lose the
`Send`/`Sync` auto traits it had before the lazy-entry change. The profiled
state continues to accept caller-scoped, non-`Send` callbacks. The package
profiling helper explicitly performs the first body access while its recorder
is live, then converts to the ordinary edit before the recorder is finished;
the open phase therefore measures real lazy materialization instead of zero
events. Failed profiled opens retain both observers for a later attempt, and
the profiled strict/fallback path receives the selected protection policy just
like the ordinary path. The DOC transaction coordinator propagates the
fallible coalesced snapshot charge without swallowing its error.

## Checks

The full tracked consumer audit used `git grep HEAD -- crates tools` and
separated the DOC-owned sites from the PPT/XLS sites assigned to their
workers. Direct `rustfmt --config skip_children=true --edition 2024` on all
changed Rust files and `git diff --check` passed. Cargo validation is deferred
to the coordinator's next reserved build slot; the prior 0362ugd tests and
strict Clippy logs remain historical evidence for its unchanged paths.
