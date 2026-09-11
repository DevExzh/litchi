# 0512 operation scope review

Independent read-only review of the current harness confirms these boundaries:

- `run_xlsx_update_commit` opens and stages before the clock. Only
  `Edit::commit` is timed; verification and retained-commit replacement/drop
  are outside. `prepare_xlsx_updates` does not commit.
- `build_xlsx_workbook` commits while constructing the fixture, outside the
  case runner. A commit toggle alone therefore includes unwanted setup.
- `run_xlsx_update_commit_save` constructs expected output with another commit
  and save after entering the runner but before its loop. Its clock brackets
  commit plus sequential write. Sink reservation, equality/reopen/cell oracles
  and drops are outside.

The selected profile uses one commit-only case per process, collection off at
start, an exact commit enter/leave toggle, and zero-before the exact update-
commit runner. Require three commit calls after reset for three samples and
zero warmups. Source opening, edit staging, fixture construction, final
verification, save and drops are outside collected costs. Simulated instruction
references are not native cycles or phase clocks.

The save runner cannot use this reset to remove its expected-output commit:
that oracle runs after helper entry. Exact save phase attribution therefore
requires a post-oracle marker or a dedicated equivalent region. No such
production or harness change is made in this investigation.

Generic commit cases have no operation allocation measurements. Save's
`promote_sink_operation_metrics` carries deterministic accepted bytes, write
calls, largest write and buckets; it does not measure allocation, copies or
compression. Reusing `allocation_metrics::begin` / `Region::finish` around the
existing clock, paired with `InProcessObservation` and the separate allocator
binary, is the required enabler before an operation-memory claim.

Dense-wide has two 256-by-256 sheets, 131,072 cells and 1,311 updates. The two
native repetitions provide current descriptive timer observations; whole-child
hardware includes generation, opening, staging, warmups, final readback and
teardown. These scopes must remain distinct in every result.
