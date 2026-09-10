# Development receipts

Receipts under `validation/` are retained even when they are not final-source
evidence. `final-gates.json` identifies the accepted commands; successful exit
alone does not make an earlier gate eligible.

- `substrate-build1` prewarmed dependencies while source edits were in progress.
  Its source manifest changed during the command.
- `opc-check1` found an unnecessary trait-object cast in the new integration
  test. `opc-check2` found an unused fixture constant. Both were corrected.
- `harness-check1` used a workspace package selector for the independently
  rooted benchmark package. `harness-check2` used the explicit manifest path
  and passed.
- `prewarm-build1` overlapped benchmark edits and found a missing exhaustive
  `(Exact, None)` match arm. The arm was added. That build is not retained as
  an accepted executable.
- `opc-tests1` passed 386 library tests and failed 12 filesystem tests with
  `QuotaExceeded` while seeding temporary files in `/tmp`. The source was also
  changing during this development run. Subsequent test commands use the
  explicit batch-owned `TMPDIR` recorded in their argument vectors. A direct
  write/read probe of that directory succeeded.

The cleanup-only turn between implementation steps freed 114.6 GiB of old
build artifacts and rebased onto the fetched remote branch without changing
the local source. It preserved the active spec-gap worktree and all source
worktrees. That separate cleanup receipt is outside this measurement bundle.

- `opc-tests2` exposed four unsuitable fixture assumptions (whole tiny archives
  may be prefetched; opaque non-Parts are not topology members; precompressed
  authorization requires a relationship leaf; generic materialization does not
  synthesize an opaque Part) and one exact-path zero-length provider-trace
  regression. Fixtures now use typed leaf members, and exact mode delegates
  zero-length reads to the original source helper.
- `opc-tests3` passed but overlapped source edits. It is development evidence.
- `docx-tests1` exceeded the 64 MiB managed fixture budget during default append
  preparation. The test now selects finite limits for its small append: a
  4 KiB token window, depth 32, 4,096 events, and 2 MiB workspace/output ceilings.
- `final-opc-tests` found two unnecessary `std::sync::Mutex` qualifications in
  the newly extended test after `Mutex` became imported. These were removed.
  This attempt also overlapped the DOCX fixture correction.

- `final-docx-tests` passed the complete DOCX test/doc-test inventory, including
  the corrected managed publication test, but the OPC test qualification fix
  landed during compilation. `final-docx-tests2` is the source-stable rerun.
- Cleanup receipt and final-gate verification were strengthened during
  independent review before protocol freeze. Production cleanup now validates
  complete build/gate custody before planning and immediately before removal.

- `final-opc-tests2` passed 626 tests across 21 result groups (one ignored),
  and `final-docx-tests2` passed 1,390 across 49 groups (31 ignored). The next
  warning-denied Clippy gate requested `inspect_err` for a side-effect-only
  `map_err`; the equivalent expression was adopted and final checks rerun.

- `final-opc-docx-clippy2` and `final-opc-docx-clippy3` identified two test-only
  simplifications: a redundant `u64` conversion and `16.min(4)`. Their corrected
  form passed `final-opc-docx-clippy4`; final-source OPC/DOCX suites are selected
  by the third test attempts. `final-format2` and `final-harness-clippy2` bind
  the same source.

- `final-harness-doc` rejected a public module comment linking to private
  `CountingReadAt`. The reference is now plain code text. Although executable
  logic is unchanged, all selected final receipts bind the resulting complete
  source manifest; earlier passing gates remain development evidence.

- Final cancellation review found that a managed reader queued behind an
  active provider could wait without observing cancellation. Admission now
  checks the context before blocking and polls every 10 ms, matching the
  existing payload-cache flight wait. Gated tests cover both cancellation
  after queueing and an already-cancelled waiter while the first provider
  remains blocked. Publication still drains an already active provider before
  releasing its window. All final receipts are refreshed after this fix.
- `build-normal1` completed successfully while that cancellation change landed;
  it is prewarm evidence and is not a retained executable.

A further out-of-tree cleanup removed an inactive incremental cache after two
empty process-reference audits and acquisition of its Cargo lock. The newest
entry was over 16 hours old. `external-incremental-cleanup.json` records the
117.8 GiB of unshared blocks removed; the protected source worktree was not
modified. This filesystem cleanup is separate from the bundle's own bounded
`cleanup.json` contract and is not a performance measurement.
