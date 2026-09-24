# OpenFormula database functions: implementation evidence

The value evaluator implements all twelve OpenFormula §6.9 database functions.
The [implementation profile](implementation-profile.md) records supported inputs,
conversion, resource boundaries, and explicit host choices; the
[contract](contract.md) identifies the normative source. Release performance
acceptance remains open, and this family does not establish full Small Group
conformance.

Candidate-09 passes all 33 focused database tests, 1,140 ordinary ODS tests,
strict all-target Clippy, strict rustdoc, five doctests, and formatting. All five
gates ran against the same frozen 469-file source closure; source hashes remained
unchanged throughout. Exact sources, commands, toolchain, logs, and receipts are
retained in `candidate-09`.

## Captured diagnostics

- `baseline-01`: the initial seven evaluation test groups compiled and all
  failed because the database functions were unsupported. A group stops at its
  first failed assertion; this is not evidence that every function vector ran.
- `candidate-02`: a frozen 469-file source closure failed compilation with five
  integration errors. No tests ran. The diagnostic identifies a shadowed field
  selector, resource-error conversion, reservation member naming, header lookup
  result shape, and numeric parsing predicate signature. Numeric kernels were
  present in this snapshot but not yet connected to dispatch.

Each capture contains `source.tar.gz`, `receipt.json`, and the exact `test.log`.
The receipt records the command, toolchain, environment, source hashes before
and after execution, and archive/log digests. Root verified archive members and
receipt hashes before moving the captures from disk-backed scratch storage.

- `candidate-03`: the five original integration errors were fixed; three
  deny-unused diagnostics remained, so no tests ran.
- `candidate-04`: production compiled; 12 of 15 functional test groups, all
  seven resource groups, and all three numeric groups passed. Failing functional
  assertions exposed an incorrect sample-variance oracle and success fixtures
  that included an intentional criterion Error record.
- `candidate-05`: added lazy criteria reads, header/error profile fixes, text
  ceilings, and omitted-field dispatch; the same three oracle issues remained.
- `candidate-06`: 17 of 18 functional groups and all ten resource/three numeric
  groups passed. The final failed oracle pointed at the wrong criteria range
  (`S3:S4` instead of the populated `S4:S5`).
- `candidate-07`: corrected that reference and added an actual Empty criterion
  reference regression; all focused tests and full crate gates passed.

Release performance evidence and review follow-ups remain required before the
family can be reported as production-ready.

- `candidate-08`: preserves the VM's private Missing-element error invariant
  and removes unreachable ignored-header criteria state; all 32 then-existing
  focused groups pass.
- `candidate-09`: adds the public computed-matrix error-preservation regression;
  all 33 focused groups and all five full-crate gates pass. The accompanying
  `missing-mutation-09` check shows that the public test does not distinguish the
  defensive private Missing conversion; see [review.md](review.md).

## Performance evidence

`performance/release-07` retains the first release build and 87 successful
children (29 cases × three rounds). It predates the final Missing/criteria-state
changes, so it does not describe candidate-09's exact allocation layout. Its
measured requested bytes were fully released, with equal live bytes before and
after the measured scopes. `performance/release-09` repeats all 87 successful children against the final
candidate-09 source, with stable checked metrics across rounds and all measured
allocations released. Two common-value comparison windows now retain 324 child runs with no
result/work/read/allocation-counter changes. One initial timing flag was not
reproduced, while peak-RSS flags remain under review. The separate live-memory
probe found higher executable RSS and equal median heap/stack RSS for its
sustained range scenario. See the [performance report](performance/report.md);
no overall speedup or unconditional performance acceptance is claimed.
