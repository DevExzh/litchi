# OpenFormula database functions: implementation evidence

The implementation of the twelve OpenFormula §6.9 database functions is in
progress. [The contract](contract.md) records the local specification source,
required behavior, and explicit repository profile choices. This directory does
not yet establish production readiness or performance acceptance.

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

Focused functional, resource, and numeric regression tests are being developed
alongside the implementation. Full crate gates and release performance evidence
remain required before this family can be reported as complete.
