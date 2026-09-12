# Executable correction of the new MCE tests

The initial candidate common-crate test command failed four of the nine new
regressions. Its raw receipt and stdout remain in `candidate/`. Static review
had missed two established processor behaviors: comments outside the root are
not emitted, and an emitted fallback child can carry inherited namespace
attributes. Neither failure identifies a difference between the old and new
namespace-search predicates.

The final tests place the lexical-trigger comments inside the root, preserving
the requested byte offsets, and assert the fallback element/text independently
of added namespace attributes. They still require owned processing, exact
borrowed source sharing, every single-byte near-match mutation, typed limit
and malformed errors, report counters, and wrapper `Arc`/`Cow` behavior.

All nine corrected tests pass against the original baseline production under
the separately bound `test-reference/` receipt. `final-source-binding.json`
binds the corrected test. Final source additionally contains the equivalent
`repeat_n()` spelling in an existing XLSX `cfg(test)` helper, recorded in
`test-lint.patch`; code before that test module is unchanged. The complete final quality campaign then exercises the corrected tests
with the candidate implementation. The original formatted test, final test,
and exact correction patch are retained. Initial failed tests are not counted
as a passing quality gate or hidden by the final rerun.

The first complete lint attempt is preserved in `quality-before-lint/`. It
found the existing XLSX test helper spelling under warning-denied Clippy.
The final quality campaign reruns after the test-only lint correction.
