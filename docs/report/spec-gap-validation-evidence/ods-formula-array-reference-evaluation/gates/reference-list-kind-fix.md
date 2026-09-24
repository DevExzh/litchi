# Reference-list kind correction

The candidate preserves `ReferenceList` separately from rectangular references while
planning evaluated reference operands. Reference operators retain list geometry;
unary and ordinary function consumers apply list conversion rules. `IFERROR`
can catch the resulting scalar `#VALUE!`, while matrix operands still keep their
array kind when combined with a list.

The initial list correction passed the existing 62 integration tests but collapsed
list-plus-array results into scalar errors. Three additional nested regressions
reproduced that mismatch; the final correction makes arrays and rectangular
references dominate the scalar list conversion for matrix mapping, in either
argument order. A list consumed by a range operator remains a reference.

Validated `value.rs` SHA-256:
`1ddaf1adcf4157c3f84bdbae62aa882867037da85a97aab1bd9a9a225eda7e5e`.

Evidence:

- [Initial counterexamples](reference-list-kind-draft-42.json), with exact draft
  source, tests, and logs in its adjacent archive.
- [Corrected directed checks](reference-list-kind-draft-44.json): all 62 tests
  pass, followed by an additional nested error-priority control.
- [Full integration gates](planner-integration-45.json): tests, strict Clippy,
  strict rustdoc, four doctests, and formatting all pass. The 455-file scoped
  source closure matches the canonical workspace and is unchanged during checks.

The original RRK-1 counterexample and direct/unary list fallback variants now
pass. Both ordinary and nested matrix operands reject the new list/array cases
with `Unsupported(ReferenceOperator)` as required by the existing runtime contract.

Two directed checks for a suspected error-priority discrepancy also pass: a
left `#N/A` plus a list is caught by `IFNA`, and the corresponding nested handler
returning an inline array retains the reference-operator refusal. A subsequent
unvalidated [proposed error-priority change](deferred-list-error-priority.diff)
was deferred because these checks did not reproduce a failure. It is not part of
the production candidate. Further review of scalar error precedence remains
separate from the reproduced and corrected list/array classification bug.

This closes the reproduced list-kind regression for the tested profile. It does
not establish complete formula support or performance acceptance. The retained
release-03 capture predates this fix and is not performance evidence for it.

## Follow-up review

The reviewer independently checked the validated `1ddaf` snapshot and confirmed
that list-plus-array precedence matches `map_binary` and `map_function` in both
operand orders. No confirmed public `IFNA` blocker remains in current call paths.
The kind fold can report a different scalar error order than ordinary `XOR`, but
its current callers use it for shape rejection and recompute scalar values.

A separate source-only concern remains unproven: the geometry helper maps omitted
handler alternatives to `Empty`, whereas ordinary evaluation preserves `Missing`.
The reviewer did not establish an externally visible mismatch. Neither concern
is presented as a reproduced failure or as evidence of complete VM acceptance.
