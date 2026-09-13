# OpenFormula complex functions

The implementation provides all 26 ODF 1.4 Part 4 §6.8 functions in the
scalar and value evaluators. [Contract](contract.md) records exact local source
hashes, conversion and sequence rules, selected interpretations of conflicting
IMSQRT/IMSECH text, and independently computed extreme-value oracles.
[Implementation profile](implementation-profile.md) records the public value,
ordinary coercion, numeric, and resource policies.

## Missing-behavior baseline

[Baseline 06](baseline-all-06/receipt.json) compiles both integration targets
against the pre-complex matrix-06 implementation, with a 460-file source closure.
Eight conformance tests fail; of eight resource tests, three existing
budget/cancellation boundaries pass and five fail. Failures show the missing
complex evaluation path. The conformance test source contains vectors for all
26 functions, but a test stops on its first failed assertion: this baseline
does not claim that every vector executed.

[Baseline limits 03](baseline-limits-03/receipt.json) records the earlier
resource-only result against 459 source inputs. Diagnostics 01/02/04/05 are
compile failures in new test drafts, not production runtime regressions. Root
fixed nonexistent limit-builder calls, helper shadowing, unused fixture code,
unnecessary qualified paths, invalid value dereferences and Result handling.
The ownership test explicitly converts a complex result to Text before testing
text reservations, rather than requiring a Text representation for Complex.

Every capture retains compiler identity, command, environment, exact source
archive and before/after hashes. The reused matrix capture runner is frozen in
each archive; its descriptive name does not restrict the selected test targets.
All archive members and log hashes were verified before captures were moved
out of scratch. The single existing disk-backed build workspace is retained;
no temporary repository copy was created under /tmp or /var/tmp.

## Candidate 09 validation

[Focused capture](candidate-09/receipt.json) passes 30 tests: 10 conformance,
12 resource/ownership/error tests, and 8 array/reference tests. The exact
464-file dependency source closure and frozen capture/hash runners are in
`candidate-09/source.tar.gz`. [Full gates](candidate-09/gates.json) match the
canonical source before and after: 1,105 tests pass, Clippy passes with warnings
denied, documentation passes with warnings denied, five doctests pass, and
formatting passes. Full logs and the gate runner are in
`candidate-09/gate-logs.tar.gz`.

The final source incorporates independent reviews of public/owned values,
streamed references, reservation lifetimes, numerical extremes, and scalar
aggregate cancellation/error ordering. Tests include finite exponential and
hyperbolic outputs whose naive scale overflows, mixed division preserving a
subnormal denominator component, product cancellation, and signed-zero axes.
Earlier candidate diagnostics retain compile errors and corrected test-oracle
mistakes; they are not final gate results.

[Performance report](performance/report.md) retains 225 successful new-family
process runs and an existing-value comparison of 162 processes. Six RSS flags
and one small allocation-byte change remain explicit review items; passing
correctness checks is not broad performance acceptance. Database functions and
other missing audit work remain outside this implemented family.
