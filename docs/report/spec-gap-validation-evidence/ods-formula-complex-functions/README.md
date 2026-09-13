# OpenFormula complex functions

The active implementation targets all 26 ODF 1.4 Part 4 §6.8 functions in the
scalar and value evaluators. [Contract](contract.md) records exact local source
hashes, conversion and sequence rules, selected interpretations of conflicting
IMSQRT/IMSECH text, and independently computed extreme-value oracles.
Implementation and final validation remain in progress.

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

## Remaining work

Complete and review the fixed-size complex representation, scalar kernels,
array/reference sequences, owned conversion, and ordinary scalar coercions.
Then run focused and full ODS checks, execute the independent performance
harness, and compare existing workloads. Passing the baseline's three existing
resource checks does not establish bounded complex execution or support.
