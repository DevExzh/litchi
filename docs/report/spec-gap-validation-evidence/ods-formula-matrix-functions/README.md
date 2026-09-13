# OpenFormula matrix-function evidence

The intended implementation covers all five ODF 1.4 Part 4 §6.5 functions:
`MDETERM`, `MINVERSE`, `MMULT`, `MUNIT`, and `TRANSPOSE` in the bounded value
evaluator. This family remains in progress.

- [Contract](contract.md): local specification sections and context rules.
- [Implementation profile](implementation-profile.md): explicit choices for
  element conversion, numerical errors, and host limits.
- [Performance plan](performance-plan.md): corpus, accounting, and acceptance.
- [Capture runner](capture.py): focused test execution with archived source
  inputs, toolchain, command, environment, and before/after hashes.

## Baseline 01

The isolated workspace matched the 455 canonical source inputs at `1f2029994`
before the new test file was copied into it. With that file, the capture has
456 source inputs. All 12 new tests compiled and failed because the existing
value evaluator returned `Unsupported(Function)`. This establishes an
implementation gap; it is not a passing validation gate or a performance
baseline.

The [receipt](baseline-01/receipt.json), [test log](baseline-01/test.log), and
[source archive](baseline-01/source.tar.gz) retain exact evidence. The receipt's
Git head belongs to the sparse build workspace and is contextual only; the
archived source bytes and source hashes identify what ran. The test source
may gain further coverage after this capture; the archived version remains
authoritative for these 12 failures.

The capture used the existing disk-backed build workspace, target cache, and
scratch directory. The temporary capture directory was moved into this
evidence directory after its archive and receipt were verified. No repository
or build copy was created under `/tmp` or `/var/tmp`.

## Expanded baseline 04

The [expanded receipt](baseline-04/receipt.json), [log](baseline-04/test.log),
and [archive](baseline-04/source.tar.gz) cover 457 inputs and 28 tests: 21
matrix conformance tests and seven resource/ownership tests. Two existing
boundary checks pass (zero work/storage admission and pre-cancellation),
while 26 tests fail. These include computed identity dimensions, nested
ForceArray/lazy state restoration, mixed-type transpose ownership, and
reference fallback after a singular inverse. Matrix implementation is still
required; the passing boundary checks do not establish bounded matrix
arithmetic.

Intermediate captures in `diagnostics/test-api-draft-02` and
`diagnostics/test-import-draft-03` are test-source compilation failures, not
production regressions. Root corrected a nonexistent test-builder method,
shadowed helper names, and an unused import before baseline 04. Each retains
its exact archived source, compiler log, and receipt. All temporary capture
directories were moved here; no loose source copies remain in scratch.
