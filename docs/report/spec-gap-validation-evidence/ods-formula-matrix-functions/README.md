# OpenFormula matrix-function evidence

The intended implementation covers all five ODF 1.4 Part 4 §6.5 functions:
`MDETERM`, `MINVERSE`, `MMULT`, `MUNIT`, and `TRANSPOSE` in the bounded value
evaluator. All five functions are implemented in candidate 06; broader
OpenFormula coverage and performance acceptance remain open.

## Current result

[Candidate 06](candidate-06/receipt.json) passes 34 focused tests and all five
[ODS gates](candidate-06/gates.json): all-feature/all-target tests, Clippy with
warnings denied, documentation, doctests, and formatting. The 458-file source
closure is unchanged across the checks. The [source archive](candidate-06/source.tar.gz)
and [gate logs](candidate-06/gate-logs.tar.gz) identify the exact implementation.

The implementation adds iterative argument-context continuations, isolated
shape-planner state, bounded numerical kernels, and reuse of literal-only
matrix branches. Nested shape/value probes have a checked Depth ceiling;
resource and cancellation errors remain distinct from formula errors.
`MUNIT` uses the first parameter cell in matrix mode, including selected lazy
branches and offset caller positions. Its shape discovery follows that same
conversion. `TRANSPOSE` preserves matrix calculation context inside a lazy
branch while retaining its non-ForceArray signature.

The release harness has independent numeric/shape/error oracles, with full
validation outside the timed loop. Release measurements are initial absolute
baselines; they do not close existing-workload regression or end-to-end
performance requirements.

The [current release report](performance/release-06-analysis.md) retains 99
successful processes over 33 cases, with zero deterministic result/counter
changes between the two release windows. Ten cases trigger RSS review flags in those sequential windows. The
[paired follow-up](performance/pairs-01-analysis.md) passes 198 processes with
unchanged deterministic results and counters. Fourteen cases exceed the RSS
review threshold, including five of the earlier ten; no median p50/p95 latency
change exceeds 5%. RSS acceptance and existing-workload checks remain open.

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

## Candidate history

- Draft 01 did not compile; the kernel's nested result handling and a nested
  mutable borrow were corrected.
- Draft 02 passed 28 of 30 tests. Root fixed shape-planner state reentry. The
  second failure was a fixture advertising numeric cells it actually supplied
  as Empty; its rectangular storage was corrected without weakening numeric
  element rules.
- Candidate 03 passed 30 tests and all five gates. Draft 04 added bounded
  probe depth and passed 31 tests. Candidate 05 added context and work-reuse
  coverage and passed 33 tests plus all five gates. Candidate 06 adds the
  first-parameter-cell rules described above.
- The first runtime harness smoke stopped on an incorrect inverse oracle:
  for the fixture `I + J`, the diagonal of the inverse is `n/(n+1)`, not
  `2/(n+1)`. The corrected oracle passes all 33 runtime cases. This was a
  harness defect, not an inverse-kernel defect.

Historical captures retain their exact sources and logs. The current source
archive and gate receipt, rather than earlier counts, establish current
validation. Temporary capture directories and loose gate logs were removed
only after archived contents were verified.

## Existing-workload comparison

The [common-value comparison](performance/common-01-analysis.md) against the
pre-matrix evaluator passes 162 processes with no deterministic or memory-counter
changes. Ten cases cross a latency or RSS review threshold, including text
arithmetic p50 +6.57%. These findings remain open; current matrix support does
not establish a blanket performance pass for existing workloads.
