# ODS value-inspection batch disposition

Status: source, correctness gates and performance evidence accepted; owned
temporary checkout and build-target cleanup complete.

The batch implements sixteen OpenFormula §6.13 functions through shared private
kernels and the scalar/value evaluator entry points. The
[contract](contract.md) defines each function's raw-value, conversion, error,
reference and matrix behavior. It leaves host/reference metadata functions and
configurable locale/date contexts outside this batch.

The implementation preserves Empty before coercion and distinguishes omitted
optional separators from explicitly empty Text. TYPE consumes complete Any
arguments, scans reference cells without retaining them, and returns the array
type for computed arrays. N preserves scalar reference intersection and selects
the first element of an array. Ordinary inspection functions stream matrix
reference coordinates. Computed arguments use the enclosing calculation mode;
explicit scalar descendants retain the current position. Formula errors remain
data for raw inspectors, while provider, cancellation, resource and source
failures remain typed evaluation failures.

VALUE provides the fixed en_US numeric, currency, percent, fraction, date, time
and datetime profile. Shared Gregorian helpers keep TEXT and VALUE on the same
calendar. Long grouped numbers use budgeted scratch and exact decimal
conversion. NUMBERVALUE normalizes explicit separators in the specified order,
with linear group removal and periodic execution checks. This adds no public
options, ambient services, recalculation or cache publication.

## Correctness and source custody

- Independent [semantic review](semantic-review.md): PASS.
- Independent [resource/cache review](resource-review.md): PASS.
- Seven isolated gates: PASS, including 1,623 ODS tests with no failed or
  ignored tests, strict Clippy, rustdoc, formatting, boundary and diff checks.
- Independent oracle: 107 observations across all sixteen functions; the
  retained Python model and Rust consumer agree.
- Native fixture: 123 rows, with 49 exact matches, 46 numeric-tolerance
  matches, 11 error-kind matches and 17 documented profile differences.
- The 82-file freeze and complete workspace manifests match the isolated
  checkout before and after the accepted gate run. The baseline is
  `d623f3c2ecc0c837017f700174656f0e443759a5` and the retained gate lock is
  `58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`.
  The ambient workspace lock is a separate input and remains unchanged.

Regressions cover signed negative exponents, multi-digit fractional seconds,
near-minute rounding, omitted separators, scalar versus matrix TYPE results,
computed arrays in projected branches, MUNIT scalar descendants, 3-D plane
selection, zero-read refusals and typed later-read failures.

## Performance and cleanup

The frozen workload contains 87 cases: 28 matched controls and 59 new-function
cases, measured in evaluate and parse-evaluate phases. Release preflight checks
exact values, shapes, reads and typed refusals for every case. Two complete
captures independently verify, totaling 6,900 samples. Their candidate source,
harness, profile inputs and executable hashes agree. The second capture arose
from a handoff visibility error, not a decision to select better measurements.
Both complete captures remain part of this disposition.

Across the same 56 matched control groups in each capture, allocator calls,
requested/released bytes, peak live bytes, retained result budget, work,
reference reads and output bytes remain unchanged. The earlier capture's
latency shifts range from −3.46% to +7.72%; the later capture ranges from
−3.05% to +3.56%. No general runtime speedup or universal non-regression claim
is made.

The earlier SUMIFS parse-evaluate lane increased by 7.72%, approximately
6.3 microseconds per repeat. Its independently bootstrapped 95% interval is
−1.96% to +10.91%. This remains an accepted review flag: the samples do not
establish its direction reliably, but the observation is not discarded or
declared cleared by the later capture. The new functions and shared dispatch
pass the source, resource and correctness reviews; no speculative rewrite was
introduced to chase this one uncertain microbenchmark result.

RSS shifts across both captures range from −4.67% to +6.67%. One earlier
comparison exceeds +5% by 172 KiB; eight later comparisons exceed it by
168–212 KiB. These bounded process-footprint increases are explicitly accepted
for this function batch. Evaluator allocation and budget metrics are unchanged
across every matched sample group. Those facts do not identify the cause of
the process RSS changes; the report makes no such attribution.

The [independent root audit](root-performance-audit.json) records every matched
comparison, both sets of flags and reproducible uncertainty estimates. Its
[script](root_performance_audit.py) checks raw receipts, source manifests,
preflights, exact accounting sets and target-cleanup receipts before computing
statistics. Earlier incomplete harness attempts remain diagnostic evidence and
are excluded from the equivalent complete-capture comparison.

Capture-owned worktrees and targets have been removed; their receipts remain
with the measurements. The owned gate target, frozen gate
checkout and active-run registry are also removed; see the
[final cleanup receipt](gates/final-cleanup.json). Verification runs entirely
from retained evidence and Git source objects. Raw native and performance
evidence remains in the batch.
