# ODS trigonometry performance plan

This profile is a paired, bounded process benchmark for the ODS OpenFormula
trigonometric scalar/value evaluator.  The baseline is the committed
`069c65870` snapshot.  The candidate is frozen only after the implementation,
tests, review, and root validation have finished.  The candidate may evaluate
the new trigonometric family; the baseline is expected to return a typed
unsupported result for those cases.  Paired ROUND and arithmetic control lanes
are therefore the evidence for dispatch or evaluator regression.  A baseline
trigonometric refusal is retained as capability evidence and is never treated
as a numeric speedup comparison.

The final capture will use the same harness, fixtures, release profile, Rust
toolchain, and lockfile in fresh processes.  Each measured case/phase has
three warmup evaluations inside each child and fifteen measured fresh child
processes.  The baseline runs the ten matched arithmetic/ROUND controls; the
candidate runs all fifty named cases.  The timed process repeats a fixed
operation count so short scalar calls are measurable.  It records elapsed
time, allocator calls and requested/released bytes, live-byte balance, peak
live allocation delta, execution work and result-live budget memory, and
maximum RSS from `/usr/bin/time -v`.  The binary runs one direct `f64`
numerical and shape oracle, independent of evaluator dispatch, before
resetting counters and timing.  This oracle checks evaluator projection
consistency; it is not a separate libm-accuracy implementation.

Cases cover:

- scalar trigonometric, inverse, hyperbolic, and degree/radian calls;
- literal 4x4 and 16x16 arrays through the value evaluator, using five
  representative unary trigonometric kernels (`SIN`, `COS`, `ASIN`, `ACOSH`,
  and `TANH`) alongside the controls;
- scalar and rectangular local-reference values supplied by a synthetic
  immutable resolver (not the worksheet adapter), using representative
  `SIN`, `ACOSH`, and `ATANH` kernels;
- matched existing ROUND and arithmetic scalar/array/reference controls.

The resolver is an immutable in-memory fixture.  It performs no I/O, formula
recalculation, cache publication, workbook mutation, networking, or ambient
thread scheduling.  The profile measures evaluator entry points and fixture
reads only; it does not claim native LibreOffice/Office acceptance, complete
OpenFormula conformance, full-workbook recalculation, save performance, or
package-level RSS behavior.

Capture custody:

1. Freeze and hash this directory's harness, runner, verifier, fixtures and
   lockfile before either paired capture.
2. Build and capture the ten-control baseline from commit `069c65870` in a
   temporary detached worktree and external Cargo target.
3. Build and capture all named candidate cases from the root agent's exact
   selected-file freeze with the same inputs.
4. Run the fail-closed verifier and retain raw JSON, stdout, stderr and
   `/usr/bin/time -v` receipts.  Do not modify profile inputs after capture.
