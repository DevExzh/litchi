# OpenFormula bitwise evaluation

This batch implements all five OpenDocument 1.4 Part 4 §6.6 functions in the
existing explicit scalar evaluator: `BITAND`, `BITOR`, `BITXOR`, `BITLSHIFT`, and
`BITRSHIFT`. It advances the full OpenFormula evaluation gap in §10 of
[the audit](../../spec-gap-audit.md); reference/array resolution, other function
families, workbook recalculation and cache publication remain open.

## API and deterministic profile

Call `codec::formula::evaluation::{evaluate_scalar, evaluate_scalar_with_context}`
on an already-parsed immutable `Expression`. Existing caller-owned cancellation,
work, stack and storage limits apply. The operation has no workbook, cache, I/O,
ambient provider, new dependency, or new public settings.

Integer parameters first use the existing scalar Number conversion, then truncate
toward zero. This rounding is an explicit implementation choice under §6.3.6,
not a universal OpenFormula rule. Data operands and results use the inclusive
unsigned range `0..=281474976710655` (48 bits). Values outside that domain produce
a formula Number error. Domain checks happen after conversion: `-0.9` truncates
to zero and is accepted, while `-1.9` becomes a negative integer and is rejected.
Numeric text and Logical values use the existing Number conversion policy.

Shift counts are signed finite integers after conversion, with no arbitrary
count cap. Negative counts reverse direction. Right shifts of 48 or more return
zero. Left shifts of zero return zero for every finite count; otherwise an
unrepresentable result produces a Number error. The implementation checks the
48-bit result bound before shifting, since Rust's `checked_shl` validates only
the count and does not detect high bits shifted out of the integer. No loop is
proportional to shift magnitude, and no intermediate shift can discard bits and
then pass a domain check.

Arguments are evaluated eagerly in source order, including arguments of a
malformed-arity call. The scalar error policy retains the leftmost formula error;
a wrong arity without such an error returns a Value error. References, arrays,
unimplemented functions, cancellation, and resource/allocation failures retain
typed Rust refusals. `IFERROR` can catch bitwise formula errors; it cannot catch
those Rust failures. A surrounding lazy `IF` still leaves an unselected bitwise
branch unevaluated.

## Validation results

The final source passes **903 tests across 54 Cargo targets**, including nine
focused bitwise integration groups, and **2 executed doctests**. All-feature
all-target testing, warning-denied Clippy and rustdoc, formatting, and crate
boundary checks pass. Gate receipts bind 212 exact source inputs before/after.

The common corpus contains 117 rows per revision, and the candidate-only
bitwise corpus contains 93 rows. The first measured candidate had repeated 8–11%
text-evaluation regressions. Keeping `apply_bitwise` out of inlined evaluator
code removed those regressions in fresh captures and four interleaved repeats.
The final repeat covers 46 flagged/carried cases (368 rows), retaining the four
previously regressed lanes even when they no longer cross the screening limit.
Four parse-only lanes remain 5–8% slower in median latency; no final repeated
evaluation or parse/evaluate p50 delta, or peak-RSS delta, exceeds the 5%
review threshold.
Tracked common-case allocations and peak live heap remain identical. These
results are scoped to the retained corpus and binaries, not a general speedup.

## Production boundaries and evidence

The implementation reuses existing bounded evaluator frames and fallible stacks.
It introduces no unsafe code or changes to execution, reservation lifetime, package
preservation, or public API ownership. ADRs 0002/0004/0023/0024 govern API and crate
ownership; ADRs 0005/0006 govern resources, performance evidence and preservation.
Source parsing and the existing bounded numeric converter retain their existing
cooperative-cancellation boundaries.

[Specification review](spec-review.md), [implementation review](implementation-review.md),
[gate receipts](gates/results.json), and [performance report](performance/report.md)
record the final validation and its limits. `candidate.patch` is the exact change
from baseline `c32455df83ea67f840a6d4f8da8953f113b590f4`. The retained workspace
Cargo lock is an ignored build input, explicitly saved as `gates/workspace-Cargo.lock`.

The standalone measurement harness lives at `performance/bitwise-harness`.
Its four files are identical for baseline and candidate. Build in a fresh
baseline checkout using the retained build-command receipt and lockfile, save
the executable, then apply `candidate.patch` and rebuild in a fresh target
directory to avoid stale fingerprints. Run the harness runner with a fresh output
directory, `--group comparable` for both revisions and `--group bitwise` for the
candidate; use the captured CPU, phase, warmup, iteration and repeat settings.
The orchestration script records its original scratch paths and must be adapted
to fresh output directories when repeating measurements. Existing evidence should
not be overwritten. Baseline refusals are not speedup baselines for new functions.

Run `python3 docs/report/spec-gap-validation-evidence/ods-formula-bitwise-evaluation/verify.py`
to verify retained source, patch, gate, measurement and cleanup evidence without
rebuilding. Build directories and saved binaries are removed after validation;
provenance receipts retain their hashes and sizes.

The single disk-backed scratch checkout, target, and TMP directory were removed
after their processes finished, reclaiming 1,546,616,832 allocated bytes. The
22 unique scratch files were hash-verified in a recovery archive outside temporary
storage. Three saved measurement executables were reverified and removed after
their captures; source, binary, build, and raw measurement receipts remain. See
`gates/cleanup.json` and `gates/root-binary-verification.json`. This batch did not
create repo or build duplicates in `/tmp` or `/var/tmp`.
