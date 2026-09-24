# OpenFormula logical evaluation

This batch continues the explicit ODS evaluation work in §10 of
[`spec-gap-audit.md`](../../spec-gap-audit.md). The scalar evaluator gains the
nine functions in OpenDocument 1.4 Part 4 §6.15: `AND`, `FALSE`, `IF`, `IFERROR`,
`IFNA`, `NOT`, `OR`, `TRUE`, and `XOR`.

The public entry points remain
`codec::formula::evaluation::{evaluate_scalar, evaluate_scalar_with_context}`.
They consume the immutable expression tree and caller-owned execution context.
Formula errors remain scalar values. Cancellation, resource/allocation failures,
and unsupported capabilities remain typed Rust failures.

## Conditional execution

`IF` evaluates its condition and selected branch. `IFERROR` evaluates its first
argument once and evaluates its alternative only for a formula error. `IFNA`
selects its alternative only for `#N/A`. An unselected reference or unsupported
function is not resolved or evaluated; parsing still validates the entire input
and applies the expression parser's byte, node, and depth limits.

Omitted `IF` arguments follow §6.15.4: `IF(condition)` returns a Logical value;
a two-argument call defaults the false result to `FALSE`; an explicitly empty
true or false slot produces numeric zero when selected. This distinguishes an
absent optional argument from a semicolon-delimited empty slot. Missing required
handler slots produce a Value error only when evaluated: `IFERROR(1;)` returns
1, while `IFERROR(#N/A;)` returns Value. `IFERROR(;2)` catches the missing
first slot's Value error; `IFNA(;2)` retains it. Malformed conditional arities
return Value before branch scheduling. These are explicit bounded profile
choices for calls outside the functions' ordinary signatures.

Boolean aggregation is eager. A decisive false `AND` argument or true `OR`
argument does not bypass later arguments' errors, resource limits, or unsupported
capabilities. Logical conversion maps zero to false and nonzero finite numbers
to true. The deterministic scalar profile refuses direct Text-to-Logical
conversion with a Value error. `AND` and `OR` instead accept numeric text through
their NumberSequenceList parameter conversion, using the existing
locale-independent decimal policy.

## Boundaries and remaining work

The existing limits and accounting exclusions remain documented in the
[scalar profile](../ods-formula-scalar-evaluation/README.md): default limits are
1,000,000 work units, 32,767 bytes per text value, 4 MiB requested evaluator
storage, and 65,536 entries per stack. Numeric-conversion and allocator calls
remain indivisible cooperative-cancellation boundaries.

Selected output storage retains the existing reservation lifetime. Execution
uses bounded explicit stacks and caller cancellation; it does not mutate caches
or packages, introduce ambient providers, or recursively execute nested branches.
This retains the format ownership and typed-operation contracts in ADRs
0002/0004/0023/0024 and the resource and preservation contracts in ADRs 0005/0006.

This is scalar-domain support for the logical family. Reference and array
aggregation, workbook/name resolution, other function families, dependency-graph
recalculation, host locale policies, and explicit cache publication remain open.
It does not claim a complete OpenFormula evaluator or workbook CRUD speedup.

## Validation

The frozen candidate passes **894 tests across 53 Cargo targets**, including
**15 logical-function integration groups**, and **2 executed doctests**.
All-feature/all-target ODS tests, warning-denied Clippy and rustdoc, formatting,
and the workspace boundary checker pass. The filtered depth-check subprocess is
recorded separately and not counted as another Cargo target. The gate receipt
records 211 exact source-input hashes before and after execution.

Independent tests distinguish formula errors from evaluator failures, exercise
every optional IF form, prove eager aggregate and lazy branch behavior, bound
once-only error-handler input evaluation, check cancellation and resource
failures, and retain/release selected owned text. The final doctest demonstrates
`IF(FALSE();[.A1];IFERROR(1/0;42))` returning numeric 42 without resolving the
unselected reference. Initial test expectations for variadic XOR and the omitted
false IF arm were corrected against the normative signatures and defaults.

See [spec-review.md](spec-review.md) and
[implementation-review.md](implementation-review.md) for review evidence,
[gates/results.json](gates/results.json) for exact commands and source hashes,
and [performance/report.md](performance/report.md) for measurement scope and
per-scenario results. `candidate.patch` replays against `3f907e9e1`; `verify.py`
checks retained source, gate, specification, profile, and artifact evidence.


The performance review retains every flagged scenario and four interleaved
baseline/candidate pairs per final flag. The large numeric-coercion case
(`evaluate/control-coerce-4096`) measured +3.20% median latency after the dispatch
change; `evaluate/control-flat-1024` still measured +5.19%. Some process-RSS
medians remained 6–9% higher, while comparable allocation counts, requested
bytes, and peak live tracked heap stayed identical. These are scoped costs and
uncertainties of the added capability, not a claim of universal speedup.

## Reproduction and cleanup

Run `python3 docs/report/spec-gap-validation-evidence/ods-formula-logical-evaluation/verify.py`
from the repository to verify retained evidence without compiling or recreating
temporary directories. The workspace Cargo lock is an ignored input; its exact
bytes are retained in `gates/workspace-Cargo.lock`.

To repeat measurements, use a fresh detached checkout of baseline
`3f907e9e1abcde99b2ce94e88c1122bb79f77a8d` on disk, excluding bulky report and
performance trees during checkout. Copy this evidence directory's
`performance/logical-harness` into the same relative location in that checkout,
and copy `gates/workspace-Cargo.lock` to its root `Cargo.lock`. Build the harness
with the exact command in `performance/baseline/build-command.json`, assigning
`CARGO_TARGET_DIR` and `TMPDIR` to empty directories on disk. Save the baseline
executable outside the target directory. Apply `candidate.patch` at the checkout
root, then use a fresh target directory for the candidate build to avoid stale
mtime or fingerprint reuse. For the intermediate five-frame implementation,
apply `performance/before-dispatch/candidate.patch` to a separate baseline
checkout instead. The four harness files are identical across all builds.

Invoke `performance/logical-harness/run.py --help` for output-directory and phase
options. Use `--revision baseline --group comparable` for the baseline; use
`--revision candidate` with each of `--group comparable` and `--group logical`
for the candidate, supplying `--binary`, a fresh `--output`, `--phase all`,
`--warmups 3`, and `--iterations 15`. The runner pins CPU 2; a repeat on another
machine must record its CPU selection and environment. Exact per-case commands,
repetition counts, status, stdout, RSS, and counters are retained beside each
measurement. Counter commands include setup and warmup work outside the timed
evaluation interval. The capture scripts preserve the original orchestration
paths and should be adapted to fresh output paths rather than overwrite evidence.

The original scratch checkout, target, and TMP directory were removed after
all build/profile processes finished, reclaiming 1,551,687,680 allocated bytes.
Unique scratch files were hash-verified in a recovery archive outside temporary
storage; the retained main evidence is sufficient for verification. See
`gates/cleanup.json`. Saved executables are removed after hash and measurement
verification; their provenance receipts remain. No temporary repo copies are
required by the retained verifier.
