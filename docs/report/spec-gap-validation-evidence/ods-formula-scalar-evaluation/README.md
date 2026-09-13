# Explicit OpenFormula scalar evaluation

This batch addresses the missing execution layer in §10 of
[`spec-gap-audit.md`](../../spec-gap-audit.md). The new public
`litchi_ods::codec::formula::evaluation` module evaluates the immutable ODS
expression tree directly. It implements scalar literals, typed formula errors,
all scalar operators, and `TRUE()`/`FALSE()`. This is a foundation for the
remaining runtime work, not an ODF Small Group or full evaluator claim.

## Behavior and API

`evaluate_scalar_with_context(&expression, &execution, &limits)` uses a
caller-supplied `litchi_core::ExecutionContext`. The canonical contextual
entry point is `evaluate_scalar(&expression, &EvaluationContext, &limits)`.
Both return `EvaluatedScalar`, whose `value()` exposes distinct Number,
Logical, Text, and Error values. Formula errors remain values; cancellation,
allocation/resource failures, and unsupported capabilities are typed Rust
failures. Unescaped text borrows the immutable expression. Owned text keeps
its memory reservation until the result is dropped.

The profile uses finite f64 numbers, exact numeric equality, case-sensitive
Unicode scalar ordering without normalization, locale-independent decimal
text conversion, a `0^0 = 1` choice, and a Value error for mixed-type ordered
comparisons. Prefix plus preserves scalar identity. Power is left-associative
and prefix minus binds before power, as represented by the ODF tree. Unknown
error literals become the Name error. Supplied arguments of supported boolean
functions are evaluated before the arity error, preserving input errors and
capability/resource refusals.

Opening or parsing files remains inert. This operation does not consult a
workbook, prefer stored cached results, mutate a cache, resolve a name or
reference, access an external source, or publish a cell/package change.
References, arrays, names, labels, missing arguments, and other functions
return typed unsupported results. Function recognition in the separate
393-name catalog does not imply evaluation support.

## Resource and architecture contract

The iterative frame/value stacks avoid recursive execution and destruction
for long flat expressions. Default local limits are 1,000,000 work units,
32,767 UTF-8 bytes per text value, 4 MiB aggregate requested heap capacity, and
65,536 entries per stack. Work covers AST visits/applications and admitted
numeric/text scans and copies. String scan/copy and comparison loops check
cancellation in bounded chunks; the caller controls cancellation and shared
ancestor budgets.

An operation-local child budget charges live stack and owned-text reservations
against both the local storage ceiling and the caller's ancestors. Growth is
fallible, and failed evaluations release temporary reservations. Owned-left
concatenation grows geometrically and charges existing bytes only when growth
requires copying them. Borrowed AST storage, fixed stack scratch, caller-made
copies, allocator overhead, and existing core budget/error bookkeeping are
outside the requested heap-capacity accounting; this is not a whole-process
allocation-failure or RSS guarantee. Cooperative cancellation does not promise
hard real-time interruption of allocator or numeric-conversion calls.

The ODS grammar/value/conversion semantics stay in their format owner. The
existing Excel-oriented `litchi-eval` parser and coercions are not reused.
There are no new dependencies, crate edges, executors, or ambient providers.
This follows ADRs 0002/0023/0024 for ownership, 0004 for typed public operations,
and 0005/0006 for bounded explicit execution and inert preservation.

## Validation

The frozen candidate passes **879 tests across 52 Cargo targets**, including
**24 independent scalar-evaluation integration groups**, plus **2 compiled
and executed doctests**. The existing filtered subprocess replay is recorded
separately and is not double-counted as a target. All-feature/all-target ODS
tests, warning-denied Clippy and rustdoc, doctests, formatting, and the current
workspace dependency-boundary check passed. Exact commands, logs, and 210
source-input hashes are retained under `gates/`.

The independent tests cover normative operators, error propagation versus Rust
failures, text and Unicode behavior, large/small finite formatting, decoded
literal limits, aggregate and ancestor memory admission, output reservation
lifetime, pre-cancellation, work/stack limits, eager boolean arguments,
unsupported capabilities, escaped literals, and long arithmetic/concatenation
chains. Source inspection additionally checks cooperative loop checkpoints;
there is no hard latency or scheduler-dependent mid-loop cancellation claim.

See [`spec-review.md`](spec-review.md),
[`implementation-review.md`](implementation-review.md), and
[`scalar-cases.json`](scalar-cases.json) for normative and review evidence.
The independent case corpus distinguishes mandatory outcomes from host or
implementation choices; its initial overly strict 0^0 expectation was corrected
against Part 4 §6.16.46 before the implementation was accepted.

Performance phases and raw data are described in
[`performance/report.md`](performance/report.md). An isolated ASCII-name parser
optimization was **reverted** after repeatable reference regressions; its exact
patch, baseline/candidate/ABAB measurements, and disposition remain in
[`performance/ascii-experiment.md`](performance/ascii-experiment.md). No parser
speedup from that experiment is shipped or claimed.

`candidate.patch` replays exactly from `9072e969e`; `verify.py` checks the
retained specification entry, source hashes, gate commands/results, patch
replay, profiling provenance, and artifact manifest. Run it from any working
directory with `python3 path/to/ods-formula-scalar-evaluation/verify.py`.

## Remaining audit requirements

Function-family execution, reference/name resolution and implied intersection,
array evaluation, host-dependent case/locale/date/display policies, dependency
graphs, volatile recalculation, external providers, and explicit cache
publication remain incomplete. This batch does not close those audit rows or
the broader thread goal. Profiling here concerns parser/scalar calls; it does
not establish workbook-level calculation or end-to-end document CRUD speedups.
