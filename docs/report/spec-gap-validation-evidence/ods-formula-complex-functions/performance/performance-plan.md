# Complex-function performance plan

This diagnostic harness measures the bounded value-evaluator path for the
complete OpenFormula §6.8 family. The normative behavior and explicit local
choices are recorded in the [complex-function contract](../contract.md). The
current catalog recognizes these names, while the evaluator implementation is
being completed separately; an unsupported baseline is therefore useful as a
capability reference but cannot supply a successful before/after performance
comparison until both revisions implement the same result contract.

The base corpus has one independent case for each of the 26 functions:
`COMPLEX`, `IMABS`, `IMAGINARY`, `IMARGUMENT`, `IMCONJUGATE`, `IMCOS`,
`IMCOSH`, `IMCOT`, `IMCSC`, `IMCSCH`, `IMDIV`, `IMEXP`, `IMLN`, `IMLOG10`,
`IMLOG2`, `IMPOWER`, `IMPRODUCT`, `IMREAL`, `IMSIN`, `IMSINH`, `IMSEC`,
`IMSECH`, `IMSQRT`, `IMSUB`, `IMSUM`, and `IMTAN`. Inputs use real-only,
imaginary, `i`, and `j` forms where applicable. The independent oracle uses
fixed `f64` pairs and principal branches. Complex results are validated before
timing as the public `Value::Complex` variant, checking both `real()` and
`imaginary()` plus the preserved `suffix()` (`i` or `j`); the original
expression is also wrapped in `IMREAL` and `IMAGINARY` for independent
component checks. Numeric-returning functions are checked directly as numbers.
Timed checksums fold both complex components and the suffix.

Additional lanes cover nested complex operations, a scaled finite `IMDIV`
(`1e308+1e308i` divided by `1+i`), `IMSUM` and `IMPRODUCT`
with 1/4/16/64/256/1024 operands, and read-only reference sequences at 16/256/
1024 cells. The in-memory resolver returns static complex text, plus explicit
Empty, Logical, and formula-error cells, allowing sequence omission and error
precedence to be checked without I/O. Selected and unselected `IF` branches,
`IFERROR`, and `IFNA` exercise lazy behavior; the unselected branch includes a
large reference range. Invalid text, suffix, arity, zero/domain, divide-by-
zero, overflow, ignored `IMSUM` text, zero-arity policy, and 4 KiB text
cases exercise conversion and formula-error boundaries. Over-limit text and separate
work, storage, and cancellation lanes expect typed evaluation failures.
The reference corpus reads Text, Empty, Logical, and Error cells; numeric
reference-cell conversion is covered by integration tests, not this harness.

`evaluate` reuses a parsed expression and constructs a fresh bounded resolver
and execution context outside each timed sample. The timed region consists of
evaluation, checksum folding, and dropping the result. Warmups and measured
iterations are configurable; the default is three warmups and fifteen
iterations, with a bounded repeat chosen from the case size. Allocation calls,
requested/released bytes, live and peak deltas, retained execution-budget
memory, work, and resolver calls are emitted per JSON line. `memory_retained`
is a budget reservation observed while the result is live, not allocator peak;
RSS must be supplied by an external `/usr/bin/time -v` wrapper. Raw component
oracle work, parsing, resolver construction, and fixture generation stay out
of the evaluate timer.

Capture candidates serially on one pinned CPU with fresh processes and the
same harness bytes, binary flags, limits, and corpus. Preserve successful
result fields exactly across revisions and review every individual lane under
the approximately 5% latency/RSS trigger in `docs/GOAL.md`. A single capture
is diagnostic evidence: repeated AB/BA rounds are required before attributing
latency, allocation, or resident-memory changes, and no aggregate or broad
throughput claim follows from this plan.
