# 0442: Share ODP auxiliary staging traversal

Starting an ODP transaction previously tokenized the same content independently
for settings, declarations and page metadata. Those three state machines now
consume one namespace-aware borrowed-event stream. Individual public parsers
remain available. The grouped traversal preserves full settings-then-declarations-
then-pages error priority by deferring lower-priority errors and finishing in
historical order. Source fragments, semantic slides and publication readback
retain their existing independent validation paths.

The frozen A1/B1/B2/A2 experiment retains 24 reports, 720 samples and four
profiles across 64/4,096/8,192 source slides. Medium normal p50 improves
12.891%/9.830%, large 13.307%/10.220%, and tiny 7.020%/6.923% (R1/R2).
The predeclared 5% medium/large normal gate passes. Each size saves only
38 allocation calls and 11,902 requested bytes; peak and retained bytes are
unchanged, and the practical allocation/peak gate fails.

No matched adverse metric exceeds 5%. All five repeat flags remain: four
baseline allocator medium timing changes and candidate large normal p99 +7.091%.
The paired large R2 p99 still falls 2.268%, but no general tail-speedup claim
is made. Whole-process cycles/instructions fall 10.092%/8.824%, including setup,
warmups and Rust oracle work. These are not operation-only causal fractions.
See [measurements](../results/change-0442/measurements.md) and the
[keep decision](../results/change-0442/decision.json).

Three test-only original parser bodies are retained byte-for-byte. Differential
tests cover 353 input comparisons plus explicit error-priority assertions,
including namespace changes, malformed XML/attributes and limit edges. The
[bundle](../results/change-0442/README.md) records validation, source custody,
all attempts and portable evidence verification. No public API or dependency
changes; lower-priority outputs may accumulate before higher-priority failures
within existing finite limits. Invalid-input performance is unmeasured.

The registry remains 436 selectors and the default matrix 36 cases. This
optimization adds no CRUD coverage promotion. Remaining scans, one-shot caching,
bounded existing append, Part addition, repackaging, native breadth, cold/range
and scaling remain open under the unchanged non-iWork goal.

Final validation passes 355 ODP tests, 368 harness tests (one existing ignored),
strict owner Clippy, workspace/ODF feature checks, warning-denied rustdoc,
formatting and crate boundaries. Portable controls and 11/12 corruption probes
pass before/after cleanup. Four temporary executables totaling 1.83 GB were
removed; build caches and user-owned GOAL.md were preserved.
