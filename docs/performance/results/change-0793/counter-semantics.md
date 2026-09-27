# Post-capture counter-semantics correction

The frozen plan is retained unchanged and its owner-count reconciliation fails.
This note was added after all ten traces and twenty decodes completed. It is a
source explanation and supplementary diagnostic, not a replacement gate.

`probe-src/src/allocation_metrics.rs`, `Counters::reallocation_locked`, increments
both `allocation_calls` and `reallocation_calls` for each successful realloc.
The latter is a subset counter. Consequently the frozen formula
`allocation_calls + reallocation_calls` counts each successful realloc twice.
The appropriate supplementary comparison is owner stack cost against
`allocation_calls` alone. Failed allocations are recorded separately.

For the first large trace, the frozen expectation is 72,106 + 2,603 = 74,709;
the observed wrapper-owned allocation cost is 72,106. This does not pass the
frozen rule. The supplementary equality explains the difference but does not
retroactively qualify the experiment or authorize owner allocation fractions.

Filtered heaptrack flamegraphs contain only matching stacks. The filtered
print summary and size histogram remain whole-process outputs: for the same
large trace they contain 608,290 allocation calls. The unfiltered stack sum
and histogram agree with that count. Only flamegraph stacks are used for the
exact wrapper subset check. Histogram-derived byte totals are process scoped.

No production code, counters, frozen inputs, or captured observations were
changed in response. A future allocation qualification must use the existing
counter semantics explicitly before collection.
