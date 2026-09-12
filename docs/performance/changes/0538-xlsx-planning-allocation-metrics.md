# 0538: measure source-backed XLSX planning allocations

The cell-value benchmark previously observed allocations only during commit
and publication. Worksheet decoding happens earlier, inside `edit_sheets`,
so those observations could not evaluate the transient attribute ownership
candidate identified in 0537.

The harness now records `plan_allocation_metrics` for all eight managed and
unmanaged source-backed cell-value publication selectors. Its region surrounds
the existing planning timer after selector construction. The returned edit
remains alive; cell updates, commit, publication, evidence collection, reopen
and verification follow outside the region. Each phase has a separate region,
and planning samples preserve acquisition order and exclude warmups. Normal
executables report unavailable allocation counters explicitly. Lifecycle
controls without this observation continue to omit the optional field.

This is a measurement enabler, not a production optimization. The 0537 `Cow`
candidate remains unapplied. The [evidence bundle](../results/change-0538/README.md)
records focused tests, checks and real normal/allocator binary schema smoke
captures for medium and dense-sparse corpora. Their debug elapsed values are
not accepted performance baselines or speedup evidence. Process-global allocator
region peaks are not RSS or allocator-internal peaks.

The next experiment must freeze fresh matched release sources and gates before
applying the candidate, then compare native latency, planning instructions and
allocations, commit/publication and eager-read guards. Correctness, preservation,
error order and retained cell-type ownership remain mandatory. OLE2/OOXML stay
first; ODF is deferred until their optimization goal completes, and iWork is
excluded.
