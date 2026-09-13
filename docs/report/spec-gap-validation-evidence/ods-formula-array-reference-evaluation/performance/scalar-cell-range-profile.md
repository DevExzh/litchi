# Range-1 CPU profile diagnostic

The retained [`scalar-cell-range-profile-01.tar.gz`](diagnostics/scalar-cell-range-profile-01.tar.gz)
contains the raw `perf.data`, reports, and process output for the before ELF
and the current `ScalarCell`-last candidate. Each run used CPU 6,
`cycles:u`, two warmups, five measured iterations, and repeat 200,000 for the
`reference-range-1` workload. The samples cover the whole process, including setup and warmups.

The before report's largest sampled symbols were `Budget::reserve` (25.51%),
`RawVecInner::finish_grow` (15.75%), `ExecutionContext::consume` (7.88%), and
`memmove` (6.87%). The candidate report showed the same leading regions:
25.07%, 16.12%, 9.07%, and 6.50%, respectively. Candidate stacks also show
the demand-aware evaluator path; this profile alone cannot establish whether
that path caused any timing change. Both reports had zero lost samples, but
several unwound frames remain raw addresses, and the reports provide no
causal attribution.

This is supporting evidence for the paired range-1 control review only. It is
not a speedup claim or a replacement for repeated A/B timing and source-level
analysis; the residual control regression remains open.
