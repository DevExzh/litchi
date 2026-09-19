# Native repeated-query diagnostic

This uninstrumented binary opens one owner, performs two preparation queries,
and then executes N identical selected queries on that owner. Results are consumed
and dropped per iteration; owner/cache remain alive through the entire loop.
The reported time excludes setup. External perf/time measurements include setup;
N=10 versus N=1010 instruction/cycle differences isolate the additional 1,000
queries, while process peak RSS includes the owner, mandatory catalogs and cache.
Use owned and file modes. File mode is ordinary warm OS-cache positional I/O,
not physical cold-cache evidence. This package is separate from the counting
allocator and from the corpus/phase-timing probe.
