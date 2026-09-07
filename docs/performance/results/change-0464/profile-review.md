# 0464 profiling scope

The applicable diagnostic is `profiling-r1/profile-summary.json`, captured by
`profile-r1.py`. The original `profile.py` stopped at an erroneous UID-zero
guard without trying perf. That result remains as historical driver evidence;
it does not establish an environmental profiler limitation. The amendment
probes actual user-mode counters and recording, and both succeed without a
permission change. Sampling disables perf's build-id cache.

Both diagnostic workloads use the original normal binary and the measured
source epoch. The source-compatibility check permits only the independently
built semantic-inventory binary source to differ in the live tree. Each
workload executes three warmups and 30 measured iterations. These 60 retained
diagnostic rows are separate from the 240 formal timing samples.

User-mode whole-process counters report 1,758,215,198 instructions,
631,274,012 cycles, 327,372,796 branches, 3,514,541 branch misses,
2,295,690 cache misses and 1,709 page faults. PMU events ran for about 83% of
the window under multiplexing; page faults ran for 100%. The raw generic
cache-reference value is zero and is not used to calculate a hit/miss ratio.
Dividing totals by 33 is only a descriptive whole-process normalization.

The DWARF cycle profile has **20 samples**, with no lost samples reported.
The symbol table lists SHA-256 compression at 19.96% and zlib inflate at
14.25%. The parsed symbol strings also contain perf's trailing IPC placeholder
columns; the raw `samples/top-symbols.txt` remains authoritative. This small
sample does not establish a stable ranking or a production API hotspot.
Binary/input hashing, setup, post-operation ZIP/XML oracles, output/report
serialization and teardown are included in the process window. They are
excluded from the four formal API clocks. An API-specific profile or a longer,
controlled sampling window is required before using this diagnostic to choose
a production optimization.
