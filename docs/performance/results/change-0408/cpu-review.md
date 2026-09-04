# CPU capture review

The 7,124 raw sample records report zero lost samples. Dominant application
frames resolve through frame-pointer callchains and inline paths. An unresolved
kernel frame contributes 0.61%, and each postprocessor retains two addr2line
warnings. Complete unwinding is not claimed.

Whole-process inclusive samples: verifier 71.82%; into_opc_package 16.37%;
materialize_opc_package_with_accounting 16.34%; corpus construction 3.83%;
binary identity hashing 2.95%. These rows overlap, are not timed phase shares,
and must not be added or used as a speedup comparison. The normal selector
calls into_opc_package with accounting disabled despite the shared private
materializer's function name.

Self samples: SHA-256 74.49%, memmove 10.00%, CRC32 7.30%, payload generation
0.66%. The resolved SHA callers include post-timer verification at 71.45% and
binary hashing at 2.89% of process samples. Memory movement includes 6.19% in
the materialization read/decompression path and 3.74% in source construction.
These are descriptive sampled attribution, not exact timer/copy measurements.

The timer excludes oracle/source/package setup, verification, report assembly
and final owned-package destruction. Allocator deltas bracket that timer;
process live/peak/RSS fields retain their documented absolute scope.

L1 aliases return unvalidated zeroes; LLC aliases report unsupported. No
zero-cache-activity or before/after performance claim follows.
