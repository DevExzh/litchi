# OPC source-cache release evidence: next change

This run is a structural and distribution capture from committed revision
`a1b692297b2493a3f523aa064e6be366271c4f52` on an isolated `git archive`.
The 30-cell contention matrix passed its cache, gate, Budget, source-I/O,
worker-team, and package-drop invariants. The exact-budget boundary passed as
well.

No production performance claim is accepted. The source delay is a deterministic
10,000-us test instrument and is not production storage, scheduling, allocator,
or decompression latency. No cell passed both independent ABBA directions with
95% confidence intervals excluding zero, so no managed-versus-control speedup
is reported.

The next evidence change should add per-sample allocation count/bytes and peak
live bytes, process RSS, and attributable hardware counters (cycles,
instructions, branches, branch misses, cache misses, and page faults) to this
same release harness. It should preserve the same corpus, worker widths,
10,000-us disclaimer, 3 warmups, 30 samples, CPU affinity, and balanced
`control-A, managed-A, managed-B, control-B` order, then repeat the structural
and directional gates before any optimization or speedup decision.

Until that instrumentation exists, allocation, peak RSS, and hardware-counter
results remain explicitly unavailable rather than inferred from elapsed time.
