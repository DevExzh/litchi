# 0921 operation-local profile method

This is a diagnosis packet for the rejected 0920 CFB output-reservation
candidate. It does not revisit the 0920 adoption gate and does not authorize a
performance recapture.

The frozen population uses the retained 0920 native binaries, whose source,
candidate patch, probe tree and fixture identities are recorded in
`profile-protocol.json`. The base and candidate binaries are checked by SHA-256
before each capture. The candidate is the unchanged 0920 base revision plus its
retained patch; no source or probe file is rebuilt or edited.

Callgrind reuses the 0918 scoped mechanism. Instrumentation starts at process
launch, collection starts only on the non-inlined `measure_one` symbol, and the
single `run_case` edge is checked in each profile. The collected operation is
`Snapshot::open_bounded` through repeated paragraph-zero replacement, commit,
output-vector copy and local destruction. Source I/O, route preparation,
reference routes, correctness and public-reader checks, hashing and JSON are
outside this region. Callgrind `Ir`, data references and simulated cache misses
are attribution counters, not native elapsed time.

`perf stat` is feasible for user-space hardware/software counters under the
current `perf_event_paranoid=1`. Its page-fault and context counters cover the
whole child process, so they are reported as process controls and interpreted
alongside the retained 0920 `/usr/bin/time -v` page-fault receipts. Kernel
syscall tracepoints are unavailable; six bounded `strace -f -c` runs record
whole-process syscall counts for the two FloatingPictures endpoints and the
NoHeadFoot control. They do not become operation-local claims.

Native `perf record` was also run for the ten-cell matrix at 2 kHz with DWARF
callchains. The retained `perf.data` files have zero lost samples, but neither
`perf script` nor `perf report` exposed `measure_one` or `run_case` frames on
this host despite the symbols and debug sections being present. The requested
operation-only sample filter is therefore unavailable; those reports are
retained as whole-process context and are not used as operation attribution.
The scoped attribution comes from Callgrind, whose `measure_one` toggle did
collect nonzero events for every row.

The diagnostic question is whether the candidate's extra `ReusePlan::output_len`
and owned-output reservation adds CPU work inside the accepted publication path,
or whether page faults, allocator placement, or setup dominate the observed
FloatingPictures regression. The profile matrix compares Float 1/2/8/32 and a
NoHeadFoot 32-edit control with the same CPU, retention, warmup and sample
settings. Raw Callgrind, perf and strace outputs remain under
`/tmp/litchi-performance-0921-profile`; only small hash-bound summaries belong
in this packet.
