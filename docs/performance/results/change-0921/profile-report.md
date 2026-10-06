# 0921 operation profile

This diagnosis uses the retained 0920 baseline and candidate binaries. It does
not build a new candidate, rerun the 0920 comparison, change the allocator, or
alter the 0920 adoption decision. Binary and fixture identities, raw artifact
hashes, and the exact frozen population are in
[`profile-results.json`](profile-results.json), with the capture boundary in
[`profile-protocol.json`](profile-protocol.json) and the method in
[`profile-method.md`](profile-method.md). Raw captures remain under
`/tmp/litchi-performance-0921-profile`.

The ten scoped Callgrind captures completed with nonzero `measure_one` events,
one `run_case` lifecycle, exact output guards, and exit code 0. Candidate
instruction references were lower in every cell: Float 1/2/8/32 by 1.59%,
2.25%, 2.39%, and 4.91%, and NoHeadFoot 32 by 5.42%. Data references and
simulated D1 misses were also lower. In a representative Float 2 capture,
inclusive `Package::render_with_layout` fell from 2,779,245 to 2,113,479 Ir;
the writer row fell from baseline `OleWriter::write_to` 2,255,926 Ir to
candidate `write_to_vec` 1,587,256 Ir. These are simulated attribution counts,
not native elapsed time, but they provide no evidence that the reservation
path added CPU work inside the measured operation.

Native `perf record` captured ten process-wide reports with zero lost samples.
The requested `measure_one` callchain filter was unavailable: neither
`perf script` nor `perf report` exposed `measure_one` or `run_case` despite the
retained symbols and DWARF sections. The reports therefore remain context only.
Their largest symbols are libc `memmove`, glibc allocation/free routines, DOC
sprm parsing, and SHA compression. The corrected one-child `perf stat` and
six `strace -f -c` controls are also process-wide. They show candidate page
fault deltas of +154%/+8%/−2%/+22%/−28% for Float 1/2/8/32 and NoHeadFoot 32,
but the retained 12-round 0920 page-fault receipts are highly variable: their
paired median deltas are +216%/+80%/−12%/+45%/−3%, respectively. Syscall totals
are 304/300, 506/598, and 338/274 baseline/candidate for Float 2, Float 32,
and NoHeadFoot 32. This is whole-process context and does not establish an
operation-local cause.

The useful next hypothesis is launch-sensitive copy and allocator/page
placement around `Package::render_with_layout` and `write_to_vec`, especially
for the marginal Float 8 timing result. A future diagnosis would need a
separately controlled source-line/branch and page-residency profile at that
boundary. The retained evidence does not support a reservation-loop CPU
regression or production adoption.

The first control attempt omitted the required probe `--input` argument; all
16 such children exited with `ProbeError("missing --input")` and are excluded.
The corrected controls are the only control rows summarized in the result
packet. Callgrind emitted the known Valgrind `brk segment overflow` diagnostic,
but each scoped probe exited 0 and emitted nonzero events; the warning is
retained as a limitation rather than silently discarded.
