# 0956 native managed-worker experiment protocol review

Reviewed 2026-10-07 against checkout HEAD
c8e80705989183d88950e282ef1852f3e3470b10 and the retained 0955
qualification. This is a protocol review only. No checkout source, build, or
workload was changed. The checkout already had unrelated modifications in
tools/check_iwork_leaf_public_api.py, crates/litchi/examples/bench_word_parse.rs,
and docs/UNIFIED_OPS_API_DESIGN.md; this review did not touch them.

## What 0955 establishes

docs/performance/0955-xlsx-managed-compression-qualification.md:3-18,20-39
and docs/performance/results/change-0955/validation.json establish admitted
compression width and semantic/resource correctness, not elapsed performance.
The medium one-percent case admits 0/2/4 worker permits at configured widths
1/2/4, with five CpuTasks at each width. Dense-sparse one-percent admits
0/0/0 and also charges five tasks; medium one-edit admits 0/0/0 and charges
two tasks. The nine-case matrix passed exact source/output hashes, eager output
equality, semantic reopen, inverse restoration, no-op, cancellation, refusal,
and post-release resource checks. The 0955 report explicitly reserves a paired
native experiment and says dense-sparse and one-edit are controls
(0955-xlsx-managed-compression-qualification.md:75-79).

The retained identifiers for the prospective native run are:

- medium one-percent source
  dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036;
  output fc742a7ad139e2ce2e75fabfff8d2a33408dd5eeb10997717a2e2656d8423e6c;
- dense-sparse one-percent source
  893ad3f5dd6a98aec44bc541a140048072c84c579b4b9e332431f779b097cb1a;
  output 07eee3f980949f031776087ed61f06125edd82cb29394a608a3d8d5a4b28374d;
- medium one-edit source is the medium source above; output
  9b7b66a02007eeb63498fd5de4c6b7115ace0383ce37d97e1a9560ef7bfadec1.

The source report uses “cold” to mean a fresh package and explicit execution
context per case. It does not prove cold filesystem or page-cache state; a
native report should state the filesystem-cache condition separately.

## Source and boundary findings

The existing timed route is run_xlsx_cell_values_edit_save in
tools/perf-baseline/src/lib.rs:44633-44775. Lifecycle gates and eager output
construction happen before the sample loop (44646-44649), output-refusal
replay is outside the loop (44692-44709), and per-iteration sink reservation
and source backing are prepared before the timed open (44717-44731). The
timed sequence begins at the fresh managed open (44742-44764), then selector
planning (44765-44775), commit/staging, and publication. Reopen, output,
semantic, inverse, and resource checks remain after the timed phases in the
existing function. A native comparison must retain these boundaries and must
not move verification into the measured interval.

The current benchmark helper is not width-parametric:
tools/perf-baseline/src/lib.rs:46331-46349 hard-codes
NonZeroUsize::new(1) for both execution worker fields and constructs the root
limits with unbounded u64::MAX input/work fields. The prospective isolated
diagnostic therefore needs a width argument threaded into the managed context,
with finite input/work/object/depth/output/CPU limits matching the 0955
qualification. Changing a global CLI --workers selector alone is not proof
that this route changes width. Keep the 0955 memory ceiling high enough for
four worker states and fail the cell if admission narrows the requested width.

The source implementation confirms the expected mechanism. In
crates/litchi-opc/src/source_backed.rs:5436-5445,5555-5571,
write_limits derives a ParallelWriteLimits from the context and managed
packages create a private lazy ParallelWriteSession; there is no ambient
process-global scheduler. admit_deflate_wave at 5640-5685 charges CPU tasks
at every width, then reserves worker/memory permits for a parallel wave and
narrows the wave if the budget cannot grant the requested width. A permit
observation therefore proves budget admission, not OS thread overlap. The
native timing experiment must record both configured and observed/effective
width.

## Recommended bounded capture

Use one clean native release executable and one runtime width option, with
widths 1, 2, and 4 passed into the same managed-context construction. Do not
compare separately allocated or separately compiled binaries. If allocator
telemetry is desired, use the same binary for all widths or run a separate
descriptive instrumentation arm; never mix an allocator-build result into the
native width comparison or adoption gate. Bind every child to the same binary
SHA, source revision, rustc/toolchain, profile, Cargo flags, and corpus
identities.

Run the medium one-percent case as the primary. Run dense-sparse one-percent
and medium one-edit in every schedule as controls. The controls are expected
to stay serial under the 0955 threshold; a width effect there measures setup
or noise and invalidates a simple compression-scaling interpretation.

Use fresh processes for every schedule leg. A concrete balanced schedule for
each pair (1,2), (1,4), and (2,4) is four blocks: two A B B A blocks
and two B A A B blocks, where A and B are the pair's widths. This gives
16 fresh child processes per pair per case and 48 per case, 144 across the
three cases. Each child should run at least 20 warmups and 100 measured
samples; process launch and report serialization are outside each sample.
Retain every raw per-child phase vector, width, schedule position, process
terminal, and affinity record. Do not pool samples while hiding child or
block identity.

On this host, the read-only topology check found 32 online CPUs, allowed list
0-31, all on socket 0, and CPU IDs 0-31 map one-to-one to the reported core
IDs. Use a fixed mask such as 0-3 for every width, and verify each child
with /proc/<pid>/status or taskset -pc plus the recorded lscpu mapping.
Fail closed if fewer than four distinct physical cores are allowed or if a
child is not on the required mask. Pinning all widths to CPU 12, as in the
single-CPU historical setup, would serialize a width-2/4 local pool and
cannot answer the scaling question. Record CPU model, governor/frequency
state, allowed list, and host-load conditions with the capture.

Preserve the existing phase boundaries and report at least open, selector
plan, staging/commit, publication, and total durations. Compute per-child
nearest-rank p50 and p95 from the 100 measured samples; p95 is the declared
tail-5% statistic. Pair ratios at the block/child level and report the raw
vectors, then summarize medians and intervals across paired fresh processes.
A bounded diagnostic no-regression guard should be declared before capture:
the width-B p95 must be no more than 1.05 times width-A p95 for the primary
and both controls, with the corresponding p50 policy applied separately.
Retain all p95 tails and flag any failed child/cell; never relax the guard
after seeing results. This is an experiment validity/adoption input, not
evidence that a width is a production optimization.

Before accepting any timed result, run the retained 0955 nine-case
qualification and require exact source/output/eager/semantic/inverse checks,
sink accepted-byte/write-size checks, expected CPU-task and effective-worker
counts, and zero post-drop live memory/worker/I/O/object resources. A failed
or narrowed width is an invalid cell, not a smaller successful treatment.
Keep lifecycle, refusal, cancellation, reopen, and inverse checks outside the
timed window but retain their outcomes per child or preflight. Report RSS or
same-binary allocation observations separately; do not use a different
allocator binary as a native width control.

The resulting claim should be limited to the captured host, affinity,
binary, corpus, source-cache state, and fresh-process schedule. A faster
medium primary with serial controls is workload-specific scaling evidence; it
does not establish general XLSX publication speed, OS thread overlap, or an
Office-wide recommendation without the declared timing, tail, correctness,
resource, and cross-case gates.
