# 0784 — initial PPTX capture profile

The 0783 large fixture spends about 69% of its diagnostic lifecycle in initial
capture. This packet profiles that public operation directly, using a private,
non-inlined wrapper in the evidence probe. Its body only calls
`Package::opened_presentation` and returns the snapshot. No production source
is changed. Setup, serialization, readback and retained owner destruction are
outside the Callgrind collection boundary.

The exact wrapper symbol controls collection, zeroing, and one dump at return.
Two passes cover tiny/medium/large in forward then reverse order. Each process
performs one capture without warmup. Guest instruction counts describe the
emulated operation, not native wall time or hardware cycles. Six alternating
native blocks compare the wrapper build with the feature-off control to
identify timing perturbation separately. Source/output/semantic identities
are checked against the unchanged 0780 capture fixtures.

The inherited lifecycle phase feature remains available in the source but is
not enabled in either build or used for this capture. The schema/tool/binary
identify 0784. The allocator helpers are inherited but disabled. The dependency
lock and original generated marker remain unchanged.

Reproduction requires a fresh packet output directory and owned target at the
recorded base, then `build.py`, `capture.py`, and `profile.py`, serially. Retained
outputs must not be overwritten. Generated manifest paths must name the fresh
checkout. No profile-timer result is mixed with native latency. See `plan.json`
for the frozen matrix and `origin.json` for ownership.

The Callgrind/software-SHA observation led to `perf-plan.json`, frozen before
two ordinary native records. Their DWARF decodes have zero exact capture-owner
stacks and are retained as unqualified. `perf-fp-plan.json` then freezes a
separate frame-pointer build and two native records. Canonical non-inline
frame decoding restores fully qualified Rust symbol names; inline decodes are
also retained. Warmup and measured capture calls are both included in these
sample profiles. Frame-pointer code generation and unqualified stacks limit
these to descriptive diagnostics. One recovered capture stack has an unknown
interior frame, so the frozen plan forbids native phase-fraction claims; only
exact observed sample/period counts are reported.

For that complete follow-up, run `perf_capture.py`, `build_fp.py`, and
`perf_fp_capture.py`, then `perf_decode.py perf`, `perf_decode.py perf-fp`, and
`perf_frames.py`. Decoding requires the exact executables still present.
Compression retains each original SHA-256/length and the compressed artifact;
offline replay verifies the decompressed bytes before using them. The first
ordinary exploratory `.script` and `.flat.txt` decodes are retained as well;
the canonical `.decoded.gz` files have structured decode receipts.

Generate derived results with `native_analysis.py --write`,
`profile_analysis.py`, and `perf_analysis.py --write`. The seal and cleanup
witnesses are batch-specific; never copy them to a new experiment. Final
retained replay is `python3 -B docs/performance/results/change-0784/validate.py`.
The returned native sampled categories form a disjoint partition; namespace
resolution and SHA rows are nested diagnostics and must not be added to it.
