# Root performance assessment

The final instrumented profile contains 18 lanes, 54 fresh processes, and
1,620 measured samples. Root rechecked all raw semantic, forward-preservation,
inverse, changed-state, and allocator-balance gates; verified the native ZIP and
Theme hashes against the corpus; and ran the full 4,619-path source-manifest,
per-process percentile, and pooled-report verifier.

At p50, native eager open plus Theme/family read is **0.323 ms** and source-backed
open/read is **0.296 ms**, with 21 logical source reads returning 4,167 bytes.
The prepared native family update plus forward and inverse publication takes
**2.263 ms**, requests **1,360,747 allocation bytes**, and reaches **79,173
additional live bytes**. The corresponding opaque-heavy update is **8.442 ms**,
**4,963,609 requested bytes**, and **317,276 additional live bytes**. These are
instrumented absolute measurements of the stated operation, including inverse
publication; they exclude package save/reopen and are not production latency
promises. See the full [profile](performance/report.md) for tails and all lanes.

Family cloning performs zero observed allocations for both the 261-byte native
projection and 26,949-byte opaque projection. The source-sharing gates refer to
the Family fragment for `family_clone` and the complete Theme XML for `noop`.
The latter still performs validation/publication work and allocates: its native
p50 is 0.729 ms with 396,762 requested bytes. Source sharing is not a claim of
an allocation-free host no-op.

Repeated host parsing and validation remain measurable costs. The base codec
and host lanes have different scopes; no matched pre-feature host control was
built. Therefore a before/after regression or speedup percentage, including the
program's approximately 5% regression trigger, cannot be established here. This
is a limitation of this batch's absolute evidence, not a claim that the new
metadata discovery has negligible cost. The synthetic case exercises 96 foreign
namespace-bearing extension children; one larger input does not establish an
asymptotic scaling curve or worst-case performance.

Hardware counters were unavailable (`perf_event_paranoid=4`); the failed attempt
is retained. RSS is whole-process RSS, logical `ReadAt` observations are not OS
syscalls, and requested allocation volume is not retained memory. Owned profile
and caller targets, temporary files, and Python caches were removed. The
pre-existing workspace target was retained.
