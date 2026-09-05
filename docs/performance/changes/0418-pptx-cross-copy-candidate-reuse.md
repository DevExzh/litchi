# Change 0418: reuse proven owned PPTX cross-copy candidates

> Correction (0421): allocator `peak_live_bytes_*` values in this record
> may under-report peaks and must not support high-water claims.
> [Counter correction](0421-allocator-peak-counter.md) explains the defect.
> Live-byte totals, allocation requests, normal timing, RSS and Heaptrack
> metrics are unaffected by this specific bug. Raw reports are unchanged.

`performance_claim: scoped warm owned PPTX latency; memory tradeoff reviewed separately`

`claim_authorized: true; primary media lifecycle latency only`

## Production change

The candidate path retains the freshly constructed, serialized, and reopened
PPTX package produced during application-time planning.  For an unmodified
owned OPC destination, a read-only `is_unmodified_owned_source` check permits
that proven candidate to be reused instead of rebuilding the same package.
The reuse path still requires descriptor equality, source and destination
graph and physical fingerprints, patch preconditions and postconditions,
limits, the target revision, and the final physical fingerprint before
assignment.

Dirty destinations, caller-defined `Part` implementations, non-default
`SaveOptions`, and destinations whose authorization was revoked use the
existing clone-and-apply fallback.  Reuse does not restore authorization to a
changed package; the reusable candidate owns the new artifact through the
ordinary owned ingress path.  The inverse path retains its detached restored
candidate only after the existing fresh forward-plan proof.  The default
selector matrix is unchanged.

## Matched revisions and correctness gates

The matched control is `79dfee5025276676d433a80b0b5475ae03f09db8`; the
candidate is `f8f9e6667284ae28fdbf9b9313b2ff9583ec74fc`.  Both revisions use the
same corrected harness and registry assertions.

The control PPTX all-feature correctness run reported 828 tests and 2
ignored.  The candidate PPTX/OPC all-feature run reported 1,270 tests and 3
ignored.  The candidate lifecycle tests and the two selector-registry tests
also passed.  These counts establish correctness coverage for the exercised
fixtures; they do not establish a performance result or native Office
producer compatibility.

Candidate all-target/all-feature Clippy passes with command-scoped exemptions
for the three unchanged existing lints: `chunks_exact_to_as_chunks`,
`clone_on_copy`, and `needless_lifetimes`.  The matched control reproduces the
same existing lint findings.  No repository lint policy was weakened.
Warning-denied documentation builds for both crates pass. All 291 comparison,
corpus, packaging, claim and coverage-index tests pass. Portable replay and
12 semantic corruption probes pass. The strict registry validates nine claims;
`claim-0418-pptx-cross-copy-lifecycle` covers only the primary media lifecycle's
p50, mean, p95 and p99 on the named corpus, machine and builds.

## Measurement protocol

The capture uses ABBA order (`control`, `candidate`, `candidate`, `control`),
CPU 2, one worker, fresh processes per selector and leg, and warm generated
in-memory sources.  Inputs and sink reservation are outside the timed
lifecycle.  The primary selector is
`pptx_cross_copy_media_rich_lifecycle` with 500 samples after 20 warmups.
The plain lifecycle guard is also 500 samples after 20 warmups.  The existing
plain phase guard is 500 after 20, while the existing media-rich phase
selector is a 100-sample, 10-warmup diagnostic.  The protocol retains 6,400
normal observations, 480 allocator observations, and 32 preflights.

The allocator lane uses 30 samples after 3 warmups in the same ABBA order;
allocator latency is excluded from the latency comparison.  Its operation
region is available for the two lifecycle selectors; the legacy phase
selectors do not expose an operation-local allocation region.  Allocation
calls and bytes, literal live and peak snapshots, and whole-process maximum
RSS are retained for separate review.  The snapshots do not support a
per-operation peak claim, and RSS includes the complete process and
preflight.

A descriptive whole-command PMU run uses 20 samples after 3 warmups for
`cycles:u`, `instructions:u`, `branches:u`, `branch-misses:u`,
`cache-references:u`, and `cache-misses:u`.  It retains event availability and
multiplex scaling; it does not support timer-local IPC or cache-miss claims.
Same-revision drift ceilings are 5% for p50, mean, p95, and p99.  Any adverse
paired result above a ceiling requires review.  The complete protocol is in
[`protocol.json`](../results/change-0418/protocol.json).

## Measured result and resource tradeoff

Media lifecycle p50 falls from 1,175.449 / 1,172.463 ms in the controls to
720.698 / 723.741 ms in the candidates: paired reductions of 38.69% / 38.27%.
Mean, p95 and p99 also improve in both pairings, and all same-revision drift
checks remain below the declared 5% ceiling. Plain lifecycle p50 improves
4.04% / 3.33%; the existing plain phase guard improves 5.46% / 3.67%.
The 100-sample media phase diagnostic improves 40.54% / 40.79% at p50, but
is below the strict registry's 500-sample requirement and supports no registered
latency claim. The full per-leg table retains every statistic and interval.

Normal media lifecycle whole-process peak RSS rises from 819,028 / 820,056 KiB
to 884,576 / 886,172 KiB (+8.00% / +8.06%). The separate allocator lane agrees
(+8.01% / +8.15%). Media end-of-region live bytes rise from 222,238,733 to
287,885,862; these global snapshots include surrounding retained state and are
not an operation peak. Plain RSS changes stay below 0.15% in magnitude.
This is a latency optimization with an explicit retained-memory cost.

Media allocation calls fall from approximately 66,621 to 61,745 per lifecycle,
while cumulative requested allocation volume changes only from approximately
36.8615 GB to 36.8119 GB (decimal). Those totals count allocator requests,
including repeated allocation/reallocation, and do not mean that many bytes
are simultaneously resident or physically copied. Plain lifecycle allocation
volume falls from 29,442,859 to 24,558,673 bytes, and calls from 53,299 to 49,681.
The legacy phase selectors have no operation allocation attribution.

The [source review](../results/change-0418/source-review.md) records production
safety checks and fixture boundaries. The [memory review](../results/change-0418/memory-review.md)
assesses the RSS threshold crossing before retention. Raw reports, exact
oracles, deterministic projections, uncertainty and profile evidence are in the
[0418 bundle](../results/change-0418/).

The paired whole-command profiles retain 33,277 / 21,347 stacks, zero lost
samples and 0.375% / 0.634% unresolved leaf weight. Deflate-family leaf weight
falls from 71.89% to 57.32%; SHA-family weight rises from 20.69% to 32.20%,
while its absolute sampled weight is nearly unchanged. This supports the
repeated-compression mechanism, but these are whole-command weights that
include untimed work. Remaining hashing and candidate construction require
further attribution before another optimization.

Paired PMU observations retain 145.101 / 92.868 billion user cycles and
350.576 / 204.407 billion instructions, with 83% counter running time.
The zero cache-reference alias alongside positive cache misses is marked
unvalidated; no cache-rate or timer-local IPC claim follows from these counters.

## Remaining scope

The broader performance program remains open.  This change uses generated
owned PPTX fixtures and does not establish cold, remote, native Office or
LibreOffice, semantic CRUD, concurrency/scaling, or full non-iWork matrix
coverage.  The borrowed real-producer fixture establishes typed provenance
refusal only.
