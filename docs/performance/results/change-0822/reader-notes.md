# 0822 reader notes

`analysis.py` and `validate.py` are offline, fail-closed readers. They consume
retained 0822 receipts, JSON reports, compressed perf data and decoded frame
text. They never invoke Cargo, a workload, the probe, `perf`, `nm`, or
`objdump`. Root must run them only after every build, qualification, native,
perf, and decode child has reached a terminal receipt. The readers retain the
raw compressed and decoded perf artifacts as evidence; they do not recreate
them.

This packet measures a diagnostic perturbation of the admitted real PPTX edit
workflow. The pinned input is the 0821 `real-002-pptx` source (68,822 bytes,
SHA-256 `19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571`).
Its admitted default output is 68,284 bytes with SHA-256
`38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf`. The
public transaction opens the path, sets shape `(0, 0)`, commits/applies the
edit, serializes, hashes and reopens outside the timed edit region. These
identity and semantic checks are admission evidence and are not a latency
claim.

The three native arms are `control` (ordinary executable, direct call), `wrapped`
(ordinary executable, wrapper around the edit region), and `fp` (frame-pointer
executable, wrapper around the edit region). Qualification has three reports
with three measured samples each (nine samples total). Native capture has six process blocks, the
planned permutation order, 30 samples and three warmups per report: 18 reports
and 540 measured samples. Timings are diagnostic perturbation evidence only;
the reader makes no shipping-latency, historical-speedup, or adoption claim.

For every native report the reader computes nearest-rank p50, p95 and p99 and
retains mean, RSS and tail/spread flags. Paired comparisons are matched inside
each block: `wrapped/control` and `fp/wrapped`. Their bootstrap intervals use
10,000 resamples, seed `822822`, and sorted zero-based endpoints 250 and
9749. The statistic is the median of six matched block ratios. RSS is a whole
process diagnostic and is never subtracted from edit time.

The perf lane has two repeats of 2,000 samples, no warmup, `cycles:u` at 997 Hz,
frame-pointer call graphs, and the exact `edit_region_0822` wrapper. Raw perf
data and decoded non-inline frames retain their compressed/uncompressed
identity, period, headers, lost-event/status lines, and unknown-frame data.
The reader filters only samples containing the exact demangled owner in the
live binary/DSO identity recorded by the decoder. It reports exact-owner
sample and period counts, owner-qualified leaf ranks, inclusive call-path
counts, and separate whole-process/unattributed denominators. A zero-owner
lane is a qualification failure; it never becomes a fabricated phase
fraction. Nested inclusive counts are not added, and no Amdahl fraction is
recomputed from them.

Production quality is reused through the pinned 0821 reader proof and the
committed 0820 repair receipt. A fresh probe-quality run is checked separately.
No production or runtime-harness source changes are admitted. Build and
cleanup evidence must account for exactly two binaries (`ordinary` and `fp`).
Post-commit replay binds the measured base ancestor and exact source hashes;
the final seal may advance only to the aggregate packet HEAD.

Root final replay adds independent native and frame-audit --check gates,
including exact whole/owner/outside counts, owner periods, self leaves, and
the complete symbol leaf partition. The main analysis field
`qualified_stacks_with_unknown_interior` tests any descendant below the
owner, including the leaf. The independent frame audit separates unknown
leaf and unknown interior frames (excluding leaf and owner); the publication
table uses that explicit split. One empty stack is retained in the whole
process/outside-owner denominator.

Six unsuccessful agent reader attempts preceded the accepted output. Their
error messages/order are retained as an attestation, not original console
logs, in `reader-failure-attestation.json`; this evidence limitation is
explicit. Root retained its own duplicate-write log and final replay logs.
