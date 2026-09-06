# 0452: Retain verified compressed captures through PPTX copy planning

PPTX image/chart planning now uses the combined OPC read/capture API. An opaque
retained capture shares verified bytes and their reservations; each publication
authorizes independent writer staging under the same source lineage, version and
execution context. Metadata remains charged for empty and long-name captures.
Existing semantic and byte checks govern reuse; only memory failures permit the
ordinary decoded fallback. The consuming public PPTX editor is unchanged.

The frozen primary matrix has 16 fresh reports and 480 samples (two providers,
two corpora, two builds, reverse repeat order; CPU 2, one worker, 3 warmups/30
samples). Complete API time includes open, plan and publication. Simulated
range/media p50 improves 31.324% and 31.048%, meeting the predeclared 10% gate.
Plain medians change less than 0.4%. All output hashes and lifecycle gates match.

Primary bytes/media baseline drift makes its 36% first-pair improvement
unreliable as a stable claim. A separately frozen ABBA confirmation retains four
fresh reports/120 samples and confirms about 8% API improvement, with about 6%
extra planning and 19% less publication time. No primary sample is discarded or
pooled. All 23 primary and 10 confirmation review flags have recorded dispositions.
See [primary measurements](../results/change-0452/measurements.md) and
[confirmation](../results/change-0452/confirmation-summary.md).

Source publication data reads fall from 425 calls/16,786,581 bytes to zero;
total source bytes fall 33,584,577→16,801,705. Destination I/O is identical.
Source work after open falls 50,366,359→33,589,143 units. Source reservations grow
16,807,458→33,622,602 bytes while the plan lives, then both return to 16,807,458.
Staged decoded copies remain. Whole-process RSS and CPU profiles include untimed
fixture generation/hashing, so no timed allocator-peak reduction is claimed.
See [regression review](../results/change-0452/regression-review.md) and
[profiles](../results/change-0452/profile-summary.md).

Final-source validation passes 471 OPC, 848 PPTX and 381 harness tests (1,700),
strict lint, warning-denied rustdoc, non-iWork workspace, format and boundary
checks. Five new OPC tests cover concurrent writer reservations, retry, source
freshness, empty metadata and long metadata; the new PPTX test checks zero
publication media-range reads and exact output. A retained old-source negative
control reproduces the empty metadata accounting hole. Two intermediate compile/
fixture-limit failures are disclosed with final passing replacements. The
extended OPC fuzz target passes strict lint and 1,000 ASan/coverage runs.

The portable bundle binds source copies, both exact binaries, frozen protocols,
commands, raw results, test logs and sixteen fuzz seeds. Replay after owned
temporary cleanup with:
`python3 -B docs/performance/results/change-0452/verify.py --sealed --cleanup`.

No registry/default/native coverage is promoted. Shared staged decoded ownership,
tight-budget admission, native application breadth, cold I/O, bounded existing
append, repackaging and scaling remain follow-up work. The full non-iWork goal
remains active. Accepted ADRs and the user-owned goal file are unchanged.
