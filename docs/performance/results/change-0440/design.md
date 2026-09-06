# 0440: ODP attribute namespace ownership hypothesis

The sealed 0439 existing-append baseline shows large operation allocation cost
and repeated ElementAttrs lookups in whole-process profiles. The candidate
hypothesis is that per-attribute copies of immutable namespace URI bytes are
unnecessary while processing one element. The accepted ADR tree remains
c950b6c8be822561b498d7bbe87c460873dcbf49, read in prior goal turns.

Before editing production source, rebuild and retain the unchanged binaries
and capture the first baseline phase (A1). Compare the same unchanged existing
append selector over 64/4096/8192 source slides, normal and allocator binaries,
30 samples and three warmups, CPU 2 and one worker, with opposing repeats.
All CPU work runs serially in the root. Agents provide only external source
drafts and reviews. No iWork work is included.

The frozen order is A1/B1/B2/A2. After both candidate repeats, the root restores
the exact baseline parser file between terminal jobs for A2 and the baseline
profiles. It then restores the retained candidate file for candidate profiles
and final validation. Every workload checks its role's complete source manifest
and binary identity; no rebuild substitutes for either retained binary.

The candidate must preserve URI equality, default-namespace exclusion for
attributes, unknown-prefix behavior, first-match order, duplicate and malformed
attribute error order, decoder behavior, and drawing-attribute harvest order.
No validation or final readback may be skipped. The ordinary owned input and
materialized commit lifecycle and its retention endpoints remain unchanged.

Acceptance requires a practically useful measured change (at least 5% normal
p50 or allocation calls/requested bytes on medium/large), repeat consistency,
and review of every normal-latency or RSS regression exceeding 5%. Candidate
complexity is rejected if representative results do not justify it. Existing
0439 evidence is immutable and is context, not a substitute for fresh before
measurements. The full non-iWork program remains open.
