# Measurement and analysis review

This is a read-only review of `analysis.json` and the retained child reports.
`python3 -B tables.py --check` passed, so all six CSV tables match the admitted
analysis. No native command was run for this review.

The native lane has 280 child receipts and 70 groups: 56 timed policy groups
(seven corpora, lifecycle and atomic publication, four policies) plus 14
default-only edit/counting controls. Every group has four process blocks. The
analysis contains 50 process-spread flags: 40 p99, eight p95, and two mean;
there are no p50 flags. The largest is the refused DOCX counting-sink p95 at
58.6367%. Whole-process RSS is analyzed separately and has zero flags across
all 70 groups. The default-versus-Full paired control contains 14 flagged
metric series, 11 p99 and three p95, with no mean or p50 series flagged.

The report's native table values agree with the CSV medians, which are the
medians of four process p50 values converted from nanoseconds to milliseconds.
No process quantiles are pooled. The retained trace analysis has 28 valid rows
and agrees with the stated measured sync contract: default and Full each have
one file and one parent-directory sync, FileOnly has one file sync, and NoSync
has neither. This trace lane remains diagnostic rather than native timing
evidence.

The allocation lane has 140 child receipts and 70 groups, with two blocks and
three samples per process. Across all 28 timed `(corpus, phase, block)` keys,
the four policies have exactly equal raw operation vectors for allocation,
deallocation, reallocation, failed-allocation counts, and allocated and
deallocated bytes. The derived `peak_above_start` and `net_live` vectors are
also exactly equal. Absolute live and region-peak gauges retain deterministic
entry-state offsets between policy processes; they are not claimed to be
policy-equal. Across the 70 groups and 13 reported allocation metrics, all 910
two-block p50 pairs are equal, giving zero repeat-spread flags.

The user-counter summary has 32 whole-process children and 16 marginal pairs
(two repeats for each of four corpora and the default/NoSync policies). All
four events report 100% runtime coverage for both sample counts. Recomputing
the instruction marginal means gives the report's rounded NoSync-versus-default
changes: −0.007% generated DOCX, +0.131% generated XLSX, +0.034% generated
PPTX, and −0.008% NumberedList. These are whole-process `(23−3)/20`
observations that include report work and do not establish operation-level
causality.

No numerical inconsistency or analysis blocker was found. The results support
the report's descriptive current-source policy comparison. They do not support
a causal latency claim, a physical durability claim, an RSS/allocation
equivalence claim, or a device-floor conclusion.
