# Shared-format and process-counter review

All four frozen ABBA guard processes pass, retaining ten rows and thirty
samples per row: 1,200 samples in total. Each selector has exactly one tiny
and one large result. Complete corpus, source, sink and output projections
are identical across all four arms/repeats. XLSX intentionally has no top-level
source object in this producer schema; its sink and operation metrics remain
present and checked.

There is one paired latency statistic above the positive five-percent review
threshold. Tiny XLSX p99 changes from 129,370 ns in G1-control to 153,571 ns
in G2-candidate (+18.707%). With thirty samples, the producer's nearest-rank
p99 is the maximum. G2's next-largest sample is 108,220 ns. Its p50 improves
from 116,090 to 97,030 ns, its p95 from 125,441 to 108,220 ns, and its mean
from 117,587.333 to 99,910.467 ns.

The second frozen pair does not reproduce that p99 penalty: G3-candidate has
p99 112,620 ns versus G4-control 124,131 ns. G3's next-largest sample is
105,761 ns. Candidate p99 drift between G2 and G3 is -26.666%, so this tail
instability remains a review flag. The observation is an isolated maximum
in G2; the evidence does not establish its cause or a globally regression-free
result. No capture was discarded or replaced.

Tiny ODS candidate p99 changes from 187,041 to 175,780 ns (-6.021%) between
the two candidate runs. This is the other repeat-drift statistic above five
percent. Neither matched ODS pair exceeds the positive five-percent regression
threshold. All other guard mean/p50/p95/p99 matched and repeat checks stay
within their applicable threshold.

Whole-process maximum guard RSS in ABBA order is
288,380 / 288,664 / 288,616 / 288,856 KiB. No paired or repeat RSS flag
exceeds five percent. This is the lifetime maximum of a ten-case process and
includes setup/preflight; it is not a per-operation peak.

The four large PPTX process-counter captures also pass. User cycles in ABBA
order are 159,464,877,622 / 150,296,131,798 / 149,885,952,842 /
157,445,867,055; instructions are 690,860,463,972 / 676,077,523,507 /
676,221,144,601 / 690,887,378,817. Hardware events have approximately 83%
scheduled time and perf scales multiplexed counts. Software events have 100%.
LLC load misses are unsupported. L1 data-cache load misses are reported as zero,
which does not establish that no misses occurred. These counts include
preflight, warmups and observers and cannot be substituted for writer-only
phase costs or normal elapsed samples.

The main large normal maximum RSS is 82,708 / 82,744 KiB for control and
82,780 / 82,744 KiB for candidate in R1/R2. Allocator-process values are
82,732 / 82,792 and 82,616 / 82,796 KiB respectively. Every main paired
normal statistic, normal repeat statistic, incremental operation-peak and
process-RSS check meets its frozen threshold. Allocator elapsed values remain
separate and do not authorize a latency claim.
