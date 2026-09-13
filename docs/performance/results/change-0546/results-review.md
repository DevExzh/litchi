# 0546 isolated scanner result review

The frozen screening rule passes; this only authorizes a fresh integrated
campaign. All 4,000 samples and ten flagged metric/repeat comparisons remain.
The automatic summary uses a generic adverse-label template; the specific
interpretations below apply to each metric, including medians and means.

| Kind | Fixture | Variant / repeat | Metric | Change | Review |
| --- | --- | --- | --- | ---: | --- |
| drift | sparse | baseline | p95 | 6.304% | Repeat tail variability retained without rerun or filtering; limits precise tail claims. |
| drift | sparse | counted | p99 | 44.438% | Repeat tail variability retained without rerun or filtering; limits precise tail claims. |
| adverse | clustered | 1 | p50 | 174.012% | Bulk counting scans the long sparse suffix twice after the probe. Clustered p50 adds about 28 microseconds; full valid sparse planning guards are required. No runtime gain claimed. |
| adverse | clustered | 1 | mean | 171.451% | Bulk counting scans the long sparse suffix twice after the probe. Clustered p50 adds about 28 microseconds; full valid sparse planning guards are required. No runtime gain claimed. |
| adverse | clustered | 1 | p95 | 175.885% | Bulk counting scans the long sparse suffix twice after the probe. Clustered p50 adds about 28 microseconds; full valid sparse planning guards are required. No runtime gain claimed. |
| adverse | clustered | 1 | p99 | 100.844% | Bulk counting scans the long sparse suffix twice after the probe. Clustered p50 adds about 28 microseconds; full valid sparse planning guards are required. No runtime gain claimed. |
| adverse | clustered | 2 | p50 | 174.597% | Bulk counting scans the long sparse suffix twice after the probe. Clustered p50 adds about 28 microseconds; full valid sparse planning guards are required. No runtime gain claimed. |
| adverse | clustered | 2 | mean | 173.058% | Bulk counting scans the long sparse suffix twice after the probe. Clustered p50 adds about 28 microseconds; full valid sparse planning guards are required. No runtime gain claimed. |
| adverse | clustered | 2 | p95 | 174.688% | Bulk counting scans the long sparse suffix twice after the probe. Clustered p50 adds about 28 microseconds; full valid sparse planning guards are required. No runtime gain claimed. |
| adverse | clustered | 2 | p99 | 112.490% | Bulk counting scans the long sparse suffix twice after the probe. Clustered p50 adds about 28 microseconds; full valid sparse planning guards are required. No runtime gain claimed. |

The measured helper contains a 64 KiB cap check and two indirect count-dispatch
calls with needles 0x3c and 0x26 at 0x236d6 and 0x236e9. The pinned memchr
source confirms the specialized single-byte count path. The complete measured
binary disassembly is retained; this is not a hardware-counter, instruction-count
or cross-CPU performance claim. No allocator, RSS, concurrency or cold-cache
benefit follows from the diagnostic.
