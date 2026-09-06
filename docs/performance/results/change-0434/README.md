# Change 0434: ODS bounded text-span evidence

The production candidate is `65dfb1b141190271763367defce4086ef146c056`;
the baseline is `91de70866df67bafd7d2f4c8931269ea540e1a1d`. The private ODS
encoder groups ordinary validated UTF-8 text into borrowed spans of at most
256 bytes. Scalar fallback preserves deterministic row-window and Work-limit
refusals. Public APIs, escaping, XML auditing and archive publication stay
unchanged. See [design](design.md) and [ADR compliance](adr-compliance.md).

The frozen [protocol](protocol.json) specifies 24 reports and 720 timed samples:
before/after, normal/allocator, 64/8,192/32,768 rows with four cells per row,
two repeats in ABBA order, CPU 2, one worker, three warmups and thirty samples.
Both historical binaries run from the candidate checkout; build source identity
is retained separately from runtime checkout identity. Exact archive/XML/output
hashes, semantic digest and sink accepted-byte counts must match across roles.
Pilots are retained separately and do not enter the formal summary.

The verified [summary](summary.json) retains these operation medians:

| Rows | Repeat | Before (ms) | After (ms) | Change |
| --- | --- | ---: | ---: | ---: |
| 64 | R1 | 0.184826 | 0.161726 | -12.499% |
| 64 | R2 | 0.192766 | 0.161516 | -16.211% |
| 8,192 | R1 | 20.694993 | 17.853292 | -13.731% |
| 8,192 | R2 | 20.636727 | 17.928946 | -13.121% |
| 32,768 | R1 | 82.441442 | 71.260562 | -13.562% |
| 32,768 | R2 | 82.719983 | 71.559096 | -13.492% |

There are no matched >5% regression flags for latency, throughput, requested
allocation metrics, or process RSS. Baseline tiny p99 repeat drift is +9.844%
and remains flagged. All allocation count/byte/regional-peak vectors are
unchanged across the matched roles; the regional peak above entry is 419,347
bytes at every tested size. Exact archive/XML/output/semantic identities pass
for all twelve matched pairs. These observations support retaining this narrow
encoder change; `claims: []` leaves broad performance claims unregistered.

Whole-process sampled `ExecutionContext::consume` self CPU proportion changes
from 25.01% to 10.57%. The recording includes setup and untimed oracle work,
so it does not establish an operation-local CPU fraction. Raw perf data,
decoded stacks, symbol reports and counters are retained under `profiles/`.
The symbol reports retain 10 baseline and 11 candidate addr2line warnings;
function-level observations do not imply complete inline/source resolution.

Measurements use Rust 1.98.1 release executables with frame pointers and debug
symbols, the system allocator, Linux 7.0.0-1011-aws and an AMD EPYC 9R45.
Normal operation elapsed vectors own latency observations. Allocator vectors
record requested allocation calls/bytes and regional peak above entry. GNU-time
RSS and perf profiles cover the whole fresh process, including corpus setup
and the untimed oracle. Bootstrap intervals describe within-process samples;
they are not independent-process confidence intervals. Unsupported counters
and physical-copy attribution must remain explicit limitations.

Root runs CPU work serially. Reproduction commands, from the repository root:

```sh
python3 -B docs/performance/results/change-0434/capture.py --phase A1 --attempt formal
python3 -B docs/performance/results/change-0434/capture.py --phase B1 --attempt formal
python3 -B docs/performance/results/change-0434/capture.py --phase B2 --attempt formal
python3 -B docs/performance/results/change-0434/capture.py --phase A2 --attempt formal
python3 -B docs/performance/results/change-0434/profile.py --role before-streaming --kind stat
python3 -B docs/performance/results/change-0434/profile.py --role before-streaming --kind record
python3 -B docs/performance/results/change-0434/profile.py --role after-streaming --kind stat
python3 -B docs/performance/results/change-0434/profile.py --role after-streaming --kind record
python3 -B docs/performance/results/change-0434/summary.py --check
python3 -B docs/performance/results/change-0434/portable-probes.py --stage final
```

Capture refuses to overwrite retained results. Rebuild the two source revisions
using their build receipts and capture into a new bundle with fresh binary
identity descriptors. Sealed evidence verification needs no Rust build or
retained temporary executable. The bundle inventory and lossless deterministic
gzip records cover all retained logs and profiles. Failed development receipts
remain visible in [development notes](checks/development-notes.md).

The broader non-iWork goal remains open. The [next-work audit](next-work.md)
prioritizes bounded ODT paragraph creation, then ODP. Existing-document append,
package-Part addition, repackaging, native producer breadth, cold/range I/O and
explicit scaling remain separate obligations.

Copied-bundle verification, independent summary rederivation and six mutation
probes pass before and after task cleanup. The successful precleanup receipt is
`checks/precleanup-portable-v3.json`; the two preceding failed attempts remain
retained with their reasons. Cleanup removed 1,816,806,376 bytes across the three
exact task directories and preserved both shared target directory identities.
`docs/GOAL.md` remains untouched. The final staged inventory replay is retained
separately in the terminal evidence receipt.
