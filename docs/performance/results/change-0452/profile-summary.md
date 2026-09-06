# Whole-process CPU corroboration

These four diagnostic reports (stat/record per binary) are separate from the
formal timing matrix. cycles:u excludes sleeping but includes untimed corpus
construction, gates and output hashing. Frame-pointer stacks do not resolve the
lifecycle iteration frame in either recording, so no timed-region attribution
is asserted. All raw perf data, text, binary hashes and command receipts remain.

| Counter | Baseline | Candidate | Change |
|---|---:|---:|---:|
| cycles:u | 24539976382 | 24166564716 | -1.522% |
| instructions:u | 47206894735 | 46723608275 | -1.024% |
| branches:u | 6282035809 | 6248483550 | -0.534% |
| branch-misses:u | 132041897 | 130592070 | -1.098% |
| L1-dcache-load-misses:u | 0 | 0 | uninterpreted zero |

The zero L1-miss event is not evidence of no cache misses. Whole-process cycles
fall about 1.5% and instructions about 1.0%; untimed work dominates these totals.
SHA self period is about 38% and medium Deflate about 29–30% in both records.
Those totals cannot quantify the timed publication gain. Deterministic owner
read counters and matched API clocks provide the direct evidence.

Weighted symbols, missing callchains (7 baseline / 2 candidate), sample counts
and exact period totals are reproducibly derived in `profile-summary.json`.
