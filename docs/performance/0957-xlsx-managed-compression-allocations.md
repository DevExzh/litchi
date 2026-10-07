# 0957 — managed XLSX compression allocation qualification

The cold managed medium one-percent route adds little cumulative allocation
at widths two and four, but width four raises its observed peak by **6.23–7.00%**
across the three process means. That exceeds the 5% review threshold. Width two
raises the same peak by **2.35%**. This batch qualifies the allocator harness;
it does not adopt a production change or clear the unresolved native timing
and RSS findings from [0956](0956-xlsx-managed-compression-native-widths.md).

## Capture and measurement boundaries

The dedicated `xlsx_managed_compression_alloc` executable reuses the existing
fixed fixtures, finite managed contexts, bounded sink, admission preflight,
and exact-byte/semantic/inverse oracles. It uses the existing serialized
process-wide `CountingSystemAllocator` observer. No production Rust path was
changed. The parent native benchmark only gains the child module declaration.

The frozen matrix contains three cases × widths 1/2/4 × three fresh processes:
**27 children**, each with two warmups and ten measured samples. All 270 rows
passed. Execution was serial on CPUs 12–15, with a 900-second child deadline,
no retries and stop on first failure. A separate width-four medium preflight
passed before formal capture. No child timed out or required forced cleanup.

Setup measures fresh context/source/sink allocation. The continuous core
region splits ingress, edit/commit, publication, commit drop and worker drain.
`core_combined` covers the first four; `core_combined_with_drain` includes all
five. Verification is outside those regions. `postdrop` measures the remaining
budget and sink destruction after verification. Every row enforces all six
additive counter totals, adjacent live-byte boundaries, combined peak maxima,
zero allocation failures, exact fixture hashes, finite limits, task/output
accounting, inverse restoration and released live resources.

The peak below is the maximum of setup and drained-core region peaks, less
the live bytes at setup entry. It is an observer-ordered high-water increment,
not RSS or unique heap. Process-lifetime peaks can be dominated by the
executable hash read before measurement and are not used as phase costs.

## Results

Values below are process means. Ranges preserve differences between the three
processes; these are descriptive counters, not native timing estimates.

| Case | Width | Drained-core allocation calls | Drained-core allocated bytes | Setup + drained-core peak increment, bytes |
| --- | ---: | ---: | ---: | ---: |
| Medium one-percent | 1 | 111,025 | 19,966,270 | 17,842,418 |
| Medium one-percent | 2 | 111,063 | 19,988,230 | 18,261,051–18,261,526 |
| Medium one-percent | 4 | 111,087 | 20,006,502 | 18,953,676–19,091,958 |
| Dense-sparse one-percent | 1/2/4 | 210,321 | 27,673,682 | 22,061,838 |
| Medium one-edit | 1/2/4 | 28,142 | 5,255,741 | 14,381,184 |

Medium width two adds 38 allocation calls (+0.034%) and 21,960 allocated bytes
(+0.110%). Width four adds 62 calls (+0.056%) and 40,232 bytes (+0.201%).
Its individual sample peak increments span 18,594,611–19,093,866 bytes;
width two spans 18,257,252–18,262,001 bytes. All three width-four process mean
peak ratios exceed 1.05. Both derived peak fields retain those flags in the
[full summary](results/change-0957/allocation-summary-v3.json).

The dense-sparse and one-edit admission controls remain serial at every
requested width. Their primary counts, allocated bytes and peak increments
are identical across the captured processes. This does not establish actual
OS worker width or overlapping compression tasks.

Phase regressions are retained separately. Medium publication allocation
calls increase by more than 5% at both parallel widths. Commit drop records
new nonzero allocation calls and bytes against a zero serial denominator;
those comparisons are explicitly undefined ratios with new-nonzero flags,
not a ratio of one. Width-four publication deallocation counts also trigger
individual flags. The final width-four process also has new-nonzero drain
allocation/deallocation calls and bytes. The drained aggregate does not erase
these phase findings.

## Asynchronous worker cleanup

The [source trace](results/change-0957/review/commit-drop-source-trace-v2.md)
establishes that pinned Rayon's pool destructor signals termination without
joining workers. The previous report now distinguishes handle destruction
from completed worker exit; its historical raw strings remain unchanged.

The allocator diagnostic uses a preopened `/proc/self/status` file and a
fixed buffer. Before each sample and after commit drop, it requires
`Threads == 1`, bounded by two seconds, 10,000 polls and one-millisecond sleeps.
The post-core wait stays inside the fifth allocator segment. Calibration runs
100 probes in a separate measured region and internally verifies no allocator
callbacks or live/peak changes. The report serializes four zero call counters
and the calibration verdict; byte/live/peak calibration internals are not
serialized independently.

All medium width-one and width-two post-core observations started at one
thread. Width four started at two threads in 12 of 30 samples, and at one in
18; all finished at one within one or two polls. One width-four sample recorded
allocator callbacks in the drain segment. These observations justify including
drain in aggregate accounting. They do not identify individual callback owners
or explain the 0956 commit-drop latency penalty. Instrumentation changes
execution, so these results cannot be used to quantify native cleanup overlap.

## Validation, provenance and remaining work

Release build, scoped rustfmt, both Clippy targets with warnings denied,
26 allocation-observer tests and allocator-binary tests passed. Workspace-wide
formatting failed on inherited files and is retained separately. Five CLI
negative controls passed, including preservation of an existing output
directory. Synthetic report/custody checks and independent source/tooling/
result reviews passed. The frozen source manifest reproduced exactly.

Two offline reduction failures are preserved: a missing import, then an
assumption about a redundant top-level validation index. The corrected reducer
still requires the complete child identity and rejects contradictory indices.
It changes no statistics, workload, plan or raw capture. The
[packet](results/change-0957/README.md) records the amended reducer binding,
external artifacts, regeneration commands and cleanup.

Width two remains the narrower candidate for further investigation, with no
adoption in this batch. Width four retains both this peak regression and the
0956 RSS finding. The earlier ordinary one-edit p99 regression also remains
unresolved: this explicit managed-context route cannot retire it. Native
commit-drop attribution and the broader `docs/GOAL.md` completion audit remain
open.
