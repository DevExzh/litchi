# Change 0409: XLSX profiling and corrected member-read accounting

Date: 2026-09-04

`performance_claim: none`

`claim_authorized: false`

## Result and implementation

This batch extends the current baseline to XLSX selected-cell queries and
matched one-cell edit/save. Profiling exposed a harness bug: the source-backed
cell-value selectors constructed `InstrumentedSource::new`, leaving XLSX
member ranges empty. Their workbook and worksheet read vectors were false
zeroes. Total source reads, timing, output, and preservation checks remained
valid; those member vectors could not establish unread members.

Commit `193ca7b4297ec35d07e5cdd5f64d144b822c5be9` configures compressed member
ranges through `new_xlsx`, classifying selected worksheets from the actual
update set. The conditional configuration marker
`xlsx_cell_values_range_accounting: "compressed-member-intersections-v1"`
identifies the eight affected source-backed selectors. The comparator rejects
unknown versions, markers on unrelated selectors, and comparisons between
marker-bearing and historical markerless configurations. Historical omission
remains readable. Tests cover single-sheet and all-sheet ranges, matched
semantic output, raw preservation, managed budgets, and marker scope.

These counters accumulate compressed data-range intersections across open,
planning, commit, and publication. They are not reset between phases.
Positive unselected-sheet counters include copying untouched raw members to
the output; they do not establish semantic parsing or decompression. Range
classification and source construction occur before timing, while the read
observer itself runs inside the timed source calls. Production library code
and timing boundaries are unchanged. No old/new latency claim is authorized.

This harness-only correction follows ADR 0005's explicit observation and
measurement contract and ADR 0008's evidence gates. ADR 0001/0002/0010/0011/0024
ownership boundaries, ADR 0003 transactions, and ADR 0006 preservation and
validation paths are unaffected. No new runtime, pool, dependency or public
API is introduced.

## Capture identity and scope

The seven initial reports bind source
`abe38a9570129c6646bb1b1d7207c407fc86c3d6`; the four corrected reports bind
`193ca7b4297ec35d07e5cdd5f64d144b822c5be9`. Both captures verify that the only
worktree entry is the user's untracked `docs/GOAL.md`, SHA-256
`bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1`.
They are descriptive observations, not a clean ABBA experiment.

Rust 1.98.1 release binaries use debug level 1, forced frame pointers and
unwind tables, four build jobs, and disabled incremental compilation. The
separate harness workspace does not inherit root release LTO settings. Runs
are pinned to CPU 2 on AMD EPYC 9R45, Linux 7.0.0-1011-aws, perf 7.0.12,
a KVM guest with 32 affinity CPUs and about 128 GiB RAM. Exact commands,
runtime/build settings, executable hashes and build IDs are retained in the
[artifact bundle](../results/change-0409/).

All reports use the deterministic medium corpus
`litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1`: four 48-by-48
worksheets, 9,216 cells, 4 MiB of media, 17 ZIP members, 4,226,429 archive
bytes, and 4,231,168 uncompressed payload bytes. Archive SHA-256 is
`dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036`.

The filesystem selected query opens the workbook and prepares the query before
timing. The timer selects case-insensitive `bEnCh01`, resolves canonical
`Bench01` at zero-based position 1, and reads M29 (row 28, column 12), exact
numeric lexical value `1028012`. Each retained sample uses a fresh child with
warm filesystem cache. Full semantic/hash verification follows the timer and
procfs snapshots; deterministic oracle rebuilding and compression also occur
outside that interval. External perf and time commands include this surrounding
work and inherited children. They do not measure operation-only CPU work.

Edit/save changes Sheet1!A1 and times open + selector planning + commit +
sequential publication in process. Reopen, semantic/hash verification and
cache/budget diagnostic sampling are excluded. The generic filesystem
configuration flags do not turn these in-process edit selectors into isolated
filesystem runs. Neither physical-cold nor scaling evidence is captured.

## Descriptive observations

| Run | Warmups / samples | Observation and scope |
| --- | --- | --- |
| Initial selected normal | 20 / 500 | Query p50/p95/p99 3,528,178 / 3,585,663 / 3,634,612 ns; mean 3,531,808.138 ns, 95% Student-t mean interval [3,529,089.610, 3,534,526.666] ns |
| Initial selected allocator | 3 / 30 | Constant 81,918 allocation calls (including 12 reallocations), 10,690,444 allocated bytes, zero failed allocation calls per query |
| Initial source one-cell save | 20 / 500 | p50/p95/p99 8,023,488 / 8,111,320 / 8,169,520 ns; member read vectors invalid as described above |
| Initial eager one-cell save | 20 / 500 | p50/p95/p99 8,290,619 / 8,363,149 / 8,403,609 ns; same output as source case |
| Corrected source one-cell save | 20 / 500 | p50/p95/p99 8,108,978 / 8,206,399 / 8,261,579 ns; mean 8,113,128.372 ns, 95% Student-t mean interval [8,109,376.922, 8,116,879.822] ns |
| Selected perf stat | 5 / 100 | Whole process: 123,470,821,661 cycles; 352,955,457,181 instructions; 61,465,328,689 branches; 732,439,547 branch misses; 197,808,801 generic cache misses; 1,528,307 page faults |
| Selected native L2 | 5 / 100 | Whole process: 3,733,177,440 access requests; 3,593,856,498 hits; 139,320,942 miss requests |
| Selected perf record | 10 / 300 | cycles:u at 999 Hz with frame-pointer callchains; 80,527 stack blocks and 337,888,625,521 weighted event periods; zero lost samples reported |

Allocator timing is instrumented and is not comparable to normal latency.
Allocation calls/bytes are operation deltas; live and peak-live allocator
values remain absolute process/lifetime snapshots. Allocation volume does
not measure copied bytes. Normal selected maximum process RSS is 116,240 KiB
and command wall time is 132.88 s. Initial source/eager edit process RSS is
82,736/82,684 KiB, with command wall time 23.88/13.38 s. Their different
surrounding verification work prevents treating that wall-time difference as
an API publication regression. The edit selectors currently have no
operation-local allocator or decoded-byte observation; these are explicit gaps.

The source one-cell runs retain 257 total logical reads and 4,233,005 returned
bytes per sample, with three payload materializations/cache loads and 64,321
retained cache bytes. Initial source, eager and corrected source runs share
output SHA-256
`9b7b66a02007eeb63498fd5de4c6b7115ace0383ce37d97e1a9560ef7bfadec1`,
4,226,480 accepted sink bytes, 201 writes, and largest write 32,768 bytes.
The 15 untouched members retain their raw preservation digest
`7105fcbce160328f666e69fcfd18da9e19fd71dd7b63961e7cddd29d5da1a17d`.

Corrected member observations are constant within each run:

| Source case | Retained samples | Workbook calls / bytes | Selected worksheet calls / bytes | Unselected worksheet calls / bytes |
| --- | ---: | ---: | ---: | ---: |
| One-cell save | 500 | 1 / 226 | 1 / 6,816 | 3 / 20,330 |
| Two-sheet edit smoke | 3 | 1 / 226 | 2 / 13,593 | 2 / 13,553 |
| All-four-sheet batch smoke | 3 | 1 / 226 | 4 / 27,146 | 0 / 0 |
| Managed one-cell smoke | 3 | 1 / 226 | 1 / 6,816 | 3 / 20,330 |

The three smoke cases use one warmup and establish attribution/guard behavior,
not latency. The managed case verifies typed output-budget refusal with zero
accepted output. The first recapture validator incorrectly expected four
selected sheets from the two-sheet selector. The successful reports were
retained and revalidated with the correct count; the all-sheet batch was
captured separately. Both the initial validator and failure log are retained.
The validated 500-sample report was not rerun to replace its observations.

## CPU attribution and cache-event limits

Whole-process self attribution is dominated by surrounding Deflate work
(32.80% `deflate_medium`, 15.80% `longest_match`), SHA-256 (7.79%), and corpus
payload generation (4.69%). The retained all-symbol/no-inline export separates
the library query using the exact `SelectedWorksheet::cell::<&str>` ancestor:
2,528 blocks and 8,291,214,922 weighted event periods. Within that subset,
`clone_bounded_name_part` is 16.22% leaf attribution. Of that helper's leaf
weight, 75.50% comes directly through expanded-name cloning and 24.26%
through namespace expansion; the remaining 0.23% comes directly from
`parse_element`. `parse_element` is 6.08% self and 23.31%
inclusive. Lexical `clone_bounded_bytes` is only 0.038% self and 0.361%
inclusive; inlining and shared allocation/copy routines limit finer attribution.

These are period-weighted sampled CPU fractions, not elapsed phase timers.
Inclusive rows overlap and cannot be added. Frequency adaptation and short
fresh children also limit interpreting the selected subset's share of the
whole process. Rich inline postprocessing completed with retained addr2line
warnings; the supplementary no-inline commands completed without warnings and
preserve physical callchains. The parser binds input hashes and original Git
source blobs and retains its exact ancestor and denominator definitions.

Local sysfs/perf inventory and contrasting 32 KiB hot / 256 MiB streaming
probes validate the available native L2 request events. In the actual same-run
three-event capture, access minus hit equals miss request exactly, with all
three events scheduled at 100%. Probe events were separate invocations; no
cross-run probe ratio or subtraction is inferred. Generic `cache-misses` maps
to event 0x64/umask 0x09 here and is not an exact independently verified LLC
counter. Generic cache references and L1 load/miss aliases return untrustworthy
zeroes on both probe workloads. Exact LLC events are not exposed in this
guest; perf's L3 metric names are not valid direct event selectors. This is a
recorded host limitation, not a claim of zero activity. No operation-local
instructions-per-byte or cache-locality claim follows from process totals.

## Validation and remaining work

The release harness library passes 258 tests with one ignored. Final focused
comparator tests pass 86 tests; the full Python suite runs 862 tests, with 842
passing and 20 skipped. Scoped Rust formatting and crate-boundary checks pass.
The boundary command's initial pinned-1.95.0 invocation failed because that
installation lacks Cargo; the explicit 1.98.1 rerun passes. Initial formatting
reported three whitespace changes, which were applied before the final build.
Strict verification passes all six registered performance claims. The report
classification checker continues to validate the existing 167 rows.

All eleven report/catalog pairs pass corpus binding and structural metric
validation. Selected-query raw children additionally pass exact value/digest,
unique PID, sample alignment and operation schema checks. Corrected source
reports validate configuration identity, positive workbook/selected overlap,
update-set classification, phase-sum/sample-order identity, and preservation
witnesses. Source and binary identities remain fixed around each capture.
The bundle includes complete commands, validation logs, raw perf data,
symbolized reports, a flamegraph, and losslessly compressed large perf scripts,
with a standalone published-artifact hash verifier.

The next measured candidate is expanded XML-name/frame ownership in the
selected-cell parser. A narrow lexical `ElementData` borrowing experiment is
plausible but is not established as the dominant cost. Any change must retain
bounds/error ordering, decoded owned values, namespace/MCE behavior, callback
lifetimes, buffer reuse, and recovery, then repeat correctness and profiling.
Current XLS eager/source selected-cell baselines and the wider CRUD,
native-producer, cold-source, allocation, copy/decode, and scaling matrix remain
open. No production speedup or full-goal completion is claimed by this batch.
