# Change 0411: XLS read allocation baseline and XLSX lint gates

Date: 2026-09-05

Status: current diagnostic baseline; `performance_claim: none`.

This batch closes the XLSX Rust 1.98 all-target lint gate and adds allocation
observations to the six existing opt-in XLS open/list/selected-cell cases.
Production XLS/CFB behavior is unchanged. The non-iWork goal remains open.

The generated `litchi-xls-comments-opaque-heavy-v1` fixture contains two sheets,
256 comments on `Comments`, a numeric witness `Untouched!E21 = 42`, a sentinel
comment, eight 2 MiB opaque payload streams, and opaque metadata. The runner
checks eager/source worksheet names and the selected value against this
fixture. The historical 0278 mismatch on a different real-world workbook does
not apply to this generated oracle. The archive is 16,995,840 bytes and its
Workbook stream is 80,946 bytes. Final source hashes are bound in the retained
catalog and verification result.

Both timers start after source construction or the owned archive clone and
cover a fresh workbook open plus the named projection. They stop before
validation, reporting and owner drop. Allocation regions bracket those same
operations in the separate counting-System executable. The source timer also
includes the diagnostic `InstrumentedSource` range accounting and atomic
counters; its results are not a plain `OwnedSource` or `FileSource` proxy.

Four fresh normal processes each retain 500 samples per case after 20 warmups;
two separate allocator processes each retain 30 samples after three warmups.
Samples within a child share its heap state. These are warm generated-memory
measurements, with no cold file, filesystem I/O or remote-source claim.
Process RSS includes generation, layout/oracle, archive copies and reporting.
Allocation high-water fields remain process-lifetime snapshots rather than
isolated operation peaks. Normal and instrumented elapsed results are not
combined.

The normal table reports the minimum–maximum across four independent process
reports, in milliseconds. It does not pool their samples. Full means, sample
vectors, exact IID median intervals and repeat spread are retained in
`capture/verification.json`. The tabled allocation calls and allocated-byte
values are identical across both 30-sample allocator runs; allocation calls
include successful reallocations.

| Case | p50 ms | p95 ms | p99 ms | Allocator calls | Allocated bytes |
|---|---:|---:|---:|---:|---:|
| `xls_semantic_open` | 0.509–0.516 | 0.522–0.528 | 0.528–0.534 | 7,260 | 1,015,781 |
| `xls_source_backed_open` | 5.581–5.675 | 5.603–5.691 | 5.609–5.708 | 458 | 274,206 |
| `xls_eager_open_list_worksheets` | 0.511–0.519 | 0.523–0.533 | 0.529–0.536 | 7,263 | 1,015,894 |
| `xls_source_backed_open_list_worksheets` | 5.575–5.587 | 5.592–5.603 | 5.605–5.616 | 461 | 274,271 |
| `xls_eager_open_one_cell` | 0.510–0.515 | 0.524–0.530 | 0.531–0.534 | 7,260 | 1,015,781 |
| `xls_source_backed_open_one_cell` | 6.016–6.052 | 6.046–6.066 | 6.058–6.083 | 488 | 280,542 |

The source-backed open and list each make 334 logical reads totaling 138,459
bytes: 136,704 CFB structural bytes and 1,755 Workbook-global bytes. One-cell
makes 362 reads totaling 138,593 bytes, adding 134 selected-worksheet bytes.
Every retained source sample reads zero unselected-worksheet and opaque payload
bytes, retains the complete 16,995,840-byte input outside the operation region,
and verifies stable source identity. The archive materialization counter of zero
does not count the untimed input clone or parser-owned globals. Version counts
(168 open, 162 list, 224 one-cell) include the external before/after checks and
untimed open-name verification; they are not timed-only probe counts.

Normal six-case process RSS ranges from 144,672 to 145,368 KiB; allocator process
RSS is 145,120 and 144,900 KiB. These are whole-process observations.

The source one-cell CPU export contains 7,057 stack blocks and 30,600,358,794
weighted event periods. `InstrumentedSource::read_at` is 87.35% of whole-process
leaf weight (87.72% inclusive), 96.81% of the observed source-open subset leaf
weight, and 99.58% of the selected-cell subset. This locates the main observed
cost in the diagnostic adapter, not in XLS parsing. The eager export contains
1,290 blocks and 5,270,983,655 periods: whole-process leaf weight is 32.23%
`memmove` and 17.94% SHA-256. Its `Workbook::new` ancestor subset is 44.87%
of whole-process weight; sector collection and sector claiming appear among
its leading leaf symbols. Both exports have zero lost samples and zero unparsed
frame lines. Stacks containing unknown frames account for 0.69% source and
1.64% eager weighted periods. Setup/oracle and warmup calls are included.

Three whole-process PMU repeats per family report scaled IPC ranges of
2.006–2.046 eager and 4.221–4.235 source-backed. Hardware events are multiplexed
at 80–84% scheduling for eager and 83% for source; software events run at 100%.
All event counts, scheduling percentages, branches, faults and native L2 request
and hit counts are retained in `capture/resources.json` and raw CSVs. These
ratios characterize the complete diagnostic commands, including the adapter.
They do not measure operation-local IPC or an XLS/CFB efficiency improvement.

CPU profiles and hardware counters cover whole processes, including setup and
postvalidation. Production-ancestor call-chain subsets are sampled CPU
attribution, not timed phases. Native L2 events use the encodings validated in
0409; exact L1 and LLC counters remain unavailable on this guest. Request-size
distributions, exact memory-copy volume, lock wait time and scaling remain
unmeasured for these cases.

The default 36-case/198-row benchmark contract and 422-selector registry are
unchanged. The coverage index's `measured` status still binds the default
matrix. This independent opt-in baseline is documented separately; extending
that schema to admit independent baseline/catalog references remains work.

XLSX changes replace constant-width chunk iterators with fixed-array slices,
remove redundant namespace borrowing, and update test-only error wrappers and
temporary UTF-8 conversion. Existing divisibility and malformed-input checks
are retained. Rust 1.98.1 warning-denied all-feature/all-target Clippy and all
1,238 XLSX tests pass. The new benchmark status test shares the allocator tests'
lock to avoid observing their temporary global enable state.

The two focused XLS harness tests pass. After the shared-lock correction, the
XLS and allocator tests pass together: nine tests, four test threads. Crate
boundaries pass, all seven existing strict claims validate, and the report
classification gate remains unchanged. The read verifier accepts 12,000 normal
and 360 allocator observations; four mutations (duplicate timing order,
duplicate metric index, opaque read, wrong oracle) are rejected without changing
the original reports. A separate verifier checks all six RSS, two profile and
six PMU captures.

Production/lint commit: `38563aa0fe1b3f6af3306e9f83d2dcbf1e6bbe59`.
Instrumentation commit: `1f54d359e1ad72eb72c46134e256b80eb976201a`.
Final captured revision: `44edf790669a0aa4dc0aff73af6f7b5f5e709b6d`.
The final clean release builds use Rust/Cargo 1.98.1, release debug level 1,
frame pointers and unwind tables, CPU 2 and one worker on AMD EPYC 9R45 / Linux
7.0.0-1011-aws. They use the standalone benchmark workspace's release profile,
not the root workspace's LTO settings. Binary SHA/size, build commands,
predeclared protocol, raw captures, logs, flame graphs and verifiers are in the
[0411 evidence bundle](../results/change-0411/).

Static review identifies two observer costs: classification scans every stored
range for every read (the opaque 2 MiB streams alone contribute 32,768 physical
sector ranges), and `new_xls` enables a locked range-union calculation whose
repeated-range result is not published by the XLS summary. These are benchmark
implementation observations; this batch does not modify either path.

The next evidence work should measure and reduce the adapter's range-accounting
cost or add a matched plain `OwnedSource` observation while preserving a separate
full locality replay. Current evidence does not justify a production CFB/XLS
rewrite or a source/eager speedup claim.
