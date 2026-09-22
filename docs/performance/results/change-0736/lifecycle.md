# 0736 probe lifecycle audit and controlled cadence design

This is a source audit and a prospective measurement design. It makes no
production change, does not reinterpret the rejected 0735 result, and does not
attribute the secondary regression to any particular mechanism. The design
keeps the complete 0735 preservation oracle for every output that is used as
measurement evidence.

## What 0735 actually measures

The format operation's owner interval is the call to
`measured_public_format`. `timed_format` starts its `Instant` immediately
before that call and reads the elapsed time immediately after it returns
(`probe/src/lib.rs:2449-2473`). For PPT, the called route copies the source
bytes into `litchi_ppt::slide_order::Snapshot`, removes slide position one,
commits, and returns the committed bytes (`probe/src/lib.rs:2421-2447`). The
returned output `Vec<u8>` is moved out of `TimedFormat`; its length is passed to
`black_box`, and the vector remains alive after the timer stops
(`probe/src/lib.rs:2467-2473`). Thus `whole_ns` excludes output validation and
includes the public open/edit/commit work and construction of the returned
output.

The initial source inventory and the expected output are built before either
warmups or measured samples. `inventory` opens a copied input buffer and stores
a separate `Vec<u8>` for every stream in its `stream_bytes` map
(`probe/src/lib.rs:526-580`). The expected output is itself produced by one
untimed public edit (`probe/src/lib.rs:3000-3025`). Those source and expected
inventories remain owned by the process for the whole run. The expected oracle
and the eight corruption controls are also constructed before the sample loop;
the controls are not repeated for each sample (`probe/src/lib.rs:3015-3039`).

For the three configured native warmups, 0735 executes the timed owner and
then drops only the returned output. It does not call `output_sample`, so it
does not run the per-output structural, stream, directory, or semantic oracle
for a warmup (`probe/src/lib.rs:3046-3071`). The warmup owner call is therefore
followed by a different amount of work and a different ownership transition
than a measured iteration. This is a lifecycle fact, not a correctness failure
in the retained packet: all retained measured outputs passed the full oracle.

For every measured native sample, the timer stops first, and the output is then
passed to `output_sample` (`probe/src/lib.rs:3073-3110`). That function hashes
the output, builds a complete output inventory, and invokes
`oracle_for_output(..., enforce_policy = true)` before returning the sample
(`probe/src/lib.rs:2677-2717`). The oracle checks structural validity of source,
expected, and output bytes; exact stream paths and bytes; unchanged streams;
allowed changed streams; directory shape and metadata; CLSIDs; raw-directory
policy; and public semantic reopen witnesses (`probe/src/lib.rs:2276-2418`).
Every measured sample in the sealed packet has `oracle.oracle_ok == true`.

The validation is outside `whole_ns`, but it still runs between owner calls.
`inventory` copies the complete output into stream-owned vectors
(`probe/src/lib.rs:526-580`). `structural_valid` makes another owned source
copy for each validation (`probe/src/lib.rs:582-586`). The PPT semantic reopen
creates owned source and output inputs in both the raw and public projections
(`probe/src/lib.rs:1116-1178`, `probe/src/lib.rs:1276-1364`, and
`probe/src/lib.rs:2162-2272`). These allocations are not part of the reported
owner interval, but their allocator/cache state is present when the next owner
interval starts.

There is a second, less visible lifetime effect. `PptSlideWitness` contains
`live_record`, text, outline, notes, and comment values; the actual payload
fields are marked `serde(skip)` (`probe/src/lib.rs:1236-1270`). The combined
projection stores those values in each witness (`probe/src/lib.rs:1366-1407`).
`semantic_reopen_check` then stores source order, output order, and a cloned
removed-slide witness in the returned `PptSemanticWitness`
(`probe/src/lib.rs:2242-2272`). The measured loop pushes each returned
`Sample` into `samples` (`probe/src/lib.rs:3079-3109`), and the samples vector
survives until the final JSON serialization (`probe/src/lib.rs:3157-3217`).
Consequently, the output `Vec<u8>` is dropped when `output_sample` returns, but
the hidden per-sample semantic witness remains live for all later owner calls
in the same process. The serialized JSON does not show that hidden retention.

For the primary 11-slide fixture, one retained PPT semantic witness contains
22 `PptSlideWitness` values (11 source, 10 survivors, and the removed slide).
For the secondary 36-slide fixture it contains 72. The source projection also
copies the `PowerPoint Document` and `Current User` streams before taking live
record slices (`probe/src/lib.rs:1116-1178`). These counts describe ownership
shape; they are not a live-byte measurement. The current probe intentionally
does not expose the skipped payloads in JSON.

The allocation binary has a separate but related boundary. In
`allocation_format`, the output is placed in an outer `Option` while
`alloc_metrics::region` takes its after snapshot, so the returned output is
included in the region's boundary-relative retained ownership
(`probe/src/lib.rs:2616-2627` and `probe/src/alloc_metrics.rs:118-139`). The
region snapshots before and after the owner body, then drops only the closure
return value; validation happens afterward through `output_sample`
(`probe/src/lib.rs:3111-3154`). Allocation counters therefore exclude the
oracle, while the oracle still changes the allocator state before any later
region in the same process. The reported `peak_live_bytes` and
`retained_bytes` are allocator-counter boundaries, not RSS.

## Evidence that sample position must remain visible

The sealed 0735 native reports retain all 50 owner samples, and the 0736
post-hoc order analysis verifies their original indices and full-window
medians. Across the nine primary baseline processes, the median of the last
ten samples is 18.35% above the median of the first ten (range 15.89% to
19.53%). The corresponding candidate median drift is 0.25%. The secondary
baseline and candidate drifts are both downward (medians -2.78% and -1.25%),
while the secondary candidate-minus-baseline p50 remains positive in both the
first-ten and last-ten descriptive windows. These windows overlap in the
post-hoc report and are not independent process replicates. They are evidence
that sample position is a required diagnostic dimension, not a reason to
discard early or late observations. The original full-window rejection and
its nine independent process-pair statistics remain the decision evidence.

The compact serialized sample totals in `order-analysis.json` are about 1.016
MB for a primary process and 2.739 MB for a secondary process. They are JSON
bytes written at the end of a run, not measurements of the hidden witness
heap, and must not be used as a memory attribution.

## Controlled design for the next qualification

The next probe should make the lifecycle an explicit factor while retaining
the same exact oracle. The owner timer must remain unchanged: start immediately
before `measured_public_format`, stop immediately after it returns, and never
include validation or receipt construction in `whole_ns`.

Every warmup and every measured output in the strict arms follows this sequence.
Both 50-sample strict arms drain all full warmup witnesses before measurement;
only the measured witnesses differ in lifetime:

1. Execute the owner operation and obtain its output. For a native sample,
   retain only the owner interval in the timer. For an allocation sample,
   retain the existing `alloc_metrics::region` boundary.
2. While the output is still alive, invoke the existing full
   `output_sample`/`oracle_for_output` path with the same source, expected
   inventory, replacement set, and `enforce_policy = true`. Do not substitute
   a digest-only check, a stream-count check, or the once-per-process expected
   oracle. Require every nested boolean and `oracle_ok` to be true.
3. Create a compact receipt containing the sample index, owner time or
   allocation region, output SHA-256, output inventory summary, and the
   serialized visible oracle fields. The full semantic comparisons have already
   happened before this conversion; the receipt is not the validation itself.
4. Explicitly drop the output, temporary output inventory, reopen snapshots,
   and full semantic witness before the next owner operation. Keep only the
   compact receipt or its digest. For measured samples, the strict-retained arm constructs and keeps
   the same compact receipt before additionally retaining the full current
   `Sample` witness; the strict-drained arm drops that full witness at this
   point. A retained receipt may be written to the packet after the process
   exits; it must not keep the skipped `PptSlideWitness` payloads alive between
   samples in the drain arm.

The strict per-sample oracle is therefore preserved even in the drain arm. The
drain changes only post-validation ownership. If an implementation chooses to
retain the visible oracle JSON for independent audit, both variants must use the
same receipt construction path and the hidden witness must still be dropped.
The raw output report, source and expected hashes, and per-sample oracle status
remain mandatory evidence; no sample is accepted merely because its receipt
serialization succeeded.

Use these arms, with both PPT fixtures and the existing fixed pair rotation:

| arm | process lifecycle | oracle cadence | retained state | purpose |
| --- | --- | --- | --- | --- |
| legacy-50 | exact 0735: 3 owner-only warmups, 50 samples | full oracle after each measured sample | unchanged full witnesses | contemporaneous reference; preserves the original measurement contract |
| strict-drained-50 | 3 validated warmups, then 50 measured samples in one child | full oracle after every warmup and sample | compact receipt only | primary controlled native result; removes cumulative witness retention while keeping oracle work between owner calls |
| strict-retained-50 | 3 validated warmups, then 50 measured samples in one child | full oracle after every warmup and sample | full current `Sample` witness retained | matched diagnostic; estimates sensitivity to accumulating the full witness, including hidden payloads |
| strict-fresh-1 | 3 validated warmups, then 1 measured sample, one child per sample | full oracle after every operation | process exits after receipt | removes all post-sample state from the next owner call and checks whether the 50-sample trajectory is a process-local effect |
| allocation-0-drained | no warmup, one measured allocation region in a fresh child | full oracle after the measured output | compact receipt, then exit | preserves the 0735 no-warmup allocation boundary |
| allocation-3-drained | 3 validated owner warmups, then one measured allocation region in a fresh child | full oracle after every warmup and sample | compact receipt, then exit | controls allocator state and oracle cadence before the measured region |

First run an A/A harness qualification on the unchanged accepted owner, comparing
legacy, strict-retained and strict-drained arms under one binary where feasible.
Do not change production source during this qualification. Only after the
controls are audited should a separate, newly frozen baseline/candidate trial
reintroduce the archived rejected candidate.

For each native arm in that subsequent trial, use the same three cycles and three repeats per fixture as
0735 (nine baseline/candidate process pairs per arm and fixture), with CPU 12
pinned, serial execution, rotated baseline/candidate order, and a manifest that
records arm and process order. Use the same source fixtures, build flags,
limits, operation, and full preservation constraints. Allocation arms retain
the allocation lane's no-latency-claim status; use three paired repeats for
each warmup level and report every region field.

Strict-fresh-1 has one timed sample per child and nine independent process
pairs per fixture; it is a first-sample control, not a 450-sample distribution
or a within-process trajectory. Do not compare its tail precision to the
50-sample arms.

The strict-drained-50 versus strict-retained-50 comparison isolates retained
full-witness lifetime (visible fields as well as skipped payloads) because the owner operation, warmup count, full oracle calls,
and receipt cadence are otherwise the same. Strict-fresh-1 tests whether a
long-lived child is required for the observed trajectory. The allocation
0-versus-3 comparison measures the combined effect of validated warmup and allocator
state on the next measured region without treating the oracle as measured
owner work. None of these comparisons identifies why a candidate is faster or
slower; they only say whether the result changes when a named lifecycle factor
is controlled.

The 50-sample arms must retain the complete ordered sequence and report the
original full-window p50, mean, p95, p99, and maximum. Also report first-ten,
middle-thirty, and last-ten windows as descriptive diagnostics, with explicit
overlap and non-independence labels. Do not select a favorable window, drop
warmup-adjacent samples, or replace the nine process-pair bootstrap with
within-process samples. A candidate can be considered only if the full-window
decision and both fixtures remain acceptable under the existing 5% review
rules; lifecycle arms explain sensitivity and do not weaken that gate.

## Confounders and controls

| confounder | control and evidence |
| --- | --- |
| hidden semantic witness accumulation | strict-drained-50 and strict-retained-50 execute the same full oracle, then differ in whether the full witness, including skipped payloads, survives to the next sample |
| unvalidated warmup transition | validate every warmup in all strict arms; retain the exact old behavior as the contemporaneous legacy control, and keep its full-window review gate |
| allocator/cache state changed by validation | record oracle work outside the owner timer, drain before the next sample, and compare fresh-child and 0/3-warmup arms |
| process startup, CPU frequency, and host load | keep CPU 12 pinning and serial execution, record process order and environment, rotate variant order, and use independent process pairs |
| fixture shape and output size | run both already-qualified fixtures with identical operation and oracle scope; do not generalize from one deck |
| final JSON serialization | keep final serialization after the last measured sample; in the drain arm use equal receipt construction for both variants and record its scope outside `whole_ns` |
| allocator instrumentation | keep native and allocation binaries as separate lanes; do not combine allocation counters with native timing or call allocation counters RSS |
| quantile/window selection | retain all samples, keep the original full window as the gate, and label overlapping order windows descriptive only |
| changing correctness scope to make a result pass | bind the same source, expected output, stream, directory, raw-CFB, live-record, survivor, text, outline, notes, comment, and corruption-control requirements; any false nested oracle witness rejects the process |

No implementation or measurement from this design is present in this file's
scope. It is ready to guide a later probe revision without weakening the
preservation contract or assigning a causal explanation to the 0735 regression.
