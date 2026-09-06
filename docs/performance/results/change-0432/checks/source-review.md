# Change 0432 source review

This is a read-only review of the XLSX streaming writer, the
`xlsx_streaming_create` harness, the process/allocator observation boundaries,
and the evidence-bundle verifiers at the frozen source revision. No build,
test, capture, verifier, replay, or CPU workload was run for this review.

## Scope and source findings

The production path remains the restricted one-sheet, inline-scalar
`StreamingWorkbookWriter` in
[`crates/litchi-xlsx/src/streaming.rs`](../../../../crates/litchi-xlsx/src/streaming.rs).
It preallocates one reusable row buffer, publishes through an owned non-seek
sink, and charges finite row, cell, output, object, work, cancellation, and
scratch limits. The existing writer tests cover scalar round-trip and compact
six-member topology, strict row/cell ordering, XML escaping and invalid
characters, finite-number and error-value rejection, exact and one-over
limits, cancellation, short/interrupted/zero/failing sinks, typed partial
output, deterministic output, and fixed scratch capacity. The source review
found no production writer regression in this change.

`write_streaming_xlsx` includes context and writer construction, deterministic
row-text allocation, XML encoding, compression, package finalization, and all
logical writes to `HashingDiscardSink`. The complete `Vec` artifact,
reopen/readback, archive-member extraction, digest finalization, and summary
validation are outside that interval. Sink construction is before both the
timer and resource snapshots. The allocation region starts before the timer
and finishes after it, so the observed allocator interval includes all writer
callbacks through writer return while observer and procfs boundary work stays
outside elapsed time.

The change correctly enables the same resource envelope for XLSX and RTF by
using `Case::uses_streaming_creation()`. The ordinary binary has process
observations and an explicit unavailable allocation provider; the allocator
binary can publish measured v3 samples. Resource observations retain only
post-warmup samples and are aligned by the existing elapsed-then-sample-index
ordering.

## Memory interpretation

`allocation_metrics::Sample` publishes absolute values from the wrapped
`std::alloc::System` callbacks:

* `live_bytes_before` and `live_bytes_after` are absolute process live-byte
  counters at the region boundaries;
* `region_peak_live_bytes` is the absolute callback-order high-water value
  while the region is active, initialized to the entry live value; and
* `peak_live_bytes_before/after` are process-lifetime allocator high-water
  values.

For a size comparison, the operation transient increment is therefore the
checked derived value `region_peak_live_bytes - live_bytes_before`, and the
retained change is the signed comparison of `live_bytes_after` with
`live_bytes_before`. Raw region peaks must not be compared across fresh
processes as if they were operation-local allocations. The region includes
pre-existing live bytes and callbacks from other threads or background work;
it excludes allocator-internal realloc overlap, allocator metadata, other
unwrapped allocators, and physical RSS. It measures requested live bytes seen
by the wrapped global allocator, so it does not prove total process memory or
a language-level resident-memory bound.

The writer's 4 KiB reservation and reusable row buffer are a modeled
authoring window. A flat derived allocator increment across 64, 8,192, and
131,072 rows would be useful empirical evidence for this narrow generated
workload. It cannot establish a general bound for arbitrary text, workbook
features, append/repackaging, compressor internals, or total heap. The
allocator region also sees the harness's pre-existing corpus and retained
sample state. In particular, each completed sample's hexadecimal digest is
created after the region and retained in `digests`, so later samples have a
larger absolute entry baseline. Subtracting the entry live value is required;
an absolute peak trend would be a harness-retention artifact.

## Findings and resolution

The following findings were recorded during the initial source review. Their
resolution is included here so the review history remains auditable.

1. **Protocol wording versus emitted fields — resolved at the evidence
   interpretation layer.** The frozen protocol still names the operation
   peak-minus-entry and exit-minus-entry quantities, while reports emit the
   absolute fields. `measurement-clarifications.md` now explicitly defines
   the emitted fields as absolute and the checked operation values as derived
   deltas. The protocol remains frozen; no source or protocol mutation is
   needed for this batch.

2. **XLSX oracle width — resolved in the frozen harness source.** At commit
   `4c9cfd1fc`, `verify_streaming_xlsx` requires the exact six-member package
   set, one `Sheet1`, the exact row count, the four expected values in each
   row, and an exact visited-cell count while scanning the complete worksheet
   coordinate range. The sparse worksheet iterator does not materialize the
   dense Excel rectangle. The focused oracle test injects both an `E1` cell
   and an extra member and checks for the corresponding coordinate/member
   errors. Digest equality remains a determinism check, while the explicit
   structure and all-stored-cell checks provide the semantic oracle.

3. **Normal-versus-allocator latency label — resolved by usage constraints.**
   The producer retains the generic `comparable_timed_operation` schema label,
   but `measurement-clarifications.md` and `derive.py` treat latency as
   descriptive and compare repeat flags only within the same mode and shape.
   No normal-versus-allocator latency conclusion is permitted. Allocation
   vectors and sink counters remain independently useful.

These resolutions leave no production or verifier blocker. The report
verifier should continue to require, for allocator lanes, measured status,
30 values for every allocation field, zero failed allocation calls, finite
aligned vectors, and checked live/peak inequalities. For normal lanes it
should require allocation absence or explicit unavailable status and should
never silently turn unavailable values into zero.

## Evidence custody and portable checks

The source review of the bundle tooling found the following controls:

* `verify-report.py` rejects duplicate JSON keys, non-finite values, unknown
  schema keys, wrong member sets, wrong stored-cell counts, malformed metric
  vectors, failed allocations, invalid live-byte arithmetic, invalid peak
  inequalities, output/catalog mismatches, and unbound binary/source
  identities.
* `verify.py` binds the source manifest, protocol, verifier, capture drivers,
  build receipts, twelve lane receipts, reports, catalogs, logs, resource
  files, and the complete inventory. It checks exact artifact sets and reruns
  the report verifier for each report.
* `replay.py` binds all replay drivers and the sealed inventory, runs the
  portable verifier outside the bundle, and writes its terminal receipt only
  after the replay. Inventory and driver-hash checks therefore cover changes
  made during replay.
* The portable mutation cases target an allocator peak, the output digest,
  and the verifier itself. Each mutation is required to fail for its specific
  custody reason. These controls do not depend on an additional source test.

The source manifest intentionally covers the Rust/Cargo inputs; the Python
drivers and evidence documents are bound separately by their hashes and the
sealed inventory. This is an evidence bundle contract rather than a claim of
an independently hermetic build environment.

Root execution clarification: the initial verifier required a false dirty
flag, but the producer includes untracked files in that flag. Before formal
capture, the verifier was corrected to require a boolean and delegate source
cleanliness to outer custody. `capture-state.json` proves the tracked tree was
clean before and after; the only untracked entries were the user goal and this
evidence bundle. All 12 formal reports retain the honest `true` flag and passed
the corrected, hash-bound verifier with unchanged full code manifests.

## Tracked README wording held for post-capture update

The tracked `tools/perf-baseline/README.md` should use wording equivalent to
the following after capture is complete and the worktree can be changed:

> `region_peak_live_bytes` is the absolute requested-live-byte high-water
> observed through wrapped global `std::alloc::System` callbacks during the
> active region. It includes entry live bytes and callback activity from
> other threads or background work; it excludes allocator-internal realloc
> overlap, allocator metadata, other unwrapped allocators, and physical RSS.
> The derived operation transient peak is
> `region_peak_live_bytes - live_bytes_before`; `peak_rss_bytes` remains a
> process-lifetime high-water value.

This proposed wording is recorded here because the tracked README is frozen
while capture custody is being established.

## Adversarial coverage and ADR applicability

The writer's focused test suite covers the relevant sink, budget,
cancellation, ordering, escaping, and scratch-reuse failures. The harness
oracle now has explicit regression coverage for an out-of-range `E1` cell and
an extra package member. The portable verifier mutations cover a changed
allocator peak and changed output digest, while the verifier-source mutation
covers tool custody. No broad feature or malformed-workbook matrix is
required for this fresh one-sheet scope.

* ADR 0005 requires explicit budgets, forward-only row consumption,
  sequential non-seek output, and measurement claims scoped to observed
  resources. The writer and timing boundary comply; the 4 KiB window must
  not be promoted to a total-memory claim.
* ADR 0006 requires validation and fail-closed typed errors. Existing writer
  tests and the deterministic reopen oracle preserve this boundary; the
  oracle-width finding is now closed by the exact package and stored-cell
  checks.
* ADR 0008 requires compile/test/evidence gates and forbids turning
  functional tests into unsupported performance claims. The capture protocol
  is descriptive and tool-only; allocator and verifier receipts remain bound
  to source and protocol identity.
* ADRs 0001/0002/0010/0011/0024 keep the production writer and physical OPC
  ownership in their existing crates and keep instrumentation in the
  standalone performance tool. This batch introduces no ownership or
  dependency-direction change.

**Disposition:** no production-code or verifier-logic blocker found. The
capture custody condition is resolved as recorded above. Root applies the
proposed README wording after capture. This source review itself makes no
measurement claim; final results are retained separately.

## Root terminal execution

The twelve formal processes and final derivation passed. Both precleanup and
post-cleanup portable replays passed isolated export and all three deliberate
mutation checks. Task binary cleanup passed with retained hashes. The final
required-pass policy includes those terminal gates. The measured tables and
repeat variability remain in the result documents; this source review does
not expand their claim scope.
