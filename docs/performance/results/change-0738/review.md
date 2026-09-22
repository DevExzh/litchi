# 0738 retained-Sample layout review

## Disposition

Pre-capture review: **pass, with the measurement and production restrictions
below**. The 0738 probe is suitable for the planned unchanged-owner
qualification. This review does not authorize a production edit, candidate
reinstatement, or an ordinary-path performance claim.

## Source comparison

The comparison is against the sealed 0737 controls source at
`change-0738/prior-probe/src/lib.rs` and the independently rebuilt 0735 source
at `change-0738/archive-probe/src/lib.rs`. The source diff in
`change-0738/probe/src/lib.rs` is limited to the intended retained-sample
metadata change and its tests:

* `Sample` has exactly the 0735 field sequence and types: `index`, optional
  `phase_ns`, `output_sha256`, `output_inventory`, `oracle`, and optional
  `allocations`. It no longer stores `retained_witness_count`. Its derives are
  unchanged (`Clone`, `Debug`, and `Serialize`).
* `SampleReceipt<'a>` flattens a borrowed `Sample` and adds
  `retained_witness_count` only while producing a receipt. The custom
  `SampleReceipts::Legacy` serializer emits that same field as `sample.index +
  1`; strict receipts receive the already-derived lifecycle count explicitly.
* The strict warmup path passes count zero. The measured strict path passes the
  same count computed before 0738 from the retained-witness or drained state.
  The top-level count remains derived after the loop from the collection length,
  so the 0737 lifecycle metadata and count values remain present at the packet
  boundary.
* The added unit tests verify that the retained payload has no count, that the
  receipt adds one, that legacy and strict receipt arrays have the same visible
  value shape, and that full-oracle rejection still occurs before receipt
  construction.

The archived source is byte-identical to the recorded 0735 probe, and the
prior source is byte-identical to the recorded 0737 probe according to
`reference-source-proof.json`. The restored source has no change to
`public_format_edit`, `measured_public_format`, `timed_format`,
`execute_operation`, `oracle_for_output`, or the allocation-region bodies.
`output_sample` still computes the output digest and inventory and calls the
same full oracle with the same arguments; it only stops carrying the metadata
field into the returned `Sample`.

`source-equivalence.json` independently reports identical source blocks for
`Sample`, `output_sample`, `oracle_for_output`, `public_format_edit`,
`measured_public_format`, and `timed_format`. Archive and prior builds
(`build-0` and `build-1`) passed fmt, library tests, clippy, docs, and binary
build. The first restored build attempt is retained as `build-2` and failed
only because its copied lockfile still named the predecessor package; after
the lockfile package-name correction, `build-3` passed all five gates and is
the usable restored build. This custody-preserves the failed attempt without
weakening the final gate.

## Wire-schema and count checks

The wrapper preserves the 0737 JSON object shape for strict warmup receipts,
strict measured receipts, and legacy measured samples. The visible count is
still present in each strict receipt, each legacy sample, and the top-level
control report. The wrapper is outside the timed owner call and outside the
allocation region. The retained `Sample` therefore has the 0735 data shape
while the report retains the 0737 evidence schema.

The count rules are explicit and bounded:

* warmup receipts carry `retained_witness_count = 0`;
* legacy measured sample `i` carries `i + 1`;
* strict-retained measured sample `i` carries the pre-insertion retained count
  plus one;
* strict-drained measured samples carry zero; and
* the top-level count is the final retained collection length, or zero for
  strict-drained.

`contract.py` checks these values, requires the exact sample and receipt field
sets, and compares every visible oracle projection with the sealed 0735
reference. The source tests cover the serializer boundary; the qualification,
preflight, independent audit, and corruption controls must still be run over
the restored binary before any capture result is accepted.

## Measurement boundaries and limitations

The change is outside `measured_public_format` and `timed_format`; it does not
add receipt serialization, oracle work, or count derivation to `whole_ns`.
The allocation wrappers and their after-snapshot ownership remain unchanged.
Both strict arms continue to reserve the same witness capacity, execute the
same full oracle cadence, and differ in full-witness retention only as defined
by the 0737 harness. The layout comparison still changes compilation and
executable placement, so a timing difference cannot uniquely identify Rust
struct layout, generated code, allocator behavior, or instruction-cache
effects. The archive/prior/restored A/A contrasts are controls for that
combined observer/build/layout condition, not a proof of one cause.

The planned 72 native processes (four arms, two fixtures, nine rounds, 50
samples and three warmups) and 24 allocation processes (four arms, two
fixtures, three rounds, one sample and zero warmups) are appropriate only with
serial CPU-12 execution, rotated order, complete raw captures, paired
process-level comparisons, and the declared 5% full-window gate. Ordered
windows remain descriptive and non-independent. Allocation counters remain
boundary-relative owner counters and do not establish RSS, heap size, or a
latency cause.

No capture result from this review may be transferred to production or used to
reopen the rejected 0735 candidate without a separately frozen candidate
packet and the existing preservation oracle.
