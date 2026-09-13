# 0557 allocation-scope review

Status: conditional source-bound review. This review made no Rust build, test,
benchmark, profile, or capture claim. It covers the harness amendment that
partitions the source-backed XLSX `edit.set`/`edit.commit()` interval.

The reviewed worktree was based on `0d812aca92b608f9de8f9398fad1c379077ce708`
and was inspected on 2026-09-13. The source bindings at review time were:

```text
fc1530f55aa68daff95cd500d840aeb09dec1ded6daf912d03f79f3c39d14122  tools/perf-baseline/src/allocation_metrics.rs
0f8f52a2a3652734257c5b6479f845337e0da089843d3954a03da8434621886b  tools/perf-baseline/src/lib.rs
c947077f9517a1d39003b4234e0eeff106652920b2aa615ab48884948bb88cf8  tools/perf-baseline/tests/xlsx_planning_allocations.rs
ad6ca165b755e6946cd117e70e15dbd2f7b880011ad912ad64ac85e853b6ab18  docs/performance/results/change-0557/alloc-design.md
c46a1e9708f6ee1200eddbeb5266e10d30fbf6acc2e5c9c76a68c0fd804f88fe  docs/performance/results/change-0557/plan.json
```

The design and plan now bind the correct base revision, the 0556 measurement
inputs, and a continuous nonnested split contract. That resolves the earlier
source-binding and plan ambiguity. The implementation remains harness-only;
the production XLSX crates and public library APIs are unchanged.

## Verified behavior

`Region::split` keeps the one active-region token while it snapshots and
rebases `ObserverState::region_peak_live_bytes` under the observer mutex. This
is the correct way to avoid a callback gap between staging and commit. The
final `Region::finish_split` boundary produces the commit-core and combined
samples together, so no callback can land between those two endpoints.

For a measured operation, the intended relationships are:

* staging starts at the original `before` snapshot and ends at the first split;
* commit-core starts at that split snapshot and ends at the final snapshot;
* the old combined sample starts at the original snapshot and ends at the
  final snapshot;
* each counter in the combined sample is the checked absolute difference over
  the full interval, and therefore equals the checked sum of the two phase
  differences when all three samples are measured;
* `live_bytes_after` for staging equals `live_bytes_before` for commit-core;
  combined live endpoints are the outer endpoints;
* `region_peak_live_bytes` for the combined sample is the maximum of the
  rebased staging and commit-core peaks, while process peak before/after keeps
  its historical global meaning.

The changed no-split path still calls `finish_sample` through `Region::finish`
and therefore retains the old combined counter and peak behavior. The helper
also now fails closed when a split region has a missing final peak: the
`(Some(prefix), None)` case becomes `Status::Overflow` through
`Sample::measured` instead of reusing the prefix peak. `finish_sample` is still
reachable from production harness call sites through `Region::finish`; it is
not dead code merely because the new XLSX path uses `finish_split`.

The counter implementation uses checked add/subtract and a sticky overflow
flag. It does not saturate counters into plausible values. A counter wrap,
live-byte underflow, peak regression, checked difference failure, or missing
region peak produces an overflow sample with numeric fields omitted. Observer
poisoning and callback reentry remain a separate sticky invalid state and
produce unavailable samples. This preserves the measured-zero versus
unavailable distinction: an enabled allocator region with no callbacks can
publish measured zero counters, while a disabled normal binary publishes the
two-field unavailable envelope. The existing combined field remains present
for the source-backed normal and allocator reports; the two new fields are
additive and are omitted where the evidence type has no allocation region.

`edit.set` and `edit.commit` errors still leave the active region to its drop
path, which releases the observer token without publishing partial evidence.
The successful path finishes the combined region before later diagnostics and
publication allocation regions, preserving the old operation boundary.

## Required follow-up before candidate capture

There is one source-level fail-closed hole in the current split state. The
`ActiveRegion` stores only `combined_peak_live_bytes`. If a split returns
`observer_valid == true` with `segment_peak_live_bytes == None`, the segment
sample is correctly `Overflow`, but the active state does not remember that
failure. A later final peak can then make `finish_split` publish measured
commit-core and combined samples, and `Region::finish` after such a split can
take the unsplit `finish_sample` path. The normal one-owner invariant should
keep the observer marker present, but the allocation observer is explicitly
required to fail closed when its evidence is uncertain. Track that a split was
performed and retain a sticky invalid/overflow state (with unavailable kept
distinct for observer invalidity); add tests that clear or otherwise invalidate
the segment marker before both `finish_split` and `finish`.

`split_region` also currently returns a missing peak without marking the
observer invalid or asserting the active marker. Make that invariant failure
flow into the same sticky overflow result rather than allowing a later phase
to repair it accidentally.

The new unit test proves ordinary counter sums, endpoint chaining, and peak
maximums. It does not cover the cases needed to close this boundary:

* a measured-zero staging or commit-core interval, including its start live
  value and region peak;
* staging allocation followed by deallocation before the split, proving the
  peak is retained across the rebase;
* missing first-segment and missing final-segment peaks, with both phase and
  combined samples remaining overflow/unavailable as appropriate;
* counter add overflow, live underflow, and checked-difference reversal in
  each phase, with no partial numeric fields;
* observer poison and callback reentry at a split boundary;
* repeated split/finish/drop paths and token release after `edit.set` or
  `edit.commit` errors.

The summary builder appends each new allocation vector conditionally, as it
does for the older allocation vectors. Source-backed execution currently
supplies all three options on every successful iteration, and the focused
test checks lengths. Add an explicit cardinality invariant (or a test that
constructs an absent option) before using these vectors for phase alignment,
so a future partial evidence path cannot silently produce shorter arrays.

There is also a documentation/JSON-policy edge to resolve before freezing the
capture contract. `plan.json` says eager controls with no allocation vector
have an explicit unavailable record, while the eager branch currently emits
no `XlsxCellValuesSourceSummary` evidence at all. Either amend the plan to
permit omission, or emit and test the typed unavailable record. For the
source-backed path, additive `staging_allocation_metrics` and
`commit_core_allocation_metrics` fields with `skip_serializing_if` are
compatible with existing consumers that ignore unknown fields; retain a
golden check that the old `commit_allocation_metrics` values and unavailable
shape are unchanged.

Finally, the first `split` executes after `edit.set` but before `edit.commit`
and therefore lies inside the unchanged native `commit_ns` timer. The design
documents this correctly, but an allocator-binary comparison will include the
new mutex/snapshot boundary overhead where the old allocator binary did not.
Keep that effect visible in the allocation-lane evidence and do not interpret
an adverse or improved `commit_ns` row as merge attribution without the
pre-registered scope and controls.

Disposition: ordinary measured split behavior is correct and preserves the
old combined interval, but candidate build/capture should wait for sticky
missing-peak handling (or an explicit invariant decision), the boundary tests
above, and the eager unavailable/omission policy to be bound. No performance
or adoption conclusion follows from this source review.
