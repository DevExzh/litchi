# 0557 allocation-scope amendment

Status: harness-only design and implementation; no performance result or
adoption claim.

The 0556 plan identifies the existing `commit_allocation_metrics` interval as
too broad for provenance attribution because it starts before the
`edit.set` loop. This amendment publishes two additional intervals for the
source-backed XLSX cell-values workflow:

* `staging_allocation_metrics` covers the existing operation region from its
  original start through the final `edit.set`.
* `commit_core_allocation_metrics` covers the next interval from that boundary
  through the returned `edit.commit()` value.

The existing `commit_allocation_metrics` field remains the combined interval
from before the first `edit.set` through the same post-commit boundary. Its
meaning, counters, endpoints, and region peak remain compatible with earlier
reports.

## Boundary design

The two new samples use one `allocation_metrics::Region`. `Region::split`
takes the counter snapshot and current region peak while holding the existing
observer mutex, then rebases the observer peak to the current live byte count
without releasing the region token. The first `split` emits the staging sample;
`edit.commit()` runs while the same token remains held; the final
`finish_split` emits the commit-core sample while also consuming the outer
region atomically. This prevents callbacks on another thread from falling
between two independently acquired regions or between the core and combined
endpoints, and does not create nested regions.

The region retains its original before snapshot and the maximum peak from all
split segments. `finish` still computes the combined sample from the original
before snapshot and final snapshot. Its combined peak is the checked maximum
of the segment peaks and the final segment peak, which is the same callback
ordered peak an unsplit region would have recorded. Counter differences remain
absolute checked differences over the full interval; split samples are
diagnostic partitions and are not added to inclusive profile rows.

The measured commit-core endpoint uses `Region::finish_split`, which captures
the final segment and consumes the outer region under one observer lock. Thus
there is no callback gap between the core sample and the combined sample;
their counters and live endpoints can be checked as an exact partition.

Observer invalidity and arithmetic overflow remain fail-closed. A disabled
normal binary publishes the explicit unavailable envelope for all three
commit-related fields; the allocator binary publishes measured samples,
including measured zero counters where an interval has no allocation activity.
If a split boundary cannot provide its observer peak, the active region records
sticky overflow state. Every later split and both final samples remain
`overflow` with numeric fields omitted, even if a later boundary has a valid
peak; an incomplete early segment cannot be repaired by subsequent evidence.
Observer mutex poisoning or callback reentry remains sticky `unavailable`
state. A sticky overflow still consumes the observer boundary during
`Region::finish`, so an error path cannot strand the active-region token. The
summary builder accepts an optional allocation vector only when it is present
for every iteration, while an evidence kind that is absent on every iteration
remains intentionally omitted (the eager controls have no region).

The native `commit_ns` timer still starts immediately before the existing
`edit.set` loop and is read immediately after `edit.commit()` returns; neither
clock boundary moves. The staging split remains inside that existing timed
end-to-end interval, immediately after `edit.set`, while the final
`finish_split` follows the existing elapsed-clock read. These are harness
boundaries around the same workflow, and production crates and public APIs are
unchanged.

## ADR and source bindings

Before editing, all 30 accepted ADR/index files in
`docs/performance/results/change-0555/adr-manifest.json` were rehashed and
matched the retained manifest. The 0556 plan and preparation packet remain the
applicable measurement inputs; their hashes at implementation time were:

```text
cf3614a930b0e9460b55f92fc783dca5d928f860959066948f7f0384143cd77  measurement-plan.md
152f962f94f91a5ac5b465ef970981b92410bef6fecf86058de2678f02f69539  preparation.json
5e7c67a28dafa62d2bdd9f428cd89467b85a15a2069c74c98de722212f058ffa  adr-compliance.md
```

The implementation is based on revision
`0d812aca92b608f9de8f9398fad1c379077ce708`. Changed files in this amendment
are:

* `tools/perf-baseline/src/allocation_metrics.rs` — add the non-nested,
  mutex-atomic region split and preserve combined-region peak accounting;
* `tools/perf-baseline/src/lib.rs` — add the two XLSX evidence vectors,
  partition the existing region at staging and commit boundaries, and retain
  the original combined field and timer scope, with explicit per-vector
  cardinality checks;
* `tools/perf-baseline/tests/xlsx_planning_allocations.rs` — verify allocator
  and normal reports, serialized field shape, phase alignment, and exact
  reconstruction of combined counters/endpoints/peaks from the two measured
  segments;
* `docs/performance/results/change-0557/alloc-design.md` — this design and
  source list.

No Rust build, test, benchmark, profile, or capture is performed by this
source-preparation agent. The root coordinator owns serial validation and
measurement after review.
