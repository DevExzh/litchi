# Allocation-vector contract review

The packet changes only Python validation tooling and performance documentation.
Production Rust, benchmark Rust, dependencies, locks, and all 35 normative
inputs retain their origin hashes. Root owns all executions and Git mutations;
review/test agents inspect sources and prepare files only.

Source anchors:

- `tools/perf-baseline/src/operation_metrics.rs:227`: eleven allocation fields
  are `MetricVector`, not scalar counters.
- `tools/perf-baseline/src/operation_metrics.rs:637`: unavailable/overflow
  observations omit numeric vectors; measured observations emit aligned arrays.
- `tools/perf-baseline/src/allocation_metrics.rs:145`: raw statuses, scopes,
  absolute live bytes and lifetime peaks, region peak, and optional fields.
- `tools/perf-baseline/src/allocation_metrics.rs:225`: checked counter deltas,
  monotonic lifetime high-water marks, and region-peak bounds.
- `tools/perf-baseline/src/allocation_metrics.rs:699`: successful reallocations
  increment allocation calls, so reallocation calls cannot exceed allocations.

Independent static review found the helper consistent with these definitions.
The initial draft treated nonzero failed-allocation counts as a schema error;
review identified that these counts are valid data. The final helper accepts
them and leaves trial rejection to the caller's explicit policy. The first
positive preflight and its helper source are retained in `preflight-attempt-0`;
it was a passing draft, not a failed command. Final preflight uses the corrected
schema/policy separation and adds the reallocation-count invariant.

The helper validates single-case schema-v1 reports with v3 allocator vectors.
Native mode requires explicitly unavailable counters with no numeric payload;
observer mode requires measured counters. Overflow/unavailable observers cannot
satisfy a request for measured evidence. It checks exact sample alignment,
unsigned raw values, signed conservation, and region/lifetime peak bounds.
It neither derives latency summaries nor resets absolute live counters.

Boundary: this helper intentionally does not validate `source.ordinary_save`,
OPC/XML semantics, corpus/output hashes, process counters, binary hashes, or
comparative statistics. A report relabeled with a matching case name needs the
caller's separate source/corpus/artifact checks. The module and report document
this limitation; it is not a replacement for artifact admission or source custody.

Next trial also needs the two known prepared-reader corrections from 0825:
resolve the accepted admission attempt from its bound descriptor rather than
hardcoding `admission-{leg}`, and compare preservation replay as canonical JSON
so ZIP timestamp tuples match their serialized arrays. Neither abandoned 0825
reader nor frozen driver is modified in this packet.
