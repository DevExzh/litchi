# 0810 driver and archive preflight review

This is a read-only review of the 0810 packet at base
`3677e31be5c9d5582a1f6d531ebb4d54db5a0acc`. It does not run Cargo, rustfmt,
the probe, a workload, Callgrind, or any packet driver. The review covers the
candidate archive, plan and policy, and the adapted build, capture, quality,
probe-quality, profile, application, and restoration drivers.

## Archive and source custody

The production notes codec is unchanged at review time. Its SHA-256 is
`8485d43c99f19bda9b3510c8323aa6bc2b117fb7df262372ce82868a5397239c`, equal to
`candidate/before/codec.rs`. The candidate manifest binds the after source and
patch hashes, the base commit, and the one-file production path. The patch
passes `git apply --check`; no production source change is present. The packet
also records the three unrelated working-tree file identities and keeps them
outside the source allowlist.

The archived after source is the reviewed direct-event candidate: it keeps the
buffered scanner and `inspect_element_oracle` byte-identical, preserves the
public `NamespaceResolver` scope transitions and error conversion, and adds
only the focused differential tests described by the manifest. The application
driver requires an accepted before-only qualification audit, verifies the
one-file manifest and patch identities, and checks the changed source census
after application. The restoration driver requires an explicit rejected
decision and restores the archived before bytes, then verifies the full source
census.

## Frozen protocol

The plan and policy match the requested 0810 protocol:

* six shapes crossed with capture, commit, and lifecycle (18 rows);
* before-only allocation qualification, one report and sample per row;
* six native paired blocks in `before/after`, `after/before`,
  `before/after`, `after/before`, `after/before`, `before/after` order, with
  30 samples and three warmups (216 reports and 6,480 samples);
* two allocation paired blocks in `before/after`, `after/before` order, with
  three samples and no warmup (72 reports and 216 samples);
* four large-capture owner-scoped Callgrind reports, one sample and no warmup;
* nearest-rank `ceil(n*p)-1` process quantiles, paired medians, bootstrap seed
  810810 with 10,000 resamples and sorted endpoints 250/9749;
* a benefit of at least 3% in one capture or lifecycle row with bootstrap high
  endpoint below one; any of the eighteen rows with paired p50 ratio above
  1.05 and bootstrap low endpoint above one vetoes adoption; and no
  per-allocation-block median may increase calls, allocated bytes, net live
  bytes, or peak above entry.

The qualification lane names the allocation binary explicitly. The probe
source and lockfile are inherited from the final repaired 0806 probe, while
the root Cargo.lock and rustfmt configuration are captured in the 0810 input
set. The packet excludes historical timing, cross-format, heaptrack, perf,
and iWork lanes. Callgrind remains diagnostic owner-scoped Ir evidence and is
not used as a latency or adoption gate.

## Driver checks

The serial build driver is bound to the 0810 schema, current base, one-file
allowlist, 9,196-file source census, and three fresh binary identities. The
quality driver has the six PPTX gates; the probe-quality driver has format,
36-test, and warning-denied Clippy gates. Both drivers retain source and
probe-input custody after every command. The capture driver binds every
receipt to its source and binary, preserves report/RSS/log artifacts, and
records the plan hash. The profile driver uses the exact owner
`namespace_uri_probe::capture_region_0793`, disables collection at start,
requires one numbered positive publication, and checks the source after each
process.

The capture driver now selects each non-qualification lane's declared
`lane_plan["orders"]`, including the allocation `AB`, `BA` schedule. The
independent raw audit below also enforces those receipt identities directly.

The adapted main analyzer now carries the 0810 schema, target, and seed and
independently asserts the qualification binary and both native and allocation
order fields. `root_audit.py` repeats those bindings from the raw report paths
as an independent numerical check.

## Readiness disposition

The archive, scope, source custody, candidate application boundary, quality
gates, and measurement policy are ready for the root-owned baseline build and
before-only qualification. The independent audit's order and binary checks
must remain in the final packet before native or allocation measurements are
accepted. No performance, correctness, or retention conclusion follows from
this static review.
