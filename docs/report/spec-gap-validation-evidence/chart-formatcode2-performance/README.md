# `formatcode2` bounded allocator/runtime evidence

This directory contains the isolated profile harness for the shared
`litchi_drawingml::chart::extension::formatcode2` owner. It measures the
source-backed `Element` and `Attribute` contracts independently of the
production workspace dependency graph. The harness owns a process-local
counting allocator, while `/usr/bin/time -v` records whole-process RSS.

The matrix has 28 lanes: small and `MAX_XML_BYTES - 1` near-limit inputs for
each owner, crossed with `read`, `read_shared` from an `Arc<[u8]>` prepared
outside the timer, exact no-op serialization to an owned `Vec`, no-op
`write_to` serialization into a counting/hash sink, scalar edit, cheap typed
`clone`, and malformed rejection. Near-limit values are also close to
`MAX_VALUE_BYTES`; the attribute near-limit source fills ordinary bounded
attributes while keeping each individual value within its limit. The small
malformed inputs use an unpaired ST_Xstring surrogate escape. The near-limit
malformed inputs append one non-whitespace byte after an otherwise valid source
and remain within the source byte ceiling.

Every final lane run uses three fresh processes and twenty measured samples
after two warm-ups. Raw per-process JSON, allocator counters, incremental
peak-live deltas, `/usr/bin/time -v` receipts, source manifests, source
hashes, host/toolchain data, exact commands, and the recomputed report are
retained under `results/` after the owner is frozen.

The runner refuses to start until the owner has been reviewed and frozen:

```sh
PROFILE_FROZEN=1 \
  CARGO_TARGET_DIR=/var/tmp/litchi-chart-formatcode2-profile-target \
  sh docs/report/spec-gap-validation-evidence/chart-formatcode2-performance/run_profile.sh
```

The guarded run was sealed against the approved source epoch: the source
manifest is identical before and after the build, all retained source hashes
match, and the report was recomputed from the raw receipts. The report makes
only scoped absolute observations; it must not be turned into a before/after
speedup claim without a named baseline, corpus, machine, build, and metric.

The shared attribute helper measures a complete host start tag containing the
qualified chart attribute. Package owners remain responsible for parent
grammar and placement, as required by the shared DrawingML ownership boundary.

The `write_to` timing includes the nonallocating sink's byte count, FNV-1a
checksum, and write-call count. Those sink lanes provide source-bound raw
metrics and make no performance comparison claim against the Vec-returning
lanes.
