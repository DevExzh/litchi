# Change 0345: OPC source-backed reader ingress

Date: 2026-08-31

Status: implementation described; focused validation recorded

performance_claim: none

## Scope

This change records the bounded reader-ingress path for the public litchi-opc
SourceBackedPackage owner-layer API. No facade API,
iWork API, archive handle, reader handle, runtime, lock, or public source
lifetime was added.

The caller-provided input is consumed once. The bounded ingress retains the
compressed archive once as the source-backed package and does not create a
second complete compressed-archive copy. The input maximum is enforced with
the existing typed limit failure rather than truncating, flattening, or
reclassifying the input. ReadLimits and try_reserve_exact bound logical input
and local admission/allocation work; they do not bound total RSS or aggregate
memory across concurrent opens.

Relative to the eager path's compressed-plus-all-decompressed retention, this
path retains one compressed buffer plus indexed metadata and deferred selected
payloads. That is an unmanaged ownership/layout reduction, not a host-memory
budget or process-wide memory manager. Callers needing tighter host-memory
limits must supply a lower max_input_bytes, serialize opens, and account for
aggregate process memory externally.

## Structural evidence

The focused evidence covers the complete open/read boundary rather than a
latency benchmark:

- Opening the source-backed package performs catalog, content-type, and
  relationship work but produces zero cold ordinary-payload loads.
- Reading one selected ordinary part performs exactly one cold, successful
  payload load. The selected bytes are checked against the expected payload;
  the source-backed owner remains the authority for the retained archive.
- An input exactly at the configured maximum is admitted, while an input over
  that maximum returns the typed limit error; the overrun contract is asserted
  with actual = maximum + 1. No unbounded reader growth is implied by the
  test.
- The focused validation was 4/4 tests passing, including
  reader_ingress_retries_one_interrupted_read and
  reader_ingress_rejects_invalid_read_count_without_panicking, followed by
  four owner-library checks. The validation used one Cargo process with one
  build job and a
  dedicated on-disk target; no parallel rebuild or parallel test matrix was
  used.

These are structural and correctness observations. They do not measure
physical filesystem reads, decompressed bytes, copy volume, allocation
attribution, or cache behavior outside the selected source-backed operation.

## Limits and cancellation

The reader ingress is bounded by the existing ReadLimits policy, including
the exact maximum-input check and its typed error path. This bounds logical
input and reduces local admitted work; it does not bound total RSS or
aggregate memory when multiple opens run concurrently. A caller-supplied
arbitrary blocking Read remains a limitation: the API cannot asynchronously
interrupt a read that blocks inside the caller's implementation. Callers that
need a tighter host-memory envelope must use a lower max_input_bytes,
serialize opens, and account for aggregate process memory externally.

## Performance interpretation

performance_claim: none is intentional. No RSS measurement and no
before/after latency distribution were collected, so this record makes no
memory, allocation, throughput, speedup, cold-cache, or physical-I/O claim.
The single retained compressed archive and deferred ordinary payload loads
are ownership/read-behavior evidence only. The change does not alter the
facade's eager smart-detection result and does not cover iWork packages.
