# Source and caller attribution review

All seven retained source files match the exact 0448 build source manifest.
Their copies live under source/ with original paths and line numbers. The current
ADR tree remains c950b6c8be822561b498d7bbe87c460873dcbf49. No Rust is edited.

## Timers and SHA

`pptx_provider_lifecycle.rs:1018` hashes the prevalidated expected output when
assembling the returned row, after all API timers and owner drops. This is an
untimed harness hash; the actual sink output was already compared byte-for-byte
at line 861. The hash does not execute inside CountingSink::write (`lib.rs:6990`).
The two production digest paths run under plan/publish API frames. Their shared
`digest_touched` (`source_cross_copy.rs:5227`) hashes metadata plus prepared image
and chart bytes using cancellable 64 KiB chunks (line 5551). Full caller stacks
are retained, including the remaining graph-digest SHA samples. Caller ancestry
partitions sampled work; it does not isolate exact retained timing intervals.

## Source reads versus logical rereads

The fixture has eight 2 MiB images (`lib.rs:255/261`) in each package, generated
as deterministic incompressible payloads (`lib.rs:14658`). Both packages retain
those media members. Cache limits are 64 MiB/128 entries (`pptx_cache_retention.rs:20`),
with a 512 MiB managed memory budget. The value one in managed_context is the
worker count, not a one-byte cache limit.

Planning calls prepare once (`source_cross_copy.rs:306`) and reads image payloads
through PartView::data at line 1629. Publication reruns prepare with the retained
plan at line 365, verifies the candidate at line 377 and builds topology at line
380. Logical image accesses in both prepare and verify_candidate use the source
package cache. Every media-rich publication sample records 23 source-cache hits,
zero cold loads and unchanged 16,807,458 retained bytes. Consequently those
logical accesses must not be counted as repeated fresh source payload reads.

Topology conversion authorizes compressed images at line 500. OPC's
`authorize_precompressed` (`source_backed.rs:4650`) captures bounded source
compressed spans, decodes/compares expected bytes and checks CRC plus source,
execution and budget fences. It reserves capture/writer staging and retains
source lineage/version. Later destination publication raw-copies unchanged
members. `soapberry-zip/src/preserve.rs:21,711,2165` uses a 32 KiB stack copy
buffer and bounded read_exact_at/write_all_counted chunks.

The observed 256 source-publication requests in the 64 KiB histogram bucket and
512 destination-publication requests in the 32 KiB bucket agree with these paths.
This is static/counter corroboration, not exact per-member call attribution:
reports do not retain offsets, callsite labels, or member-specific read intervals.
Other small source/destination validation reads remain in the same phase.

## Ranked follow-up

1. Determine whether OPC/ZIP can authorize compressed transfer while performing
   the first cold payload decode, under an explicit combined memory budget and
   existing fresh-source fences. This could remove a source compressed pass;
   it requires a concrete ownership design, mutation/refusal tests, capture
   lifetime checks and matched before/after evidence. A format-layer bare-byte
   cache or skipping token verification is not justified by these counters.
2. Measure a bounded preservation copy-chunk change for the existing 32 KiB
   destination passthrough. At 200 us fixed request cost, halving 512 full chunks
   would remove 51.2 ms nominal delay, about 2.08% of minimum-service media-rich
   API p50. This is only a model arithmetic estimate; read splitting, write caps,
   additional stack memory and local/short-read cases need evidence.
3. Retain touched-digest correctness and cancellation. The observed SHA share
   after separating the harness is insufficient to justify SIMD or weakening
   the digest contract. Broader native/cold/scaling and CRUD work stays open.

No performance optimization is accepted in this attribution batch. The next
production investigation must start from the separated ownership paths above.
