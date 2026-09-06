# Change 0430 goal scope

**Disposition: attribution only; the global goal remains open.** Change 0430
captures frame-pointer profiles on the unchanged 0429 media-rich
source-backed lifecycle. It identifies a high-value candidate for safe
compressed-media transfer; it does not modify production code or establish an
optimization result.

## What the capture establishes

The retained [`profile-protocol.json`](profile-protocol.json) uses the
unchanged release binary, `cycles:u`, 499 Hz sampling, frame-pointer call
graphs, CPU 2, three warmups and 100 retained media-rich iterations for each
of the `bytes` and warm positional-`file` providers. The retained record logs
contain no lost-sample warning. Recovered samples are:

| Provider | Iteration samples recovered / total | Publication path | Deflate + `CountingSink` |
| --- | ---: | ---: | ---: |
| bytes | 14,050 / 16,674 | 12,520 | 11,692 |
| file | 14,046 / 16,673 | 12,541 | 11,712 |

The current derived intersection of the retained iteration samples with
`zlib_rs::deflate::algorithm::medium::deflate_medium` is 83.22% for `bytes`
and 83.38% for `file`. The nested call chains reach
`SourceBackedCrossSlideCopyPlan` preparation and the source-backed publication
writer. This is a recovered-sample stack intersection, not an elapsed phase
fraction, throughput, or bytes-per-cycle result; the report's symbolization
warnings also make “recovered” the appropriate scope.

The profile is useful because it separates the historical whole-process
profile's setup and verification costs from a repeatable media publication
hot path. It does not compare control and candidate binaries, and the bytes
and file runs are warm provider observations. See the [profile index](profile-index.json),
[receipts](checks/frame-pointer-profiles.json), and the current [hotspot
inventory](../../HOTSPOTS.md).

## Highest-ROI candidate and current boundary

The source-backed cross-copy path currently validates selected image and chart
parts by materializing `PartView::data()` during `prepare`. It retains the
decoded payload in `Arc<Vec<u8>>` inside the prepared image/chart records, then
runs the preparation again when publishing a plan and sends the logical bytes
through the OPC topology writer and a Deflate entry writer. The relevant
ownership and stale-source checks are in
`crates/litchi-pptx/src/presentation/source_cross_copy.rs`; publication is
owned by the OPC/ZIP writers. The profile makes this repeated decode,
validation and recompression path the strongest immediate ROI target among the
currently measured non-iWork work.

The smallest safe candidate is a private, bounded raw-entry transfer for
validated leaf media parts. A plan could retain a validated compressed payload
and its ZIP metadata, while publication rewrites the destination member name
and offsets without running Deflate again. Any entry that cannot satisfy the
full proof must use the existing decoded/recompressed fallback. This is a
design target only; no such transfer is implemented or measured by 0430.

## Constraints for a future implementation

The candidate must preserve the existing source-backed contract:

* Validate content type, leaf-part status, declared and actual uncompressed
  size, CRC and accepted compression metadata under the existing input,
  memory, output, work, object, cancellation and source-version limits before
  emitting output.
* Recheck source and destination lineages, topology, relationship remapping,
  collision handling and publication gates. A raw source entry must not bypass
  stale-source detection, source checks, or dependency-closure refusal.
* Let the archive owner correctly rewrite local and central headers, target
  names, offsets, descriptors, extra fields and ZIP32/ZIP64 boundaries. Keep
  encrypted, signed, malformed, unsupported, or otherwise unprovable entries
  on the existing typed-refusal or recompression path.
* Account for retained compressed bytes and validation work in the same
  hierarchical budgets. Preserve deterministic output, untouched-member
  preservation, exact semantic output, and the ordinary facade's ownership
  boundaries; do not expose archive implementation types through CRUD APIs.

The first acceptance experiment should be a matched control/candidate
source-backed media lifecycle using the existing output, topology, payload,
source-version, cancellation and budget oracles. It should separately report
raw-entry transfer bytes, decompressed validation bytes, recompressed bytes,
logical reads, allocation regions and CPU attribution. A candidate claim
requires retained raw reports, repeat policy, and portable replay; the 0430
profiles alone cannot authorize it.

## Other open goal work

Compressed-media transfer has the clearest measured ROI, but it is one scoped
publication path. The non-iWork goal still requires representative CRUD
coverage, native producer breadth, true cold and caller-range I/O, bounded
semantic streaming/append, allocator and physical-copy attribution, explicit
scaling, and remaining strict gates. The 0429 provider/native evidence and
warm-file/range boundaries do not answer those questions. See the [global
goal audit](../../GOAL_AUDIT.md) for the current disposition.
