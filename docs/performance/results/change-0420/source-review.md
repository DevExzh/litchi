# 0420 decoded-payload ownership review

This review uses the retained 0419 media-rich corpus and trace metadata. It is
an ownership accounting result, not a new workload or profiler run. The 0419
destination corpus is 16,814,664 archive bytes and 16,866,277 logical part
bytes (`runs/allocator/A1/.../report.json`); its destination archive digest is
`a46fc227c453cbabbe54d6ca35fcad1cf1292c9a11e8e500bb1e4ec6a2708a4d`, and its
source digest is
`830928aafdc3ec8a5995a0d84a82ea9e2acb7f190f45008ddc017e2edfbf684b`.

## What owns a decoded member

An eager OPC open creates one `Arc<Vec<u8>>` for each decompressed physical
member. `PackageReader::read_many_serial_shared` returns those arcs;
`SerializedPart.blob` moves each one into `PartFactory::load_shared`, and the
resulting `BlobPart` or `XmlPart` stores the same allocation.

The following holders are references to that allocation rather than additional
payload copies:

| Holder | Ownership relationship |
| --- | --- |
| `OpcPackage.parts[*]` (`Part`) | Owns one `Arc` handle to the decoded member. |
| `OpcPackage.source_xml_parts[*]` | For XML only, clones the part's `Arc`; it does not clone XML bytes. |
| `PreservationProvenance.parts[*].blob` | Clones the same part `Arc`; its relationship XML is separate small metadata. |
| `OpcPackage`/`Snapshot` clones | Clone maps and metadata while cloning payload `Arc`s. |
| `OpcPackage.source_archive` | Owns the raw ZIP `Arc<Vec<u8>>`, which is a separate allocation needed for exact-source preservation. |

`PreservationProvenance::from_package` indexes the raw archive through a
borrowed ZIP view. It does not copy the raw archive. A `CanonicalRelationshipsXml::Owned`
value can own a newly serialized relationships stream, but it is not a copy of
the part payload.

## The duplicate in the media-rich lifecycle

`pptx_cross_copy_media_payload` creates eight 2 MiB media parts in each input
archive. The source and destination payloads are equal by index, while their
archive digests and slide graphs differ:

```
8 * 2,097,152 = 16,777,216 bytes (16.777216 MB) decoded media per input
16 output media parts * 2,097,152 = 33,554,432 bytes (33.554432 MB)
```

`build_candidate` first clones the destination package and adds copied source
parts with `BlobPart::new_shared`; those staged media parts already alias the
source or destination package payloads. It then serializes the candidate and,
in the 0419 path, calls `OpcPackage::from_vec(serialized)`. Eager reopen
decompresses the 16 output media members into fresh allocations while the
source, destination, and staged candidate graph are still alive. Thus the
media-only upper bound for fresh decoded output storage that can be retained by
that reopen is 33,554,432 bytes. This is an operation-live bound, not a claim
about RSS, allocator request volume, or the process high-water mark.

The 0420 `from_vec_reusing_payloads` path still validates and decompresses the
new archive. It only replaces a newly decoded payload after matching donor URI,
content type, bytes, and a no-larger donor capacity. In the clean owned-source
cross-copy path, the donor candidate contains the source and destination arcs,
so all 16 unchanged/copy media payloads can remain shared after reopen. The
maximum retained decoded-media reduction for one reopened output is therefore
33.554432 MB for this corpus; the observed value can be lower if a capacity
check fails or the donor lifetime ends before the measurement snapshot. The
whole lifecycle can retain more than one candidate's payload set, as described
below.

## Why the lifecycle delta is larger

The 0420 allocator summary reports media `live_bytes_after` of 287,885,879
bytes for control and 237,452,746 bytes for the candidate in both pairings: a
matched delta of 50,433,133 bytes. The 33,554,432-byte figure above is only the
media payload bound for one reopened output package; it is not the bound for
the whole lifecycle.

The lifecycle prepares a candidate twice. The first preparation returns a
`CrossSlideCopyPlan` whose `Patch` stays alive through application, publication,
and the allocator region's end. `Patch::capture` stores `ResourceState.blob` arcs for every changed
part. In the baseline those arcs point to the first freshly reopened
candidate, so the plan retains the nine copied source-closure payloads. The
recorded `planned_bytes` for that closure is 16,784,397 bytes, including the
16,777,216 bytes of its eight media members. With payload reuse, those patch
arcs point back to the already opened source/destination donor storage instead
of retaining a separate first-candidate decode.

The known media portions of those two retained sets are:

```
first plan Patch:       8 * 2,097,152 = 16,777,216 bytes
application candidate: 16 * 2,097,152 = 33,554,432 bytes
known media total:                         50,331,648 bytes
```

The measured delta leaves `50,433,133 - 50,331,648 = 101,485` bytes of
non-media and allocator bookkeeping residual. The retained reports do not
expose per-part allocation lengths, so that residual is not assigned to a
particular XML or relationship holder. The source accounting therefore
explains the extra 16.777216 MB as the first plan's retained copied-source
media set, while preserving the 101,485-byte residual as unassigned evidence.
This accounts for the measured lifecycle observation without turning it into a
local peak or a general memory claim.

This constructor does not deduplicate the separately opened source and
destination allocations by content hash. A future cross-package dedupe could
remove at most one copy of the eight equal source/destination media payloads,
16,777,216 bytes, but that is a different ownership policy and is not part of
the 0420 donor experiment. Raw source and output ZIP allocations remain
required and are excluded from both bounds.

The retained 0419 Heaptrack reports locate `BoundedVecWriter` allocation
stacks. Their `-H` histograms are whole-command data because Heaptrack 1.5.0
applies filtering after histogram construction; they cannot provide a decoded
blob-only byte total. The 33.554432 MB figure above comes from the fixed corpus
member count and size plus the `build_candidate`/reopen ownership path, not
from treating the writer histogram as decoded storage.

## Matched follow-up protocol

Use the current harness and the frozen common flags in `protocol.json`, with
one fresh process per selector and leg, CPU affinity 2, one harness worker,
`RUSTUP_TOOLCHAIN=1.98.1`, and `--filesystem-cache warm`. Run control and
candidate in A1, B1, B2, A2 order (`control`, `candidate`, `candidate`,
`control`) for both:

```
normal:    --samples 100 --warmup 10
allocator: --samples 30  --warmup 3
```

The command shape is:

```
env RUSTUP_TOOLCHAIN=1.98.1 taskset -c 2 /usr/bin/time -v <binary> \
  --case pptx_cross_copy_media_rich_lifecycle \
  --shape many-small --payload compressible --writer-shape large \
  --xlsx-shape medium --xlsx-cell-crud-shape medium \
  --xlsx-row-visibility-shape medium --semantic-shape medium \
  --workers 1 --filesystem-cache warm \
  --samples <N> --warmup <W> --json <report> --corpus-manifest <catalog>
```

Repeat the same command for `pptx_cross_copy_plain_lifecycle`; use the normal
and allocator binaries for their respective lanes. The frozen protocol uses
the prior 0419 candidate Heaptrack capture as the profile basis: its production
and harness sources match this batch's control. No new Heaptrack capture is
planned for 0420. Matched allocator and RSS measurements test the retained
memory hypothesis; the prior whole-command trace does not prove a local peak.

Bind every report to the source revision, binary hash, protocol hash, corpus
source/destination hashes, output digest, plan topology, source immutability,
and all existing semantic/refusal gates. Compare operation `live_bytes_after`
and the process-scoped peak/RSS fields, while retaining allocator vectors. The
donor path should not be required to reduce allocator request volume because
it still performs the new archive's decompression. The 100-sample lane remains
diagnostic; do not make a latency claim or an allocator elapsed-time claim.
Report each ABBA pair and same-revision drift, and treat any retained-memory
change as evidence for this workload only. Sharing must not drop the raw
archive or transfer donor metadata/source authorization: the new serialized
archive must establish its own preservation provenance and exact-source
authority.
