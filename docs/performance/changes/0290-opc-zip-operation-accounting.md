# OPC ZIP operation accounting

## Status

This is a production `litchi-opc` correctness and operation-diagnostic slice.
It exposes caller-owned accounting for three bounded source-backed operations:
cold single-Part reads, exact source-artifact publication, and the singular
existing-Part overlay publisher. It does not change the bytes accepted or
emitted by any operation.

```text
performance_claim: none
```

The report is logical operation accounting. It is not a benchmark, resource
profile, cache telemetry stream, or performance acceptance gate. No latency,
throughput, speedup, regression, memory, allocation, physical-I/O, or
zero-copy claim is made.

## Report ownership and shape

`litchi_opc::OpcOperationAccounting` is a caller-owned, copyable report. It
is never retained by an archive, source-backed package, cache, same-Part
flight, or publication plan. Its public accessors mirror the nine payload and
raw-source counters of `soapberry_zip::ZipOperationAccounting` without
exposing that low-level report type through the OPC API. The OPC report adds
one independently checked counter for every byte accepted by the caller's
sequential output sink.

The report contains no member names, Part URIs, ZIP entry IDs, source bytes,
payload bytes, or error strings. The operation method supplies the boundary;
the counter value itself does not retain a path, compression method, declared
size, or outcome. A caller must interpret it together with the operation it
passed the report to and the returned `Result`.

## Counter meanings

The counters remain separate because a ZIP declared size, a member payload,
decoder output, sink acceptance, and copied archive bytes are different
observations. Every value is a checked `u64` count of actual work observed or
accepted by the covered operation.

| Counter | Meaning | Boundary |
| --- | --- | --- |
| `compressed_deflate_payload_bytes_read` | Compressed Deflate member-payload bytes actually read while a cold Part is decoded. | The compressed payload range traversed by the operation; it excludes ZIP local headers, descriptors, central-directory records, end records, and comments, and is not substituted with declared metadata. |
| `stored_payload_bytes_read` | Stored member-payload bytes actually read while a cold Part is loaded. | The Store payload supplied by the ZIP source; it excludes ZIP framing and metadata. |
| `stored_payload_bytes_accepted` | Stored payload bytes accepted by the Part destination. | It can be lower than the source count on a failed or partial operation; a successful materialization normally makes the two counts equal. |
| `deflate_bytes_produced` | Uncompressed bytes produced by the Deflate decoder. | It records decoder output made available to the operation, including a partial prefix that can later be rejected by checksum, size, source, or cancellation validation. |
| `deflate_bytes_accepted` | Decoded Deflate bytes accepted by the Part destination. | It records destination acceptance and can be lower than produced bytes on a partial or failed operation. |
| `generated_deflate_payload_bytes_emitted` | Compressed member-payload bytes generated for a changed member during preservation publication. | Compressor output only; ZIP local headers, descriptors, central records, end records, and comments are excluded. |
| `stored_payload_bytes_emitted` | Member-payload bytes emitted for a newly generated Store member during preservation publication. | Store payload only; ZIP framing and metadata are excluded. |
| `precompressed_payload_bytes_emitted` | Caller-provided compressed member-payload bytes emitted without recompression. | The precompressed payload only; it excludes ZIP framing and metadata. The counter is mirrored for the low-level shape even when the narrow OPC operations do not use this path. |
| `raw_unchanged_source_bytes_accepted` | Unchanged source archive bytes accepted by exact publication or raw-member preservation. | Archive-byte accounting rather than payload accounting. Depending on the path, it includes unchanged local spans, central records, end records, descriptors, and the archive comment. It never infers the count from source length. |
| `output_bytes_accepted` | Every byte accepted by the caller's sequential output sink. | Includes all ZIP framing, generated records, payloads, unchanged source bytes, and comments that the sink accepts. It is checked independently from the nine mirrored counters. |

On a complete exact source copy, `raw_unchanged_source_bytes_accepted` and
`output_bytes_accepted` equal the actual bytes accepted by the sink. On a
changed overlay, raw unchanged bytes and generated payload counters remain
separate while `output_bytes_accepted` covers the complete accepted archive.
These relationships are observations of a particular operation, not general
identities for all ZIP paths.

## Covered operations

### Cold single-Part reads

`PartView::data_with_accounting` passes the caller's report to the bounded
cold-loader path. A loader or allocation-bypass path that performs the ZIP
read merges the low-level nine counters into the report. Cache hits and
same-Part flight waiters do no physical ZIP read and leave the caller's
report unchanged. The loader owns the physical accounting even when another
caller waits for its result.

A failed cold decode retains the counters observed before the failure and
does not publish the failed payload into the cache. A later retry performs
its own physical read and updates only the retry caller's report. A source
freshness failure after decoding but before cache publication returns
`OpcError::SourceChanged`; the already observed ZIP counters remain in the
caller-owned report. Freshness checks before and after the read prevent stale
payload publication.

### Exact source-artifact publication

`SourceArtifact::write_to_stream_with_accounting` copies the retained source
artifact through a sequential sink while recording actual accepted bytes.
Each short write contributes only the accepted prefix to both
`output_bytes_accepted` and `raw_unchanged_source_bytes_accepted`; source
length is never used as a proxy. The operation keeps its existing source
freshness, cancellation, output-budget, flush, and `IncompleteOutput`
semantics.

The method is not an atomic transaction. If a source, policy, cancellation,
flush, or sink error occurs after output begins, the report retains the
accepted prefix and the returned error retains the existing partial-output
count. A source freshness decision remains authoritative over a diagnostic
merge error.

### Singular existing-Part overlay

`SourceBackedPackage::write_part_overlay_to_stream_with_accounting` accounts
one existing Part replacement. It first follows the same cold-read ownership
rules as `PartView::data_with_accounting`.

An exact payload no-op selects exact source publication: the output is copied
byte-for-byte, raw unchanged and total accepted output are both measured, and
generated payload counters remain zero. A changed payload selects the bounded
preservation path: unchanged source bytes, generated Store or Deflate payload
bytes, and total accepted output are reported independently. Generated ZIP
framing is included only in `output_bytes_accepted`, not in a payload counter.

The publisher retains existing source-version checks before planning, during
source-backed publication, and before the final decision. A sink or source
failure after accepted bytes leaves the partial raw/output accounting in the
report and preserves the existing `IncompleteOutput`/`SourceChanged`
precedence. Accounting does not authorize a fallback to a full rewrite.

## Cache and budget ownership

`OpcOperationAccounting` is distinct from `SourceCacheDiagnostics`.
`SourceCacheDiagnostics` remains a content-free, package-lifetime diagnostic
of cache hits, cold loaders, waiters, retained state, and managed Budget
usage. ZIP operation counters are not ambient cache counters and are not
aggregated into that snapshot.

For a managed source-backed operation, the existing hierarchical Budget still
owns input, memory, work, object, and output reservations. ZIP accounting only
records logical payload and sink boundaries; it does not replace Budget
charges or claim physical I/O, allocator traffic, or retained-memory usage.

## Checked overflow and error semantics

Every OPC counter update uses checked addition. An unrepresentable update
returns the dedicated
`OpcError::OperationAccountingOverflow { counter }`; it is not silently
saturated, converted to a ZIP-format error, or represented by a fabricated
zero. A low-level ZIP report is merged in a fixed order across all nine
mirrored fields. Representable later fields are still merged, and the first
overflow is returned after the merge attempt.

The report may therefore contain a valid partial prefix when the operation
fails. The caller must use the returned `Result` to distinguish a complete
operation from a decode, source, limit, cancellation, sink, flush, or
accounting failure. Existing source freshness and incomplete-output errors
remain authoritative when they race with accounting finalization.

## Correctness evidence

The focused `crates/litchi-opc/tests/operation_accounting.rs` suite covers:

- cold Store and Deflate Part reads, including exact payload counters;
- unchanged cache-hit and same-flight waiter behavior;
- Store and Deflate CRC failures, retained partial counters, and retry-owned
  physical work;
- exact source publication with actual raw/output sink acceptance;
- singular changed overlays with separate raw and generated Store/Deflate
  payload counters;
- exact no-op overlays on cold and already-warm Part selections; and
- partial sink failures with accepted-prefix accounting.

The OPC accounting unit tests cover the low-level merge boundary and the
dedicated output-counter overflow. Existing source-backed tests continue to
cover source freshness, cancellation, managed output charging, cache flight
completion, and partial-output error behavior around these methods.

## Deferred work

This record intentionally does not cover or claim to solve:

- ordinary eager `OpcPackage` reads or writes and the uninstrumented
  `PackageWriter`/`PhysPkgWriter` paths;
- incremental `PartWriter`/`StreamingArchiveEntry` accounting;
- parallel, bulk, or aggregate multi-member operations;
- topology publishers, relationship-add/remove publication, or other
  topology-changing source operations;
- batch and multi-Part overlays;
- cache-wide or process-wide aggregation of ZIP counters;
- `tools/perf-baseline` `operation_metrics`/`SourceMetrics` schema changes,
  comparator identity rules, selectors, default cases, or ABBA evidence; or
- interpreting logical counters as physical I/O, decompression allocation,
  recompression volume, RSS, allocator, scheduler, or zero-copy evidence.

Those paths need separately reviewed operation boundaries and provenance. In
particular, the existing performance `SourceMetrics` fields for compressed,
decompressed, and recompressed bytes remain unchanged: the nine ZIP counters
have different dimensions and cannot be safely collapsed into those vectors.

## Explicit exclusions

No performance claim is made for:

- latency, throughput, speedup, regression, or benchmark stability;
- physical filesystem or device I/O, syscall counts, or remote-range traffic;
- allocations, deallocations, RSS, peak memory, or retained-cache memory;
- decompression or recompression work beyond the logical byte boundaries
  explicitly counted above; or
- the absence of copying, materialization, or parallel work.

`performance_claim: none` is authoritative for this OPC propagation slice.
