# Change 0291: OPC ZIP operation-accounting evidence
Date: 2026-08-27

Status: additive machine-readable evidence; `performance_claim: none`

## Decision

The existing opt-in `opc_source_overlay_one_part_save` runner may publish an
optional `operation_metrics.opc_zip` object. This object exposes the complete
`litchi_opc::OpcOperationAccounting` report for the singular one-part overlay
operation. It is evidence of the bounded operation boundary only. It does not
change the existing `SourceMetrics` fields or reinterpret their generic
`ReadAt` scopes as ZIP counters.

The object contains one aligned `MetricVector` for each of these ten counters:

| Counter | Meaning |
| --- | --- |
| `compressed_deflate_payload_bytes_read` | Compressed Deflate payload bytes read by the low-level source operation. |
| `stored_payload_bytes_read` | Stored payload bytes read by the low-level source operation. |
| `stored_payload_bytes_accepted` | Stored payload bytes accepted by the low-level operation. |
| `deflate_bytes_produced` | Bytes produced by Deflate decoding. |
| `deflate_bytes_accepted` | Decoded Deflate bytes accepted by the low-level operation. |
| `generated_deflate_payload_bytes_emitted` | Newly generated Deflate payload bytes emitted by the writer. |
| `stored_payload_bytes_emitted` | Stored payload bytes emitted by the writer. |
| `precompressed_payload_bytes_emitted` | Precompressed payload bytes emitted without generating a new Deflate payload. |
| `raw_unchanged_source_bytes_accepted` | Unchanged source bytes accepted by the output sink. |
| `output_bytes_accepted` | Total output bytes accepted by the output sink. |

The counters are caller-owned accounting values. They are not inferred from
archive length, compressed member metadata, requested write lengths, or a
process I/O counter. `output_bytes_accepted` is the checked total reported by
the OPC operation and is not a replacement for any of the component counters.

## Status and alignment contract

On a successful accounting call, the `opc_zip.status` and all ten counter
vectors have status `measured`. Every vector contains exactly one `u64` value
for each retained sample, including a measured `0` when that counter was not
exercised by the operation. A measured zero must not be changed to
`not_applicable` or omitted.

For `not_applicable`, `unavailable`, or `overflow`, the corresponding vector
omits `values` entirely. JSON `null` is not an alternative representation for
an omitted vector, and an unavailable or overflowed value must not be
fabricated as zero. A failed or freshness-invalid operation is not published
as a successful measured sample; partial counters remain correctness evidence,
not a successful performance vector.

The runner records only retained samples, excluding warmups. The values are
reordered by the same `(elapsed_ns, original sample index)` ordering used by
the existing operation-metrics envelope. The serialized
`operation_metrics.sample_indices` is the exact permutation for that
`elapsed_ns.samples` order, and
`operation_metrics.alignment` remains
`elapsed_ns.samples_by_elapsed_then_sample_index`. Sample cardinality must
match across elapsed samples, sample indices, and every measured counter
vector.

## Scope and provenance

The stable group and vector scope is:

```text
opc_source_backed_package_write_part_overlay_to_stream_with_accounting
```

The operation-metrics latency token is:

```text
evidence_only_opc_source_overlay_accounting
```

The comparator allowlists both values, requires the group and every vector to
use the stable scope, and treats the token as evidence-only. Result `case`,
corpus manifest, binary identity, and the existing report environment provide
the remaining provenance. Dynamic member paths are not encoded as metric
scopes.

The accounting boundary follows Change 0290. Low-level ZIP work is owned by
the operation that performs it: a cache loader owns physical work, while a
cache hit or same-Part waiter does not duplicate that work in the caller's
report. Source freshness failures, partial sink output, cancellation, and
typed operation errors retain their 0290 semantics and do not become stale or
successful measured samples. The singular overlay separately accounts for
unchanged source publication, generated Store or Deflate payloads, and total
sink acceptance.

## Compatibility and comparison policy

This is an additive optional field in the current operation-metrics envelope.
The report `schema_version` remains `1`; no policy schema, default selector,
default result count, or claim-registry entry changes. Existing legacy
operation-metrics key validation remains a separate historical path.

The machine-readable comparator fails closed on unknown or missing OPC keys,
invalid status values, status/value mismatches, `values: null`, wrong sample
cardinality, non-integer or negative values, scope mismatches, and baseline or
current vector-path/status mismatches. The default performance policy does not
select these counters for regression comparison.

`performance_claim: none` is intentional. This change makes no latency,
throughput, speedup, physical-I/O, allocator/RSS, decompression,
recompression, copy, or zero-copy claim. The evidence-only latency token must
not be used to compare elapsed-time performance.

## Deferred surfaces

This change does not add machine-readable runners for
`PartView::data_with_accounting` or
`SourceArtifact::write_to_stream_with_accounting`. It also defers eager,
bulk, parallel, topology, batch, cache-wide, multi-Part, and `PartWriter`
accounting surfaces. The corresponding library accounting semantics,
including ownership, freshness, partial-output, and overflow behavior, remain
covered by Change 0290.
