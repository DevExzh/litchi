# 0439 handoff for the append runner

The authoritative candidate constants are in `protocol.json`; the acceptance
steps are in `acceptance-plan.md`. This file is the short integration contract
for the append runner.

## Candidate `SourceSummary` contract

Add an opt-in `odp_append` summary, leaving every default case and
existing summary unchanged. It should report:

```text
role = "existing_append"
implementation = "litchi_odp::authoring::edit::Snapshot::from_bytes/transaction/add/commit"
shape = tiny | medium | large
source_slide_count = 64 | 4096 | 8192
output_slide_count = source_slide_count + 1
append_count = 1
corpus_generator = "litchi-odp-existing-append-lifecycle-v1"
source_archive_sha256 / output_archive_sha256
source_archive_bytes / output_archive_bytes
source_content_xml_sha256 / output_content_xml_sha256
source_content_xml_bytes / output_content_xml_bytes
source_semantic_sha256 / output_semantic_sha256
source_order_sha256 / output_order_sha256
source_text_projection_sha256 / output_text_projection_sha256
source_text_projection_bytes / output_text_projection_bytes
opaque_member_path = Opaque/litchi-perf-odp-existing-append-opaque.bin
opaque_bytes = 65536
opaque_sha256
opaque_member_compressed_bytes / opaque_member_compressed_sha256
source_members / output_members = six complete member identity records
source_member_count / output_member_count = 6
source_manifest_bindings_verified / output_manifest_bindings_verified
untouched_members_verified
opaque_member_compressed_identity_verified
source_semantic_reopen_verified / output_semantic_reopen_verified
append_exactly_one_verified
source_unchanged_verified
patch_replay_verified / inverse_patch_verified / stale_source_refusal_verified
exact_noop_verified
runtime_output_digest_verified / runtime_sink_length_verified
lifecycle_ns / output_sha256
text_contract
timing_scope
performance_claim
```

The corpus manifest remains the source identity. The top-level output SHA and
the sink identify the committed output. Do not overwrite `CorpusManifest`
source fields with output values.

## Clock and ownership

The source archive, expected output, input clone, append strings, and sink
setup are outside the timer. The timer starts immediately before
`Snapshot::from_bytes` and includes it, `transaction()`, exactly one `add`,
`commit`, and the sink write of the actual committed bytes. Stop the clock
before endpoint observations and before dropping the commit, snapshot,
transaction, sink, or output bytes. Reopen, hash, compare, and diagnostics are
outside the clock.

The sink must digest the bytes borrowed from the committed snapshot. Do not
write the expected output buffer. Do not claim source reads, bounded memory,
source backing, physical I/O, end-to-end throughput, or cancellation. An
operation-scoped committed-bytes-per-second value derived from accepted sink
bytes and the measured interval is allowed.

## Source/output gates

The independent oracle checks the six-member source/output set, exact mimetype,
the pinned styles/meta bytes, regenerated opaque bytes, and the exact five
manifest bindings. It reopens every slide and requires the source sequence
unchanged plus one exact tail slide. It computes the existing buffered semantic
digest, a canonical text projection, and a separate order digest itself.

The append title/body is `odp_buffered_title(source_slide_count)` and
`odp_buffered_body(source_slide_count)`. The opaque byte formula is in
`protocol.json`; do not replace it with a producer hash.

## Metrics

The formal matrix is 12 reports × 30 retained samples = 360 samples: normal and
allocator modes, three shapes, two repeats, three warmups. `sample_indices` and
all retained vectors use the existing elapsed-time alignment. Allocator mode
reports measured count/byte/live/peak vectors; normal mode may report allocator
unavailable. RSS scope is process-observation only.

The only facts still requiring public-API reconciliation are the accessor names
for committed bytes, any changed/non-noop commit predicate, the output sink
fields, and the writer's observed ZIP compression/order. If no public changed
predicate exists, use the independent output identity gate rather than
inventing a typed conflict or patch result. Reconcile those names without
weakening the gates or changing the timed scope.
