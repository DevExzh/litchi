# Repeated audit investigation

This source-path investigation guides profiling. It is not an implemented
optimization or a performance result. The ordinary changed DOCX tail-stream
lifecycle decodes the main part six times:

| Stage | Decoded source passes | Authored passes | Obligation |
| --- | ---: | ---: | --- |
| DOCX source scan | 1 | 0 | Story grammar, insertion position, source hash/length, semantic counts and bounded scanner policy |
| Authored sealing | 0 | 1 | Authored event and encoded-byte proof, bounded producer/cursor contract |
| DOCX candidate scan | 1 | 1 | Candidate story grammar, exact insertion and paragraph counts, source/candidate hashes |
| OPC source audit | 1 | 0 | Independent source XML validity and policy, source hash/length |
| OPC candidate audit | 1 | 1 | Independent candidate XML validity and policy, source/replay/candidate hashes |
| ZIP measurement | 1 | 1 | Compressed framing and archive layout |
| ZIP emission | 1 | 1 | Publication, candidate artifact hash, preservation and output-progress checks |

Deterministic authored input opens five cursors. Store routes emit one producer
pass and then open four replay readers. Source artifact hashing, freshness
fences, cancellation, durable patch/inverse guards, and independent readback
are additional obligations; this table is not a total raw-I/O inventory.

Relevant production functions are `tail_append_stream::prepare_stream`,
`tail_append::scan_main_part` / `scan_reader`,
`prepare_source_part_splice_with_replay_handle`, `verify_source_proof`,
`verify_candidate_proof_replay`, `audit_splice`, and ZIP replay
`measure_callback` / `emit_callback`.

## Rejected direct removal

The candidate audit hashes original source bytes, but its XML parser sees the
concatenated source prefix, replay bytes, and source suffix. A matching source
hash proves byte identity, not independent XML validity. The public generic
OPC replay API permits a caller to supply:

```text
source:    <r>
insertion: offset 3
replay:    </r>
candidate: <r></r>
```

The original source is malformed while the candidate can pass.
`SourcePartSpliceProof` contains scalar lengths/hashes, and OPC does not impose
the DOCX producer's complete-paragraph grammar on generic replay bytes. The
DOCX scanner's earlier refusal does not justify weakening generic OPC. Removing
the source audit would also change source-before-replay error precedence and
discard independent source XML resource-policy checks. Its full proof must be
preserved by any replacement.

A possible fused verifier would consume source transport/decompression once
while maintaining independent source and candidate XML audit states, separate
hashes, bounded combined workspace, cancellation/freshness fences, and source
error precedence. It would still parse both logical streams. The remaining
source transport cost must justify the extra state and error-order complexity.
The prior 0483 profile's overlapping audit/replay ancestry does not establish
that removing one decompression pass would materially improve the operation.
Current production remains unchanged pending formal measurements and profiles.
