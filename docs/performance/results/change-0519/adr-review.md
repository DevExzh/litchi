# 0519 ADR review: reuse an immutable OPC XML proof

All 30 previously read accepted ADR/index files were hash-revalidated against
0518 before implementation. Their current hashes and base revision are in
`adr-manifest.json`; no ADR was changed.

| Contract | Preservation in this batch |
| --- | --- |
| ADR 0001/0002/0010/0011/0024: priorities and ownership | The conditional remains private in OPC. No facade API, dependency edge, or archive type changes. |
| ADR 0003: immutable snapshots and source-checked publication | Both original and assembled SourceXmlPart payloads are immutable, constructor-validated proofs. Replacement identity and original-byte checks remain. |
| ADR 0005: source, memory, Work and cancellation | Equal full ReadLimits selects proof reuse. Retain one payload-length Work charge, destination PartBytes, and source/context fences. Only unused parser working-memory admission and parser work disappear. |
| ADR 0006: preservation/security/validation | Different limits retain full destination validation. Content-type mismatch remains a refusal. Package signature/encryption/topology/output checks and DOCX candidate semantic reparse/readback remain unchanged. |
| ADR 0008: verification | Fresh paired native, method-profile and separate allocator evidence plus applicable format, workspace, lint, documentation, boundary, claims and independent ZIP64 checks gate retention. |
| Other accepted records | No changes to their format ownership, object models, typed metadata, CRUD, template, calculation, or migration contracts. ODF is deferred and iWork excluded. |

The full ReadLimits equality is deliberately stronger than comparing only XML
fields. The proof establishes XML validity under that exact policy; it does not
authorize a different destination source or bypass an output/security boundary.
There is no new unsafe code, mutable proof access, global cache, executor, or
ambient I/O.
