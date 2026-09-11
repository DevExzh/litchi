# 0511 ADR review

All 29 accepted records and their README were previously read and are freshly
hash-verified unchanged against [adr-manifest.json](adr-manifest.json), bound
to base `b0b9ebb577d616a9a0ed93f16619a1b79f3b3bc7`. No exception is required.

The production diff is confined to the private CFB FAT entry loop. The
[sector capacity review](sector-batch-review.md) proves each extension fits
the existing exact fallible reservation. It preserves reachable decode errors,
source reads, FAT/DIFAT markers, limits, ownership validation and publication.

| Records | Disposition |
| --- | --- |
| 0001 | Correctness and preservation precede measured performance; no public-layer change. |
| 0002, 0024 | CFB remains the canonical container owner; package graph and dependencies unchanged. |
| 0003, 0004, 0007 | No snapshot, edit, patch, concurrency, typed API or semantic-model change. |
| 0005 | Existing fallible capacity budget and positional reads remain; measured timings, allocation deltas, RSS and instruction scopes stay distinct. FileSource and broader scaling work remain open. |
| 0006 | Exact reservation failure, complete-chunk decoding, structural checks, source fences and malformed-input order preserved. No unsafe code introduced. |
| 0008 | Frozen source/binary/corpus identities, raw paired samples, malformed-input and consumer tests, strict checks and replay bind retention. No broad compatibility claim. |
| 0009, 0023 | ODF ownership unchanged; further ODF optimization deferred by user priority. |
| 0010, 0011 | Facade, archive and OPC ownership unchanged. |
| 0012, 0016, 0027 | BIFF8 reference/writer location types and XLS sheet anchors unchanged. |
| 0013–0015, 0017–0022 | PPTX notes, core properties, producer templates, XLSX calculation chains, DOCX settings/glossary, PPTX styles/fonts unchanged. |
| 0025 | OGraph chart transactions unchanged. |
| 0026 | OLE directory metadata binding unchanged; FAT values and directory identities are preserved. |
| 0028, 0029 | iWork remains outside this workstream. |

The source proof does not authorize changing generic fallible push callers,
MiniFAT, chain traversal or sector ownership. The first push-only variant is
rejected diagnostic evidence. Final measurements, test counts and retention
are documented in [0511](../../changes/0511-cfb-fat-entry-reservation.md).
