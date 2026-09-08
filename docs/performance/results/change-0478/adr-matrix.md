# 0478 ADR compliance and scope

All 30 previously read ADR/README hashes match `adr-refresh.json`. This batch
implements fresh PPTX plain-text streaming creation; it does not redefine
logical append, adding arbitrary package Parts, or edit-and-repackage.

| Constraint | Implementation and verification |
| --- | --- |
| 0001 priorities/API layers | Checked capability and typed prepublication refusal; ordinary semantic PPTX APIs expose a scratch policy, not ZIP types. Cross-route and public integration tests cover failure progress. |
| 0002/0024 crate topology | ZIP owns symbolic path proof/cursor and storage; OPC owns absolute PackURI validity and Part uniqueness; PPTX owns fixed PresentationML member topology. Boundary checks remain required. |
| 0003 snapshots/edits | This new route creates fresh output through the existing streaming writer. It does not mutate snapshots or change Commit/Patch semantics. |
| 0005 I/O/memory/evidence | Explicit Read+Write+Seek provider, finite serialized-byte quota and replay window, no ambient file opening in production. Complete operation allocation/timing scopes include construction and destruction. Provider storage and process RSS remain separate metrics. |
| 0006 preservation/security | Same serializer and exact byte/semantic oracles; complete sequence checked before output. Strict raw-name sequence, ASCII-equivalence and whole-component ancestor proof replace growing name indexes only in the explicit generated mode. Ordinary arbitrary-name checks stay active. |
| 0008 verification | Required test/lint/doc/boundary gates retain source-bound command receipts and development failures; measurements and portable verification remain independently reproducible. |
| 0010 archive ownership | Only the ZIP owner implements central-directory records, name normalization and checked cursor consumption. No format dependency enters ZIP. |
| 0011 OPC ownership | OPC wraps ZIP plans privately, validates production PackURI endpoints, and enforces exact checked sequencing on every physical writer route. No ZIP types enter OPC public signatures. |

The plan's representation is bounded by descriptor count and static bytes.
Indexed expansions are represented by checked endpoints; validation compares
finite descriptors and never enumerates the declared range. Runtime retains
one active member/name plus bounded compressor/replay state. A scratch provider
may retain the serialized directory externally (or in its own memory). The
library cannot guarantee constant whole-process RSS or remove that storage.

The original default constructors still retain general-purpose name indexes
and central-directory metadata. The new route requires a known complete member
sequence; it is not a bypass for arbitrary user-provided names, an unchecked
validation token, or a new serializer hierarchy.
