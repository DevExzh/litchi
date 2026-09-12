# Publication proof reuse: ADR constraints

The 30 previously read accepted ADR/index files are bound by
`adr-manifest.json`. Their SHA-256 values were checked again against the
worktree before implementation and are unchanged from the preceding batch.
The source baseline is `afb62ab7a70859dad4a4b9c8eea91402d1ef4052`.

| Constraint | Required implementation and verification |
| --- | --- |
| 0001, 0004: API layers | The optional hint is an opaque OPC `SourceXmlPart`; ordinary DOCX CRUD APIs gain no archive types, locks, physical IDs, or provider generics. |
| 0002, 0010, 0011, 0024: ownership | OPC owns current Part capture, source XML validation, and proof reuse. DOCX owns document grammar, indexes, and exact patch source matching. Dependency edges remain unchanged. |
| 0003: immutable transactions | Reuse only the immutable original snapshot retained by the patch after current-source authorization. Preserve `Patch::apply`, conflicts, inverses, candidate reconstruction, reparse, and semantic readback. |
| 0005: sources and resources | Keep the current Part read, cancellation, version fences, equal read limits, and original reservation ownership. A hint cannot create authority from a foreign provider. Full validation uses the same fresh Part data on a mismatch. No global cache, ambient I/O, parallel runtime, or unsafe optimization. |
| 0006: preservation and security | Keep XML classification, encrypted/signed package refusals, source byte equality, changed-document policy, destination topology/replacement checks, bounded output, and transfer monitoring. Exact no-ops retain their existing branch. |
| 0008: evidence | Candidate retention requires matched native, instruction, and allocation evidence plus source/error/resource/preservation tests. Raw measurements, adverse flags, and scope limitations remain reviewable. |
| Remaining accepted ADRs | No format ownership migration, semantic feature change, normalization, repair, or cross-format dependency change is proposed. ODF work is deferred and iWork is excluded. |

This matrix records proof obligations, not a claim that an unmeasured candidate
is accepted. The final source review, tests, and matched results determine
retention. In particular, the required candidate reparse/readback cannot be
replaced by an original-source hint.
