# 0436 ADR compliance

| Constraint | Evidence |
| --- | --- |
| 0001 priorities/API layers | Private encoder change; exact output and resource refusals take precedence over batching. |
| 0002/0024 ownership | Only litchi-odt production/tests change; no dependency edge or public API changes; boundary gate passes. |
| 0003 edit/publication invariants | Fresh creation only; existing edits, patches, conflicts and snapshots are untouched; failure retains accepted sink progress. |
| 0005 execution/performance | Explicit existing context and limits, per-scalar polling, at most 256-byte commit batch, hierarchical Work rollback, serialized measurements and retained regression flags. |
| 0006 preservation/security | Exact archive/content/styles/meta/semantic/sink identities; whitespace/escape paths and common XML audit remain intact; scalar threshold differential tests pass. |
| 0008 verification state | Source/binary/driver-bound receipts, explicit failed compile attempt, final 1,002 ODT tests and portable mutation proof; no unsupported native/scaling claim. |
| 0010/0011 physical ownership | ODT grammar stays in format owner, fragment/XML/ZIP publication stays in unchanged common owner. |

No unsafe code, ambient runtime behavior, new allocation strategy or parallelism
is introduced. Accepted ADR tree is unchanged at
`c950b6c8be822561b498d7bbe87c460873dcbf49`; the user goal remains pinned and unmodified.
