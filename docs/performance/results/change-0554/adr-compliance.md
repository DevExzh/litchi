# 0554 ADR compliance

The 30 accepted ADR/index hashes are unchanged from the previously read 0553 manifest and are bound in `adr-manifest.json`. This experiment retains no production runtime change: the exact baseline is restored after native gate failure.

| Constraint | Evidence and disposition |
| --- | --- |
| ADR 0001, 0002, 0024: API and ownership | Private CFB parser only; no public API, dependency or runtime ownership change. Final workspace and crate-boundary checks remain required. |
| ADR 0003: immutable state and atomic publication | Normal names transfer between staged private/public owners. Graph and cache assignment still occurs after complete directory processing. Targeted failed-load publication test passed. |
| ADR 0005: bounded resources and measured performance | Directory/resource ceilings, checked extents and remaining fallible reservations stay in place. Duplicate normal-name reserve intentionally disappears; resource-boundary review states the exact limitation. Native, allocator and instruction lanes are distinct. Candidate rejected after XLS native regressions. |
| ADR 0006: validation and compatibility | Name, UTF-16, graph, type and scalar validation remain authoritative. Classic-Mac root keeps its two distinct name views. Targeted scalar/cache, exact malformed-name and root tests passed; final legacy-format tests also passed. |
| ADR 0026: shared OLE directory metadata | SID, object kinds, CLSID, start sector, masked size and MiniFAT mapping remain owned by CFB and are checked by the focused raw/public comparison. No fabricated timestamps or new shared semantic fields. |

The review does not replace the source binding, full quality receipts or frozen performance gates. No accepted ADR was edited for this batch. OLE2 and OOXML remain the optimization priority; ODF is deferred and iWork excluded.
