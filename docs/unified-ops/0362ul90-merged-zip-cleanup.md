# 0362ul90 — Accepted ZIP worktree cleanup

Date: 2026-10-25. Cleanup wave 63 removes only the accepted ZIP author and
atomic worktrees. Their branches, tested source, atomic proposal, external
validation evidence and build target remain available.

| Removed worktree | Retained branch tip |
| --- | --- |
| `uo2-zip-alias-closure` | `f80dcebc24344691e590fdf270979e967a4a5849` |
| `uo2-zip-alias-closure-atomic` | `04879dd5036d5d79b351671f4d435e3db7af8674` |

Both owners released the worktrees; tracked trees were clean and no process
held their paths. Two ignored payloads, the approved root lockfile and one
Python bytecode file, move into same-device quarantine. They total 159,583
logical bytes and 159,744 allocated file bytes. Dry-run and actual sparse
restoration preserve hashes, modes, inodes and directory modes; the temporary
restoration is removed. Both worktrees are removed without force.

Original checkout allocation is 15,131,316,224 bytes. Final quarantine,
including metadata and receipts, occupies 208,896 bytes. Net reclaimed
allocation is **15,131,107,328 bytes**; net apparent size is 12,551,828,253
bytes. These are filesystem measurements, not memory/performance results.

Evidence remains outside Git at
`(deleted-coord-artifact)`:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `manifest.json` | 2,334 | `a0edfca489cf6e66f6351e7592ce47d72883e6e1229d284a89abec9fa2ce9cd5` |
| `restore.py` | 4,536 | `611e42369c656f66abeeb98dcb719c03c4c7d2cd8a8e551031e329afe6df3a96` |
| `restore-check.json` | 836 | `d47077f15f49a38dd0bce17d9fa70219c395fb3a926e810d4f775947e8400f7b` |
| `receipt.json` | 2,710 | `591a5615edf6b105663e0e373bebdeaa549c62ea9f27d3685d1b25645701ccf7` |

The coordinator independently verifies retained branches, absent paths and
worktree registrations, payload identities, allocation arithmetic and the
protected 15-byte constructor probe. Trunk `29eaf2e924` and integration
`76b2d275f9` remain unchanged during the hold. The holds are released. No
source, external evidence, target or native input is removed or modified.
