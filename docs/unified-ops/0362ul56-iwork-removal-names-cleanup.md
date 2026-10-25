# 0362ul56 — retired iWork naming worktree

Cleanup wave 57 removes only `uo2-iwa-removal-result-names`, after its final
`0524a814aee7ae24a6f07a879ecd69dd0285e641` was merged at `21fafcc599`.
The branch `uo2/iwa-removal-result-names-0362ul48` remains available.
Integration stayed at `fa833c7d1bc7f2b5e97efb6531a20636427e3db0` throughout
quarantine, restore verification and non-forced worktree removal.

The tracked tree and index were clean. Both holder checks found no process
using the candidate; its only state reference was historical completed
validation. External validation logs and targets were untouched. No foreign
process was signaled. Root independently verifies branch retention, absence
of the removed checkout and its worktree registration, all payload hashes,
sizes, modes and inodes, the seven protected roots and the original 15-byte
constructor probe.

Eight ignored files (`Cargo.lock` and seven `tools/__pycache__` entries) total
404,565 logical bytes. The originals are quarantined with their identities
preserved. The collision-safe helper passed dry-run and actual restoration in
a sparse checkout; that temporary checkout was removed non-forcibly.

The checkout occupied 221,532,160 allocated bytes. The complete quarantine,
including its manifest, helper and final receipts, occupies 471,040 bytes.
Subtracting that complete quarantine gives a net allocation reduction of
**221,061,120 bytes**. The worker's incremental accounting subtracts only the
446,464-byte increase after its initial 24,576-byte preflight directory and
therefore reports 221,085,696 bytes. Neither number claims an isolated
filesystem free-space change.

All artifacts below remain outside Git under:

`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave57-iwa-removal-result-names`

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `manifest.json` | 2696 | `c89ef477563d56ac1a886c2f225f8103ac9d6ecfdd496f79bfdd282bbd0d9b18` |
| `restore.py` | 3119 | `acf0333f2e18450f86c79e6419875483d45ee176b557156d7d14a56ec625a3bb` |
| `restore-check.json` | 587 | `75849c9425052b4e212d54c4a4e1d44f14bfd6012b9b9f0695d25485907caa7e` |
| `receipt.json` | 3630 | `8d70b6ce04926f92582905e0dad65696579603bee243bdedb9cffc2afc0e4362` |

The manifest records every original payload identity. Restore into an empty
sparse checkout with `python3 <quarantine>/restore.py <destination>`; existing
files cause refusal. This cleanup changes no production source or acceptance
result.
