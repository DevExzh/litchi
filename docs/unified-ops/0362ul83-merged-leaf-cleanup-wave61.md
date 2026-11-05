# 0362ul83 — Merged ODF, VBA and RTF leaf cleanup

Date: 2026-10-25. Status: complete; coordinator independently verified the
terminal receipt. This is workspace maintenance, not format or performance
qualification.

The three clean worktrees were removed with non-forced `git worktree remove`:

| Worktree | Retained source branch tip | Accepted atomic commit |
| --- | --- | --- |
| `uo2-0362ul44-odf-rawcodec` | `8a6dc1978e3cbeea532f14e0bad53d4ef8c7695b` | `677093e42f6cbdcf2770fb9633f2c972bed79572` |
| `uo2-vba-raw-policy` | `4a22c3c136ac7cf7766bc340f13eaf2ce5657a44` | same |
| `uo2-rtf-remaining-raw` | `d80cbb4fb6a1164e3bef614a22d03bfafb6938eb` | same |

All accepted commits are ancestors of held integration
`11d14a661d8ca2df4763408bc240f2772cc5ae36`. ODF's source branch was retained
separately from its rebased atomic proposal: its 98 implementation paths are
identical to the accepted atomic tree relative to common base
`30aaafbf98e097aeec8cf6fe3df6caaae62a37d6`. Source, atomic, tested and receipt
refs remain available. The worktree paths and registrations are absent.

Owners verified and retired their stale read-only audit processes. A foreign
report handle was released before removal. Cleanup sent no foreign signals.
All 13 ignored files were moved into a same-device quarantine, preserving
their hashes, modes and inodes. The inventory contains 703,700 logical bytes
and 729,088 file-allocated bytes. A dry run and actual sparse restore passed;
the temporary restore destination was removed afterward. External validation
captures and targets were retained.

The external receipt directory is
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave61/`:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `manifest.json` | 6,966 | `dec602ef95ecbadc9fd269802602fc38e3d305de12b66f9e0628d66d8c619bcd` |
| `restore.py` | 3,881 | `945523df2083523b5396562729ff25be63299c6530615c14262d7f7b89438a2e` |
| `restore-check.json` | 819 | `d0ad8ded4743f681e97a4e9bce989a06f31ee685d072d97d7acb28ee7d72c5b9` |
| `receipt.json` | 5,509 | `eff34a7fe92ced60d62d12d5f3a8df6c01f5b02b053dc5b608f6750b66a8dc66` |

The removed worktrees occupied 15,194,042,368 allocated bytes. Subtracting
the entire final quarantine allocation of 798,720 bytes gives a conservative
net reclamation of **15,193,243,648 bytes**. The receipt also reports the
incremental result, which excludes the quarantine's pre-existing setup
files. These are `du` accounting figures, not a claim about concurrent
filesystem free-space movement.

Root rechecked all payload identities, retained refs, accepted ancestry,
absent paths and registrations, held source identities, and allocation
arithmetic. Trunk stayed at `171074aac310361d589233dfb5b057ed30b5777a` during
cleanup. The protected 15-byte constructor probe retained its exact bytes,
mode and inode. The integration and trunk holds are released.
