# 0362ul88 — Merged Formula and XLDM worktree cleanup

Date: 2026-10-25. Cleanup wave 62 is complete. It removes only the two
accepted author worktrees; their branches, tested commits, atomic commits,
external validation captures and native inputs remain available.

| Worktree | Retained source tip | Integrated atomic commit |
| --- | --- | --- |
| `uo2-formula-raw-policy` | `2ae5623db931bef991d7f0d6ab7921c6db1efb15` | `20aa27d9137755243bfa02c67eee4b3d4cc14591` |
| `uo2-xldm-raw-policy` | `c2495b728f7c354a2f8480ecffa37690bff8efb1` | `c2495b728f7c354a2f8480ecffa37690bff8efb1` |

The owners retired their verified stale read-only scanner processes before
removal. Both tracked trees were clean. Five ignored files were moved into
same-device quarantine: two approved lockfiles and three Python bytecode
files, totaling 376,373 logical bytes and 385,024 allocated file bytes.
Dry-run and actual sparse restoration passed, preserving hashes, modes,
inodes and parent-directory modes. The temporary restoration was removed.

Both worktree paths and registrations are absent after non-forced removal.
An XLDM retry returned 128 after a cleanup/status race; the terminal checks
confirmed that its path and registration were absent. No force was used.

The original checkouts occupied 15,120,789,504 allocated bytes. Final
quarantine, including all metadata and the receipt, occupies 438,272 bytes.
Net reclaimed allocation is **15,120,351,232 bytes**; net apparent size is
12,542,290,091 bytes. These are filesystem allocation measurements, not
memory or performance results.

Evidence is outside Git under
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave62/`:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `manifest.json` | 3,276 | `610398ecd2e66e64bd01bbac526759fae7fe92984d85e55ef786627f71405759` |
| `restore.py` | 4,536 | `36eb459baaf0d9ab49d6dbfa6753618b51499ca9045d7848449f053197e86807` |
| `restore-check.json` | 824 | `e9d8a6d9be9f1f1fbf186be7e2d8c1b2ca37272c18b9abf05b40b98d9e2e0c72` |
| `receipt.json` | 2,376 | `334d400e0b0d90c9bc42f795ed7577eb383085dc160f06672c56c173b7b01682` |

The coordinator independently verified all five quarantined payloads,
retained branches, absent worktree registrations, exact allocation arithmetic,
external evidence paths and the protected 15-byte constructor probe. Trunk
`80d09722c5` and integration `66364d4fb0` remained unchanged during the
cleanup hold. The holds are released. No source change, acceptance claim or
push is part of this cleanup.
