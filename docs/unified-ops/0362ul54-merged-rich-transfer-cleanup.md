# 0362ul54 — merged XLSX rich-transfer worktree cleanup

Date: 2026-10-25. Wave 56 removes the clean, inactive
`uo2-xlsx-rich-transfer` worktree without force. Its retained branch
`uo2/xlsx-rich-transfer` points to
`044c37c51936ec4ae1a08efc0e2192b605b72368`, an ancestor of the held integration
`a10687e3b29583d117381710ecbc4e10685336b0`.

The only ignored payload is `Cargo.lock`: 139,592 bytes, mode `0664`, inode
309676981, SHA-256
`1d44ce42d8a66dd3967e80524d9cd830e15a2697646c43d84492628692361dc0`.
It is moved to quarantine on the same filesystem. Both a dry run and an
actual sparse hardlink restoration pass; the temporary restore worktree is
then removed without force. No target is present in the removed checkout.

An initial guard finds an unrelated recursive search holding a file in this
checkout. Cleanup leaves that process untouched. After the exact process
disappears, fresh cwd/FD/argv guards pass. The 32 bounded coordinator-state
checks find only historical cleanup metadata, with no active executable
binding. No foreign process is signaled.

The checkout occupied 7,549,100,032 allocated bytes. The final complete
quarantine occupies 184,320 bytes, giving a checkout-minus-quarantine
difference of 7,548,915,712 bytes (approximately 7.55 GB). The worker's
receipt also records an earlier allocation snapshot and its incremental
quarantine delta; those are distinct accounting points. No isolated
filesystem free-space delta is claimed.

## Evidence

External quarantine root:
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave56-xlsx-rich-transfer/`.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `manifest.json` | 978 | `e4f63edaf8aef63b227036e35fbf317a4eefe7eeeca90c7e58cdf9d7b4d2e76f` |
| `restore.py` | 2,079 | `c2b4d5917801086caa6ad62cfc556968f5dacb82c9d287872eb720189d624c25` |
| `restore-check.json` | 705 | `3bece2e31ad5e26f4425dcf8bd3b5ae3bac1f97f1ae11f1a1cc64350b13471d1` |
| `receipt.json` | 4,227 | `08bc9d549c7e3dfbc52cc0ae3d97f47ddcf30c59599a11b565495a298f4aae03` |

Root verifies all four artifact digests and the lockfile's bytes, mode and
inode. Its own checks confirm branch retention, absent checkout and worktree
registration, unchanged held integration, seven protected source/evidence
roots and the original untracked 15-byte constructor probe. The lockfile
itself occupies 143,360 allocated bytes; the receipt's 147,456-byte payload
figure includes its containing directory.

The integration hold is released. No source, active target or retained
qualification evidence is removed, and this cleanup makes no API or
performance completion claim.
