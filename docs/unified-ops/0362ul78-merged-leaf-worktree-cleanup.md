# 0362ul78 — merged font, crypto and Markdown worktree cleanup

Cleanup wave 60 removes three fully merged, clean leaf worktrees with ordinary
non-forced `git worktree remove`. The coordinator held integration at
`0dc6c76634555aa610d31f43ca96b739f93b8720` and independently verified that
every leaf tip is an ancestor. Branches and tested-source refs are retained:

| Removed worktree | Retained branch | Tip |
| --- | --- | --- |
| `uo2-shared-font-raw` | `uo2/shared-font-raw` | `07680db96980e49e7facb4d7078e7582b3b24eb6` |
| `uo2-crypto-raw-codec` | `uo2/crypto-raw-codec` | `7d765ad6e069f770842eb908347d4670a92a0fcd` |
| `uo2-markdown-policy` | `uo2/markdown-policy-0362ul70-atomic` | `e8b359eb6ac4561da5e48c54c0dd3c707fbfcd37` |

The owners verified and retired six stale read-only audit process groups
before removal. Fresh cleanup checks found no remaining holders; no foreign
process was signalled. Seven ignored payload files, including the exact
workspace locks, are quarantined with sizes, hashes, modes and inodes
preserved. They total 566,920 logical bytes and 577,536 allocated file bytes.
Dry-run and actual same-device sparse restoration pass. A preliminary
cross-device restore rejects with `EXDEV` before payload mutation; the
same-device retry verifies all files and parent-directory modes.

The external receipt directory is
`.coord/worktree-cleanup-2026-10-23/quarantine/wave60/`:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `manifest.json` | 4,458 | `4a05eaa3ff1127a2c158629e8d7331a301dbd8ccdd77863c1864ecd3cddccc89` |
| `restore.py` | 3,938 | `284476514002eba5ffa9c1a62030ace24242c61b7b9a99b7f25cdb22f635dd3a` |
| `restore-check.json` | 1,009 | `99b26bbc342d8ce9b7525cc818395e9b19238c5711dd28e2c76b15786800ed80` |
| `receipt.json` | 4,909 | `e3be6c5952fdc947638d6753d4aaef99157eea0f313ac8db801308fa3aa9cb5c` |

The removed checkouts occupied 7,874,048,000 allocated bytes. The complete
final quarantine, including receipts and directories, occupies 638,976
bytes. The conservative net reduction is therefore **7,873,409,024 bytes**.
These are `du` allocation measurements, not a claim about concurrent filesystem
free-space changes.

Root independently rechecks every preserved payload hash, size, mode and
inode, all three retained branch refs, absent worktree registrations and the
allocation arithmetic. External validation captures and targets remain
untouched. The protected constructor probe remains 15 bytes, mode `0664`,
inode `115179642`, SHA-256
`be179fcf7bbc53ebd35461556b67f6c1c77272b540d511bb365a039d5a387cc0`.
The cleanup hold is released; the separate tooling-row source hold continues
until that command completes.
