# 0362ul73 — Four merged leaf worktrees retired

Cleanup wave 59 retires the merged XLS, SpreadsheetDrawing, DOCX and PPT
worktrees while integration is held at
`b8897649d6364c5d4b90300bfb97f36748975fc9`. All four branch references stay
at their exact accepted tips, and every tip is an ancestor of integration.
The worker uses non-forced worktree removal and preserves all 18 ignored
files with bytes, SHA-256, modes and inodes unchanged. Both dry-run and actual
sparse restoration pass for every file; temporary restoration worktrees are
also removed non-forcibly.

Root independently verifies the payload identities, manifest/receipt/restore
hashes, retained branch references, ancestry, absent paths and worktree
registrations, and the protected constructor probe. Its 15 bytes, mode 0664,
inode 115179642 and SHA-256
`be179fcf7bbc53ebd35461556b67f6c1c77272b540d511bb365a039d5a387cc0`
remain unchanged. External build targets, validation captures and evidence
are retained. No production source is changed by cleanup.

DOCX initially had an obsolete audit shell (PID/PGID 3940629) and its
`git cat-file` child 3940777. Their owner verified the command, working
directory and process group, terminated only that group and checked both
PIDs absent. Cleanup proceeded only after refreshed holder checks were empty.

| Leaf | Retained branch | Commit | Checkout allocation before |
| --- | --- | --- | ---: |
| xls | `uo2/xls-raw-limits-0362ul46-sparse` | `2ffacb16e929cdad7b032f5e46f25ed72fbc4ef6` | 62,148,608 bytes |
| spreadsheet-drawing | `uo2/0362ul51` | `9f43932dd0a5a9f5aef4ce439915f1563f2e6940` | 7,559,663,616 bytes |
| docx | `uo2/raw-limits-docx-ooxml` | `f9e3bff3b30f2cda083b3b1eac2a0cf0c6c0a58e` | 410,365,952 bytes |
| ppt | `uo2/ppt-record-limits` | `6a1317ab04e57cf1fe74b1a2637a8e0be6c46003` | 434,524,160 bytes |

The four checkouts allocated **8,466,702,336 bytes** before removal. The
complete final quarantine, including receipts and summary, allocates
**1,175,552 bytes**. Net checkout-allocation reduction is therefore
**8,465,526,784 bytes**. The preserved payload is 999,817 logical bytes and
1,036,288 allocated bytes. These are `du` allocation measurements; no
filesystem free-space delta is claimed during concurrent builds.

Evidence remains under
`.coord/worktree-cleanup-2026-10-23/quarantine/wave59/`. To reproduce a restore,
create a sparse worktree at the retained leaf tip and run that leaf's
`restore.py <worktree> --dry-run`, followed by `restore.py <worktree>`.
Keep its manifest and ignored payload together. All four restore helpers
have SHA-256 `156b63adce21266d80c568f5a7bef3580baeddb240999051f60cb5b5815379f1`.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| summary.json | 3300 | `aa257545bd1bbe2c676f1f29f6d6565480e9dd404bf2410536494b98c564c20d` |
| xls/manifest.json | 2154 | `6271e6f578a6fc1d8802a322cd9b856888ebb9ccc937b2eb80dd035bb2141f6d` |
| xls/receipt.json | 2402 | `a7dbfcb72723880522a9f6363ddf85e6c19984ee9b58d5d4a71947865ae14830` |
| xls/restore-check.json | 417 | `d2dc572e1e79d6855cbf28514d6ed1ea253f9e6d81650bfb6129a471d96c3b42` |
| spreadsheet-drawing/manifest.json | 1915 | `16ee991199804f34a7f4d80023afb2a1f80553ac5270798cfcf66732486a4173` |
| spreadsheet-drawing/receipt.json | 2467 | `a3dc7f7842979215512eec8dde76d9285ed1f6bcce3387358959331570810c65` |
| spreadsheet-drawing/restore-check.json | 433 | `5714220ce4c3f80895af9e90e6e879ac80fb447bb1f5a85daef1dce7f29f15a0` |
| docx/manifest.json | 3229 | `28ba501d567da94bc4bd074913eb8cb07bec424d27e4a220c924922cf3096aa2` |
| docx/receipt.json | 2382 | `61bd6a118039df7bccca917389c69b97403ab06de2eb59e7afce3485f0dc9d60` |
| docx/restore-check.json | 418 | `935add5d066172feeb2c29b12498bfdde57d38f79dcf0ce05854534edb8a11ea` |
| ppt/manifest.json | 2945 | `50ff7ae35013090f07339811e9b163b6427458d5d6bcdb373e17fdff5694fcbb` |
| ppt/receipt.json | 2375 | `3aaec04f9231038e0b83dfea11003f8e0a8257be73a10b7085dc6de18b1833f5` |
| ppt/restore-check.json | 417 | `d2f916e08311157fcc7ce50ab5fa74bd4c22e4665f4c1d79abdb41f319c13918` |

The integration hold is released after root verification.
