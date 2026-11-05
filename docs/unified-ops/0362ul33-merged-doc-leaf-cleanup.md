# 0362ul33 — merged DOC author leaf cleanup

Date: 2026-10-24. Cleanup wave 53 removes one clean, merged author worktree
without force. It preserves the accepted DOC build and measurement sources.

| Worktree | Retained branch | Tip |
| --- | --- | --- |
| `uo2-doc-body-toggle-final` | `uo2/doc-body-toggle-final-qualification` | `a83add46ef65c2e71d1ae644e27e67ed744cba80` |

The tip is an ancestor of the held integration revision
`589d5d0adca1897779f0a056c69dfeee898a62cf`. Tracked/index cleanliness,
ancestry, process holders and current DOC command dependencies were checked
before removal. Conditional candidates were left in place.

The only ignored payload, `Cargo.lock`, is retained in external quarantine:
139,606 bytes, mode `664`, SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`.
Its inode is preserved by the move. A sparse restore exercise verified the
payload's hash and mode, then removed the temporary restored copy and
worktree without force. The original branch remains available.

Evidence is under
`(deleted-coord-artifact)`:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `manifest.json` | 1,454 | `444acd24f31af4d81053aa785419ff2cef2900eae7729cc7c64269afe65e5b11` |
| `receipt.json` | 3,255 | `001dd60d8d15f7b2472c1cdb1464e74df0253e675afdb5ca4d3bd03b67037409` |
| `restore-check.json` | 1,103 | `4806bdb3240ed903633fc7299357d1da2fa010e6a3fbcc5c5df79bc927a30cb0` |
| `restore.py` | 2,837 | `017ba78ca7787e4096472a73f3fbf1eba12df8371baa08e5efd62bdddb3b52c8` |

To restore, recreate a sparse worktree at the retained tip and run
`python3 restore.py DESTINATION --dry-run`, then the same command without
`--dry-run`. The helper rejects collisions. Existing evidence must remain
unchanged.

The coordinator independently passed 17 postchecks covering the exact payload
set, hash, size, mode and inode; branch and ancestry; absent worktree and
registration; protected source roots; the unchanged 15-byte trunk probe; and
the cleanup/restore receipt identities. This postcheck reviewed the restore
exercise receipt rather than repeating the exercise. Its external receipt is
`/home/zhuhe/.cache/litchi-api-contract/worktree-cleanup-wave53-audit-20261024.json`,
1,708 bytes, SHA-256
`672271502321342c9267bd4090ea9351a48e2fbcde665eb67f2e13787784e3e8`.

No target, measurement binary, failed-run evidence, production source or
foreign process was changed. No numerical free-space saving is claimed.
