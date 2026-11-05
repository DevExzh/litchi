# 0362ul60 — DOC/OLE and RTF worktree retirement

The coordinator retires two clean, merged leaf worktrees while integration is
held at `d14f3613e508cde514c197f2d311871aa3174a33`. DOC/OLE
`d6cc88337c523c0d40970cdaa24c5fcc9ad7e347` and RTF
`1d273acbf21daba7e88164a11b749ed3355c63a4` remain reachable on their
retained branches and from integration. Both removals are non-forced.

All six ignored payload files (the two locks and four Python cache files)
are moved into the external quarantines with bytes, hashes, modes and inodes
preserved. Each sparse restoration passes both dry-run and actual restore;
the restoration probes are then removed non-forcibly. Root independently
verifies payload identities, retained branches, absent worktree registrations,
seven protected roots and the untouched 15-byte constructor probe. External
validation targets, logs and source captures remain intact.

RTF initially had a stale Git audit holder, PID 674230 and child 674307. Its
owner confirmed the process group was obsolete, terminated only that group
and verified its disappearance. The cleanup worker sent no signals; its
refreshed holder checks were empty before either removal.

Allocation measurements include the complete final quarantine, not only its
increase after preflight. The worker reports incremental reductions that are
24,576 bytes higher per leaf because preflight metadata already existed. No
filesystem free-space delta is claimed during concurrent builds.

| Leaf | Checkout bytes before | Complete quarantine bytes | Net reduction |
| --- | ---: | ---: | ---: |
| wave58-doc-ole | 7,559,970,816 | 266,240 | 7,559,704,576 |
| wave58-rtf | 32,747,520 | 331,776 | 32,415,744 |

Total net checkout-allocation reduction: **7,592,120,320 bytes**.

Evidence remains under
`(deleted-coord-artifact)`, in the two directories named
below. The manifest retains its original preflight status; the final receipt
records the refreshed checks and successful removal.

| Directory | Artifact | Bytes | SHA-256 |
| --- | --- | ---: | --- |
| wave58-doc-ole | manifest.json | 1456 | `4190c764a12c92387ac581728d7ef4927072164c6305af643c2eabb188187311` |
| wave58-doc-ole | restore.py | 2700 | `156b63adce21266d80c568f5a7bef3580baeddb240999051f60cb5b5815379f1` |
| wave58-doc-ole | restore-check.json | 570 | `788f5855b54e83eaf2aa89082bcf8baa0aa656eef952496e200f42313abdb80b` |
| wave58-doc-ole | receipt.json | 3510 | `14bfab575859fa7919a0d53fdfe501a660beef6b58002677d2e8041be3e613e3` |
| wave58-rtf | manifest.json | 1452 | `ed7ad08e1d044a309030a3f762f24d1ba3fbcf66845055a88248cdb2e4502cb6` |
| wave58-rtf | restore.py | 2700 | `156b63adce21266d80c568f5a7bef3580baeddb240999051f60cb5b5815379f1` |
| wave58-rtf | restore-check.json | 566 | `6f556f4ae449cd213fc912eb6d0db8876c59d563f97cc85ee18e1e9de9f67acb` |
| wave58-rtf | receipt.json | 3580 | `0881ec46e11a4eab8577b486489fa67b0c39dfffbf1a0190dd4a962d6c2d4491` |

To reproduce restoration, create a sparse worktree at the retained leaf tip,
then run the corresponding `restore.py <sparse-worktree> --dry-run` and
`restore.py <sparse-worktree>`. The external payload and manifest must remain
together. This cleanup changes no source or validation result; the integration
hold is released after the verified receipts.
