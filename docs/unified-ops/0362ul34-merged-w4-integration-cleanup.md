# 0362ul34 — merged W4 integration worktree cleanup

Date: 2026-10-24. Cleanup wave 54 removes the merged W4 integration
checkout without force and retains its branch and external evidence.

| Worktree | Retained branch | Tip |
| --- | --- | --- |
| `uo2-coordinator-w4-integration` | `uo2/coordinator-w4-integration` | `f3fed35c0de7e2d67bba2d6c7c31ff021f716f16` |

The tip is an ancestor of held integration revision
`d07d30b1b646c42dcd2ffb763d2c06b8d29dd01e`. Tracked/index cleanliness,
ancestry and process dependencies were rechecked before removal. No other
worktree was removed in this wave.

All thirteen ignored regular files were moved to external quarantine:
`Cargo.lock` and twelve `tools/__pycache__/*.pyc` files, totaling 6,960,154
bytes. Their paths, sizes, SHA-256 hashes, modes and original inodes were
preserved. The initial twelve-file preaudit count was corrected before any
payload move. A sparse detached restore exercise passed both dry-run and
actual restoration, including file and destination-directory modes. The
restored temporary checkout was then removed without force.

Evidence is under
`/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave54-coordinator-w4-integration/`:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `manifest.json` | 4,627 | `b0a9d2f553259b1895875f075b108fa736fcddf155e71cfb96650af64f4996fc` |
| `receipt.json` | 5,005 | `2a9439a27380a4fda3e7d19d0b7bd517a350fea18c842e1962115018d8d18cf5` |
| `restore-check.json` | 1,149 | `c3d0f1e85204bfdbae7fa68a4a1076f177a317f54d7385dac2310c6b1a035430` |
| `restore.py` | 2,823 | `4e64c79dcbcab800869b565608d1f13111e4ef136afcda3e0cb70ee6e4024cd5` |

To restore, create a sparse worktree at the retained tip, create the parent
directories listed in the manifest with their recorded modes, and run
`python3 restore.py DESTINATION --dry-run`, followed by the same command
without `--dry-run`. The helper checks existing files and rejects mismatched
collisions. Quarantine directory modes need not equal the original source
parent modes; the helper verifies the latter at the restore destination.

The coordinator independently checked all thirteen payload identities,
branch ancestry, absent checkout and registration, retained source roots,
and the unchanged fifteen-byte trunk constructor probe. The applicable
postchecks total 44 and pass. The initial audit's three invalid comparisons
of quarantine-directory modes to restore-directory modes are retained and
explained by the corrected audit; no payload or worker receipt was changed.
The successful restore receipt was reviewed rather than repeating the
restore exercise.

Coordinator audit:
`/home/zhuhe/.cache/litchi-api-contract/worktree-cleanup-wave54-audit-v2-20261024.json`,
5,886 bytes, SHA-256
`67e52ecf6e507a5800a42d9ac13f57c95ecd401c1d73fc6810a1f95813a8095a`.

The removed checkout had 7,550,095,360 allocated bytes; quarantine retains
7,020,544 allocated bytes. Their difference, 7,543,074,816 bytes, is a
checkout-allocation estimate, not an isolated filesystem free-space delta.
No pre-removal `df` snapshot exists. Branch content, failed-run artifacts,
measurement binaries, active documentation and the foreign orphan process
remain preserved.
