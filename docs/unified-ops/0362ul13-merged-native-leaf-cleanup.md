# Merged native leaf cleanup

Cleanup wave 46 removes two merged worktrees non-forced while retaining their branch refs: DrawingML process-observer leaf `40091854020a07a233c792cac32eafc308d93120` and PPT native protocol leaf `079ad7e3bc277d16ce5cf165e612280fa15b85f0`. Both are ancestors of held staging `082fbeafe8f1e68c7118ec8561b5bd6fe1577415`.

Seventeen ignored payloads, 918,418 bytes, are retained externally as regular mode-664 files. The [restore exercise](results/0362ul13/restore-check.json) recreates the recorded branches, restores and checks every hash and mode, then removes both temporary worktrees non-forced. The [coordinator verification](results/0362ul13/coordinator-verification.json) independently checks all retained bytes, branch refs, ancestry and absent worktree paths. Their pre-removal allocated totals were 45,232,128 and 7,554,875,392 bytes; the retained quarantine is much smaller.

The original manifest and receipts remain under `/home/zhuhe/code/litchi-worktrees/.coord/worktree-cleanup-2026-10-23/quarantine/wave46/`. Their exact sizes and SHA-256 hashes are in the coordinator receipt. External targets, native endpoints, the earlier partial-checkout quarantine and the protected 15-byte trunk constructor probe remain intact. No foreign processes were signaled. OLE cleanup remains excluded because unrelated read-only holders still reference that leaf.

The short staging HEAD hold is released. This cleanup does not change qualification status or push any branch.
