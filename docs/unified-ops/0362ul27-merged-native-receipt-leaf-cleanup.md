# 0362ul27 — merged native receipt leaf cleanup

Date: 2026-10-24. Cleanup wave 51 removes the two eligible merged author
worktrees without force:

| Worktree | Retained branch | Tip |
| --- | --- | --- |
| `uo2-doc-body-toggle-472b` | `uo2/doc-body-toggle-final-production-rebind` | `eb10f976c448461a79ad11c402b598221f9eefde` |
| `uo2-iwork-native-allocator` | `uo2/iwork-native-allocator` | `2cf3200d9cb8a5ee86dda3ae8f6995f145b913d9` |

Both tips are ancestors of the held integration revision
`c204999a444b799658ed56b5e92609986b21fbd4`. Immediate clean-state, branch,
ancestry and process-holder checks passed. Five ignored files, totaling
389,976 bytes, were quarantined outside Git. A sparse reconstruction and
restore exercise verified every payload hash and mode, parent-directory
modes, exact ignored status and a second non-forced removal. Branches remain
at their recorded tips.

The glossary reconciliation worktree was retained because a text-search
process held an open file there. Cleanup sent no signal to that process.
The search owner subsequently confirmed it was an obsolete DOC overlay-hash
lookup and stopped its own process group. The retained worktree is eligible
for a later audit; this record does not claim its removal.

Evidence is under
`(deleted-coord-artifact)`:

| Relative artifact | SHA-256 |
| --- | --- |
| `quarantine/wave51-partial/manifest.json` | `4d5bc5c361d5e97d9d72accdc75d24f773eba0a536e91ee29973677d462c22c8` |
| `quarantine/wave51-partial/receipt.json` | `2a6cdcf5b18a122de06b296cd35b6821029e48f9ef71363608980a314ac6934d` |
| `quarantine/wave51-partial/restore-check.json` | `3cd540896b0f72fe2a3460fd00a5f956aa6b382abee9eb956e1a44f78d66d3d7` |
| `quarantine/wave51-scoped/blocked-glossary.json` | `d1dbe2f0591dd7432f7ad37ee075987182fa65c42cfe0c3d23ba5e89c722e4d4` |
| `coordinator-wave51-verification.json` | `8b1ac087f48a8a6de92b8f2bd170c2808373dde5a1aa6e596a6a1e76dc323b63` |

The coordinator independently verified all receipts, five quarantined
payloads, retained branch ancestry, worktree absence, the skipped worktree's
presence and protected external directory identities. The final DOC and
iWork packets, source snapshots, validation targets and historical evidence
remain intact. The protected 15-byte trunk constructor probe is unchanged.
No performance result changed, and no target-size or free-space reduction is
claimed.
