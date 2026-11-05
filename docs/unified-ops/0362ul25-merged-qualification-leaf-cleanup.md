# 0362ul25 — merged qualification leaf cleanup

Date: 2026-10-24. Scoped cleanup wave 50 removes three merged worktrees
without force. Their branches remain at the recorded tips:

| Worktree | Retained tip | Ignored payload |
| --- | --- | ---: |
| `uo2-legacy-manifest-identity` | `92ef3a2845578425655c79c5bd601f0f958af063` | 5 files, 362,177 bytes |
| `uo2-docx-ink-final-preparation` | `375ae218bf5a30d795c31c717ad5ed240a1aebb4` | none |
| `uo2-final-corpus-report-c796` | `ae2b7a425d3b37b57f0170a4827b38b62c15c3d7` | none |

Immediate ancestry, branch, clean-state and process-holder checks passed
against integration `023d0868e68c2e05623b39286869b44892bd0b19`. The
integration HEAD stayed fixed throughout removal and the restore exercise.
The five ignored Python cache files remain quarantined outside Git.

Sparse reconstruction and restoration verified payload hashes and modes,
parent-directory modes, exact ignored status and a second non-forced
removal. All three worktree paths are absent, their branch tips remain, and
the final holder audit found no process using the removed worktrees.

Evidence is retained under
`(deleted-coord-artifact)`:

| Artifact | SHA-256 |
| --- | --- |
| `manifest.json` | `cf39830d2417bb20121eca93617c11598b7e7c90f8db799e9fb24c3185dec38d` |
| `receipt.json` | `916e66c37674bed4902bc2e1787f667db55e1d8eac4c9fb199574de84f5b8fcb` |
| `restore-check.json` | `0638d5e0e0c204e43e0dcceb82d9dec2a604832ddaae147c1d5df54436862dea` |

The coordinator independently verified those receipts, all five payloads,
retained branch ancestry, worktree absence and protected directory identities.
The independent receipt is the sibling
`coordinator-wave50-verification.json`, 766 bytes, SHA-256
`845d04817a1828fc87fd74060c83d6bc6059efe7fe51d2296d373a9790b342b1`.

The final 472b source, legacy preparation packet, durable Ink build packet,
and corpus validation target and temporary directory remain intact. Actual
legacy driver manifests and the Ink build receipt have no dependency on the
removed author worktrees; future commands use retained sources. The protected
15-byte trunk constructor probe is unchanged. No foreign process was
signalled, no qualification result changed, and no target-size or free-space
reduction is claimed.
