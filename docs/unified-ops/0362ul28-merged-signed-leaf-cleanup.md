# 0362ul28 — merged signed-API leaf cleanup

Date: 2026-10-24. Cleanup wave 52 removes five clean, merged worktrees without
force. All branch references remain at their recorded tips.

| Worktree | Retained branch | Tip |
| --- | --- | --- |
| `(deleted-coord-artifact)` | `uo2-glossary-source-reconciliation` | `bddab5e261ded19ded27bf2773a8a4ec413bf0dd` |
| `uo2-signed-pptx-budget` | `uo2/signed-pptx-budget` | `50cb32a80a10bfe871df2d1cc2bd183c7d628e08` |
| `uo2-signed-replay` | `uo2/signed-replay` | `06afe43fa08a4bddbdab0745b54042f6106dde8e` |
| `uo2-signed-xlsb-budget` | `uo2/signed-xlsb-budget` | `509bf28e7531ad75f278eaec3a732714a243915a` |
| `uo2-xls-f3-qualification` | `uo2/xls-signed-runtime` | `df2d948477de487357f3ac69e4b5dd3212c8a6fd` |

Every tip is an ancestor of the held integration revision
`c5a86b55529a003c13503d01e2fd00fddd2f56b6`. Immediate clean-state, branch,
ancestry and process-holder checks passed. The removed XLS worktree's actual
path is `uo2-xls-f3-qualification`; its branch has a different name.

The cleanup preserves 13,912 regular payload files totaling 9,680,195,101
bytes in external quarantine. These bytes were moved, not deleted. Restore
exercises passed. Independent verification subsequently rehashed every file
and checked size, mode and recorded identity, with zero missing, extra,
duplicate, symlink or metadata mismatches. It also verified all five retained
branches, absent worktree registrations and 25 protected external identities.
The current sheet readiness packet and its three execution commands contain
no dependency on a removed worktree. The protected 15-byte trunk probe is
unchanged. No numerical free-space saving or performance result is claimed.

Evidence remains under
`(deleted-coord-artifact)`:

| Relative artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `wave52-partial/manifest.json` | 3,395,229 | `55dd98390cac1d8ec179e61aae1025eaa0edc13df13cd0c5abe366804eb318f0` |
| `wave52-partial/receipt.json` | 18,001 | `1aa508da57898b483b3e1d4af34b39ab406f254915c4f0718340a2f7032e56d0` |
| `wave52-partial/restore-check.json` | 2,539 | `8672c1864e793a759f582b9ffca4294afc18bf2fc98e87221ce869ffdd62e2a8` |
| `wave52-xlsb/manifest.json` | 3,182,404 | `1103e921736ed8497909aebf46878327b4e645efbc39f41218c1fb2aa28cf99d` |
| `wave52-xlsb/receipt.json` | 15,845 | `0906e3f0f1da5ab3accff56a65a8f554c0a2ddbf6eeb82dc8641726358081183` |
| `wave52-xlsb/restore-check.json` | 1,777 | `a3cde1c4e591cbea878a6125e65b48b85cb2884cd7760469d7d5c2e646068c96` |

The large manifests stay outside Git. Their original generation commands,
relative to the quarantine root, were `python3 wave52-partial/prepare_cleanup.py`
and `python3 wave52-xlsb/prepare_cleanup.py`. Reproduction requires the recorded
integration revision and recreated sparse worktrees at the retained branch
tips, followed by each wave's `python3 restore.py` and a fresh evidence
destination. Existing manifests and receipts must remain untouched.

The receipt key `prepare_cleanup.py` identifies the imported engine, not the
small launching wrapper. The partial run used the 14,979-byte engine SHA-256
`bf46a11ca95856b1361ff8f6af7451a58e443251ec4c113f0c08f37ad6d0709f`;
the XLSB run used its 14,981-byte correction SHA-256
`a5b7dcf3a075130020464dd664d18ddddde2737164dc24dbaa6baf6c0818b108`.
Their only difference is the corrected full XLSB branch tip, which the
partial run excluded. The original engine was reconstructed byte-exactly
against its recorded hash and retained as
`wave52-scoped/prepare_cleanup-precorrection.py`. The execution chains and
wrapper identities are recorded in
`wave52-scoped/helper-provenance-clarification.json`, SHA-256
`076a4249e0e85b90237c7b1ccc6d12578f698b83933458116ed8e8f7e3792c48`.
Raw cleanup receipts remain unchanged.

Independent audit receipts are in `/home/zhuhe/.cache/litchi-api-contract/`:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `worktree-cleanup-wave52-audit-20261024.json` | 10,426 | `0b0c0ee25a2bd4f5f428ad9bc37c652eb7c2d419b32946a8ea8122876ee358c9` |
| `worktree-cleanup-wave52-helper-provenance-addendum-20261024.json` | 5,486 | `4e67cf7f1ede01c6269590b92403724534a934c014b066184d5c81e4743327ed` |

The addendum independently passes all 38 helper-provenance checks. No source,
target, raw receipt or quarantined payload was modified by either review.
