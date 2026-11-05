# 0362ul22 — merged DOCX validation leaf cleanup

Date: 2026-10-24. Scoped cleanup wave 49 removes the merged
`uo2-docx-artifact-stream-report` worktree at
`accdd972c462305dd55c42d598aecc516427c3b0` without force. Its branch remains.
The immediate branch, tip, ancestry, clean-state and process-holder checks
passed against integration `9781cf6d87e4d261d1dd9670e144e2db127d869b`.

Six ignored files, totaling 350,556 bytes, were quarantined outside the
repository. A reconstruction and restoration exercise verified every file's
hash and mode, parent-directory modes, exact ignored status and a second
non-forced removal. The validation target and evidence remain untouched.
The protected 15-byte trunk constructor probe remains unchanged.

The external evidence root is
`(deleted-coord-artifact)`:

| Artifact | SHA-256 |
| --- | --- |
| `manifest.json` | `6e2b215fd10284ef334dc23a0c1c7644602b91917e4ab800d5259177037f23b0` |
| `receipt.json` | `8fd388038af9fb4445c9069ec6e6aeedbd8f78469ac13e84ada98c542a322c84` |
| `restore-check.json` | `b41a8485acc666dd6ec98c796078a2e18502836381bdfca42e281d650e8d87d5` |

The coordinator independently verified those artifacts, all six quarantined
payloads and modes, retained branch ancestry, worktree absence and the
protected validation directory's identity. Its external receipt is
`(deleted-coord-artifact)/coordinator-wave49-verification.json`,
SHA-256 `1189c2a74c69cf541529ba3439a7e426f538aeb06664b3f2330ad6833ac3df9a`.

An earlier broad audit was stopped before any payload move or removal; its
helper files remain under the sibling `wave49/` directory. The completed
scoped audit inspected only the selected leaf and protected directory
identity. No target-size or filesystem-free-space reduction is claimed.
