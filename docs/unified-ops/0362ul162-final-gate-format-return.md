# 0362ul162 — format the migrated BatchCommit test consumers

Corrected candidate C2 `e7aca04f06e46952cee223a78b1ad476927ac001` is frozen
after 161. The fresh 65-check driver passes its five source tests and dry
check. Its reviewed command sequence is unchanged apart from the candidate
identity: seven compiler checks, full53, and five runtime supplements, with
isolated rustdoc and the OfficeCLI preservation barrier.

The approved run then stops at compiler row 0. Rustfmt requires four
assertions in `docx_batches.rs` and `pptx_unified_write.rs` to be reflowed
after the BatchCommit accessor migration. No Cargo build, rustdoc or test
runs; the other 64 rows remain unrun. Root verifies the terminal logs and
traces, unchanged source and locks, and empty owned processes.

This commit applies exactly those formatting changes. No assertions or
production behavior change. The same rustfmt check passes across all 588
Rust files changed from the gate's fixed baseline. The next gate must use
a fresh committed successor and fresh state; the failed C2 packet remains
historical and cannot be relabeled.

External packet:
`/home/zhuhe/code/litchi-worktrees/.coord/final-gate-C2-e7aca04f06-20261026/`.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| Failed format log (in sibling state directory `compiler/00-fmt-changed-vs-9a3e-attempt-final-v5-compiler-00.log`) | 2,557 | `122b4951b02abe6028f532b9b92a5b082a4914f65642579c6dcbedabeb2a2470` |
| `root-terminal-verification.json` | 1,239 | `4c422716c17688faf390a0dea4e1aa63dabf83d4b9b6768bc5589646eed6e2d9` |

Native and raw-census preparation stops at its source-identity boundary.
No archive, native measurement or census capture uses the rejected C2.
Already emitted external C2 preparation remains preserved. Final gates,
census, fidelity and performance remain required; nothing is pushed.
