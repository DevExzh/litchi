# 0362uk — DOCX encrypted transaction admission

The frozen candidate `f6fe430a4d19d94aa360c0e9ed03ebd757265225` allowed an
unlocked encrypted DOCX's deferred editor to be constructed.  The unified
transaction therefore entered the user closure before the leaf publication
policy rejected the operation; the final acceptance test intentionally
panicked at that closure.

`Docx::edit_policy` now checks the eager package's retained encryption
provenance in O(1), before constructing an editor.  It returns the fixed,
content-free `UnsupportedCapability` refusal used by the DOCX capability
matrix.  The signed transaction fast path is unchanged, so exact signed
no-ops still reach their source/signature proof, and locked inputs still
return `DocumentProtected` before a closure.  The existing publication fence
continues to reject plaintext output and no-op replay to an unlocked encrypted
owner.

No format capability or performance claim is made here.  The focused
regression was the existing
`unlocked_encrypted_input_refuses_without_dropping_publication_protection`
test; the complete DOCX transaction integration test was also run with and
without encryption.

## Checks

All commands used Rust 1.95.0, `--offline`, one Cargo build job, and the
external target
`/home/zhuhe/code/litchi-worktrees/.coord/docx-encrypted-target`.
The initial size transcription, 10,585,881,195 bytes, was incorrect;
[0362uk1](0362uk1-docx-encrypted-locked-receipt.md) records the measured
1,058,588,195 bytes. These initial diagnostic checks omitted `--locked`.
The ignored lock remains a required pinned input; 0362uk1 supersedes that
omission with locked reruns and verified lock, configuration and command
identities.

* `cargo test -p litchi-ops --features 'docx encryption' --test docx_transactions unlocked_encrypted_input_refuses_without_dropping_publication_protection -- --nocapture`: 1 passed.
* `cargo test -p litchi-ops --features 'docx encryption' --test docx_transactions -- --nocapture`: 22 passed.
* `cargo test -p litchi-ops --features docx --test docx_transactions`: 20 passed.
* `git diff --check`: passed.

The full acceptance gate remains the coordinator's responsibility; this
packet contains no gate logs or compiled artifacts.

## External receipts

| command | receipt | bytes | SHA-256 |
| --- | --- | ---: | --- |
| focused encrypted regression | `/home/zhuhe/code/litchi-worktrees/.coord/docx-encrypted-focused.log` | 10691 | `94e74236ea89ac2aaa82e4c121e013e035fa6a4f20896787cca2dcb20e1ba3e2` |
| encrypted DOCX transaction suite | `/home/zhuhe/code/litchi-worktrees/.coord/docx-encrypted-docx-transactions.log` | 2426 | `674c85f7dc42c5b2ce01e4d4c1454aa0f3a77f58cfbd02daa8358146d2975fd1` |
| non-encrypted DOCX transaction suite | `/home/zhuhe/code/litchi-worktrees/.coord/docx-encrypted-docx-no-encryption.log` | 3898 | `ea3a343807536d8eed6071af38b15359ea7fd832dc1ff7fed0e3b285ba93d833` |
