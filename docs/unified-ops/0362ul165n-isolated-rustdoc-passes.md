# 0362ul165n — Isolated rustdoc passes

The isolated `wsdoc` rerun passes on
`472a89a3fae1cbe81d177b6c2f139e03d788bf43`. It takes 600.75 seconds and leaves
7,783,115,047 bytes in the bounded target. The standalone PPT check passes
next in 97.71 seconds, with 7,932,811,789 target bytes afterward.
The subsequent core/ops/facade documentation row also passes. Root verifies
its log, target trace, unchanged source/locks and successful exit; the same
diagnostic proceeds to CLI tests.

Root verifies both completed rows, their log and Cargo-target trace hashes,
unchanged source and locks, successful exit codes and absence of target
resource errors. During documentation, the observed builder processes are
one Cargo process and its rustdoc child in the Q worktree. Other workers are
source-only. The queued DOCX watcher starts no Cargo work during this run.

This closes the outstanding isolated documentation rerun. It does not
qualify the final source's all-65 gate. The same remaining-29 diagnostic
continues; its earlier leaf/workspace failures are the two DOCX assertions
already repaired on the isolated `ec243ebb44` leaf. Focused validation of
that leaf is queued after the complete diagnostic becomes terminal.

W6's incomplete v7 packet is preserved and a fresh worker prepares a v8
successor for the four concrete v6 findings. Protocol composition v12 carries
the five merged ODF admission files exactly. Root verifies all fifteen
copies across its three output trees, each tree's exact five-file delta
from v11, and unchanged signed guard, patch and phase inputs. All 22 tooling
files changed since C33 are represented: 18 match current root bytes and
four retain their reviewed composed migrations. Installing these consumers
therefore preserves the merged ODF admission. The pinned 9feb merge proves
these source bytes; it does not select final C, which must include the
validated DOCX repair and any further accepted source changes.
Neither preparation constitutes actual native qualification. Final C/P
selection, all 28 qualification groups and final acceptance remain required.
Nothing is pushed.

Root's verification receipt is
`/home/zhuhe/code/litchi-worktrees/.coord/rest29-isolated-wsdoc-verification-20261026/receipt.json`
(2,517 bytes, SHA-256
`a5d2dacce7f539a313fecc3f3089700872eb1324ad39f9d266190d0c6e443de4`).
It includes the exact commands, source/lock identities and log identities;
the adjacent `row-34.json` and `row-35.json` preserve the completed rows.

Additional receipts under the same `.coord/` directory:

| Artifact | SHA-256 |
| --- | --- |
| `rest29-core-ops-facade-rustdoc-verification-20261026/receipt.json` | `eeb2826d1c384f6eea52f95ac9f88d25f6ea8b52271dae380cc92217d97e7845` |
| `protocol-consumer-composition-C2a-20261026-v12/receipt.json` | `02f784d92608168099c60c94c8271992a2bdc59a19fb6824585cf1838076e6e4` |
| `protocol-composition-v12-root-source-verification-20261026/receipt.json` | `40d38879e9db11f57a94b5d1241790e5ff57e2f790d49624d9fc5139ef744348` |
