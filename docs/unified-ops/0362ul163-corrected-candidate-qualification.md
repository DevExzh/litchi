# 0362ul163 — corrected candidate qualification and proof preparation

Candidate C2a is frozen at
`33a0a63d3553712c1276da592f7846a0c7dcabd6`, tree
`656d66603b33aec273d529263aa10b504bd3b7f8`, in the integration worktree.
The protocol worktree remains at that same pending revision. Subsequent
coordinator documentation on trunk does not advance either frozen checkout.
The candidate identity is external at
`.coord/candidate-C2a-33a0a63d35-20261026/c2-identity.json`, SHA-256
`73754fd078e462678df0025c913e85f7323910d115c288f5a0fc641e9e91e7e3`.

## Full gate

The fresh reviewed full65 driver is running under
`litchi-unified-final65-c2a-33a0-v1.service`. Its source-only tests and dry
check pass; its executable AST differs from the reviewed driver only in the
candidate identity. The source selection contains 588 Rust paths. Coverage
remains seven compiler checks, full53 and five runtime supplements, including
isolated workspace rustdoc and the row-37 OfficeCLI preservation barrier.

Root verifies nine completed rows: all seven compiler supplements,
workspace formatting and all-feature workspace Clippy. Logs and target
traces match their recorded hashes; source and locks are unchanged and
completed rows have clean resource/process guards. The remaining rows are
not accepted until they complete. The earlier e7aca formatting failure stays
preserved as a failed run.

External packet `.coord/final-gate-C2a-33a0a63d35-20261026/`:

| Artifact | SHA-256 |
| --- | --- |
| `plan.json` | `e76fed802a2b87072285369dfae1286e90bde5107d919439fa2d74853bf24ec2` |
| `root-prelaunch-review.json` | `80a56d2a86c9b91224ba6efa2029235b585173e37bb49d38b500bf3524733e0a` |
| `root-verified-compiler-and-format.json` (2,346 bytes; nine rows) | `d0e90626cb59e501b66b4b13b76e5d0025b164832c4faa9ed57ea18e486d4eb6` |

## Source-only preparation

The following packets pass independent or coordinator source review. None
is a native measurement, final census capture, archive verification or
terminal OfficeCLI acceptance result. Paths below are relative to the
external `/home/zhuhe/code/litchi-worktrees/.coord/` directory.

| Preparation | Final review receipt | SHA-256 |
| --- | --- | --- |
| Current-native C2a v2 | `current-native-c2a-independent-followup-20261026-v2/review-receipt.json` | `0408798fc82ae8602f4bf262e3914c4eb077b2ad9f7a83423b2ceefbe099cfc9` |
| Signed-native C2a | `signed-native-C2a-author-final-recheck-20261026/review-receipt.json` | `10ce1648371ba19e5c61167f7a8d1d816523be89c75a2fef8189658947eebcc7` |
| DrawingML C2a | `native-final-bindings-20261026/drawingml-recipe-c2a-33a0a63d/root-correction-review.json` | `931d382fd0e82e4b211d640c34138998fef626b23c205ec9990fcc2e43478403` |
| Raw census C2a | `raw-census-f5-c2a-33a0a63d35-20261026/independent-review-receipt-drawingml-v2.json` | `c5a49d3fa55feb8b57915413c28a2c6dbe04a4712000faffe7b316dadc801cb6` |
| OfficeCLI C2a postcheck | `officecli-C2a-root-source-review-20261026/review-receipt.json` | `5a33e5f620041f807b9c7b2f23d26450927529da510b5b8a15b68da69ce588c4` |
| Notes memo C2a | `pptx-notes-memo-f5-c2a-33a0a63d35-corrected-c/independent-followup-review-receipt.json` | `4622230dec26acff7561f5034aa9fd48785bfe52c49d929e7fc812465cfb976b` |
| Replay scanner126 C2a | `0362ul126-pptx-replay-perf-C2a-33a0a63d35-final-c-independent-review-receipt.json` | `bfef153e0f89edfa598a51b489e37681ddfade6ac026786f07099b06d8839c1a` |

Review corrections include stale receipt metadata, consistent archive
instructions, complete consumer pins, an unambiguous DrawingML patch
composition, and correctly sealed final packet inventories. Notes and
scanner measurements bind frozen C rather than the protocol checkout, which
will advance during installation. Their source-diff descriptions now
distinguish all 134 changed paths from the six-path filtered performance
scope. No runtime behavior or acceptance threshold is weakened.

Historical failed preparation remains identified as such. The current-native
predecessor retains its already-disclosed stale receipt. One superseded
scanner author receipt (`29137…`) was overwritten during metadata repair
and its original bytes could not be recovered; the final packet discloses
this instead of claiming preservation. The independent failed-review
receipt remains intact. No executed measurement relies on either packet.

The census still needs 68-package inventory, 84 profile captures, callable
review and strict finalization. Native policies 72/74/75 need fresh candidate
archives, coordinated protocol installation, build/smoke/fidelity checks and
quiet ABBA measurements. The ten verified DrawingML baseline roots are
reused. Notes and scanner126 measurements remain required. All these heavy
jobs wait for the full gate's exclusive slot. Only memo 22 is deferred;
nothing is pushed.
