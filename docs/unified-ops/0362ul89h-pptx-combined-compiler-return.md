# 0362ul89h — combined PPTX compiler return

This private source repair follows the actual isolated F5 failure on
`00a42a7d4c9bfead038e2e40c57a0c14291e5ecb`. It is compiler preparation;
PPTX runtime, glossary, golden-transcript and performance acceptance remain
open. No exception to the API contract or selected resource policy is made.

## Retained actual failure

State: `(deleted-coord-artifact)`.
The changed-path formatting row passed (384 paths, 2.780 seconds). The
all-features/all-targets PPTX check failed with exit 101 in 56.942 seconds,
reporting 207 diagnostics across 40 primary source paths. The final target
was 210,116,269 bytes; source and all three role locks were unchanged.
Root verified log/trace identities, isolated preflight, guard samples and
terminal ownership. Later rows were not run.

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `rows.json` | 31,269 | `328f19976ba13a79459d6954156dc4fed446ed376ecce80f6bfd9957762201e9` |
| `result.json` | 32,936 | `aa3cc96c973f2cd9d41be871a9c9796cf8a2666099064f18ca62e5734c6451c8` |
| `01-pptx-all-features-check-attempt-actual-f5-01.log` | 129,953 | `d7f08d8f7a99af9566b704d86cc9ca0066f28e745d2c0f5ac0cdd3e6ec11628f` |

## Composed repairs

The source retains public const limit APIs using const-compatible matches,
conditionals and minimum helpers. Checked counters and every selected owner
ceiling are unchanged. Shape/tag/Designer/classification replacement helpers
receive the selected output ceiling; their test consumers migrate together.
The review return fixes six missed Designer test callers and removes three
unused private default wrappers rather than retaining compatibility twins.

Other repairs restore crate-private/test-only helper visibility, resolve
root type qualifications, retain namespace decoders and QName lifetimes,
fix animation text and comment URI types, and copy caption limits before a
mutable borrow. None adds a whole-package scan or payload copy.

Coordinator composition also removes an identical duplicate semantic-limit
accessor, borrows replay-index part names until an owned key is needed, and
converts metadata byte-accounting errors to the properties error type while
preserving underlying OPC errors and internal-invariant reasons.

The following independent patches were reviewed and applied in order:

| Patch | SHA-256 |
|---|---|
| `(deleted-coord-artifact)/0362ul89g-pptx-61-compiler-repair.patch` | `948df522fb5ec75e65df1ccbc74549ca14e99ce35d4e412719254abdd4478fb7` |
| `(deleted-coord-artifact)/0362ul89g-pptx-61-designer-tests-repair.patch` | `a323f9bce865fe907a66e786f5f6aac927af009de895f8ec64598b1949b81674` |
| `(deleted-coord-artifact)/0362ul89g-pptx-61-obsolete-wrapper-cleanup.patch` | `e230d24de3653a98605759fb1e72543fb7fedd97161c23c38746da1044f59685` |
| `(deleted-coord-artifact)/pptx-f5-compiler-fixes-00a42-85e6aa3884.patch` | `eff29f8b0cdc3d73eab615fd07dbf33a2f8731715c4c318635276dac4e079cf3` |
| `(deleted-coord-artifact)/metadata-compiler-repair-81-dd6affd.patch` | `55e32a27541a133c41162613ff55305b046125ff51e586c61649cb1e6a2fead8` |
| `/tmp/uo2-pptx-tracks-borrow-00a42-5bf328754a.patch` | `069a6b062d51ba0fb25d0dce5120aa233c0b25fedc7a1d92be064ccc77c56175` |

The deleted external paths were beneath `/home/zhuhe/code/litchi-worktrees`.
Direct rustfmt on all 53 changed Rust paths and `git diff --check` pass.
A fresh F5 run on the committed composition is still required; the retained
failed capture is not reclassified by these source checks.
