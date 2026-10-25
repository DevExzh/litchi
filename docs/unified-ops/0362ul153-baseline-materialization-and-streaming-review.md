# 0362ul153 — baseline materialization and streaming review

The signed streaming repair returned in [152a](0362ul152a-signed-archive-memory-return.md)
passes independent review and all 16 focused tests. Artifact identity now
uses bounded SHA-256 reads; tar members use bounded Git-blob and ignored-lock
digest reads. Exact tree, mode, duplicate, path and ignored-file checks remain.
The regression includes a 192 MiB sparse artifact under a 128 MiB address-space
limit and an intercepted member reader that rejects unbounded reads. The
patch is ready but has not changed frozen protocol P.

External repair packet `signed-archive-streaming-repair-v1` contains:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `prepare_signed_native_inputs_v2.patch` | 4721 | `1ad698901910475717da3f393d04987c82c2d67e2bfcf83898e0236fc7682a43` |
| `test_signed_native_policy_v2.patch` | 6627 | `1656a5b656054175babd923e398700adccd51f09d20203d7f2c629eb505a53dd` |

The separate signed v5 archive-reuse recipe passes 12 source-only tests but
is returned before execution. Its standalone finalizer checks path equality
without rejecting nested output paths. The successor must reject overlap
between P, the immutable generation root and final output before writing.
Reuse must also compare the input and verified staged hashes immediately and
check stable source metadata. The failed v4 tar remains unverified and has
not been read, copied or hardlinked during these reviews. Its failed receipt
remains unchanged.

DrawingML materializer v5a resolves the four returns in 152a. Full C may have
no sparse metadata; newly created worktrees admit Git's generated metadata
before installing the reviewed template. Common configuration and the
relevant source heads remain protected, while unrelated ref changes are
recorded. Final checks precede the final profile map and completion event.
Thirteen focused source/seam tests pass independently. The generated-driver
seam is refused by the existing unrelated-builder guard and has no new pass
claim under load.

Root's actual read-only plan confirms ten pairwise-disjoint profile roots:
four retained roots and six new sparse worktrees. It also confirms frozen
C=`f8cd398a770f7363f628be40edb3ce4219566c16` and
P=`4ab42bb9678769930d7c17173827fe5693370ab6`. The external plan is
`drawingml-v5a-root-materialization-plan-20261025.json`, 74,193 bytes, SHA-256
`f2a5bfaaa5e5345780760bc35d5132efa91a7b3dd4aa00aa747d9be4493cffc2`.
The independent review receipt is `drawingml-recipe-v5a-review-20261025.md`,
SHA-256 `9e978fc50d40099b3f157fa8058413656b71305d57e34e424ad7e96d3151ee2a`.

Root starts the reviewed source-only materialization under
`litchi-unified-drawingml-cf8cd-baselines-v5a.service`, with MemoryMax=4G and
MemorySwapMax=0. Its external control directory is
`drawingml-Cf8cd398a-baselines-v5a-control-20261025`; the fresh output root is
`drawingml-Cf8cd398a-baselines-v5`. The exact executable is
`native-final-bindings-20261025/drawingml-recipe-v5a/materialize_drawingml_baselines_v5a.py`,
SHA-256 `1ef65bd414064003d2f1b9e26a5610f7a7061115d4879d8c65ff2d8562ecac6a`.
Regenerate only at fresh paths after preserving the prior state, using
`PYTHONDONTWRITEBYTECODE=1 python3 -B <materializer> --execute`.
The job preserves the four old workspace locks and their stat records before
replacement, creates six new roots, and validates all ten source/adapter/lock
sets. Launch is not completion; terminal verification remains required.
P and C stay frozen throughout this job as well as the current-native emitter.

Root independently verifies 32 completed full65 rows: seven compiler
supplements and full53 rows 0–24. Logs and target traces match their hashes;
source/lock identities, exit statuses, target limits and empty owned-process
postconditions pass. The external partial receipt
`coordinator-153-partial-gate-verification.json` is 7,053 bytes, SHA-256
`87928dfb2f60af63adcbdbdd8e5b15411298ce336c1574f8f5517eeb4d5c4cd9`.
The full gate and current-native emitter remain live. No additional Cargo or
native run is admitted while the full gate owns the build slot.

All external paths above are beneath `/home/zhuhe/code/litchi-worktrees/.coord`.
Common-P installation, terminal gates, native fidelity/performance, census
review, final trunk integration and cleanup remain open. Nothing is pushed.
