# 0362ul89n — PPTX private import return

The actual isolated check of `6517f7f5c1824c95cd1fd793692de0c680cebd91`
passes formatting over all 390 changed Rust paths, then stops on two unused
imports. The earlier 113 dead-code diagnostics are absent. The all-targets
check has not yet completed, so runtime/Clippy/glossary/performance acceptance
remains open.

The media codec's `MAX_STRING_BYTES` import is needed only by its existing
cfg(test) constructor and now has the same condition. The Designer module's
two unused crate-private setter/remover reexports are removed; the actual
functions and their selected-limit consumers remain in `properties.rs`.
Exact symbol searches confirm no caller uses the removed reexports. No
public method or execution path changes.

Retained state:
`/home/zhuhe/code/litchi-worktrees/.coord/pptx-combined-0362ul89-f5-actual-6517f7f5c1`.
The compiler row exits 101 after 55.260 seconds, final target 209,550,542
bytes. Root verifies the actual command, clean source/three-lock fences,
log/trace hashes, isolated preflight, live guard and terminal process samples.

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `rows.json` | 31,627 | `a7eea7ebd4e1310413d6f5f80a9141b6f32eb1bf819528dbf58b6cf653af38bd` |
| `result.json` | 33,306 | `9ac63cef5bf7cd6d172c7a460b782af2a13865a9e47e2b96bf5c40fdb4619f16` |
| `01-pptx-all-features-check-attempt-actual-f5-01.log` | 25,637 | `fd76ddfde5ff26656f7e7fd402bd976a159bdc6bc942c21d87323b9dae5bb6ee` |

Direct rustfmt and diff checks pass. The failed capture is retained unchanged;
a fresh compiler run is required on this committed repair.
