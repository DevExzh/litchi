# 0362ul147 — raw-census Cargo admission return

The first actual C-bound v9 census run terminates after two passing rustc
identity steps. Step 2, `final-api-inventory`, exits 2 before compiling the
inventory. Its exact error is that the Cargo path is not the coordinator cap
shim: the adapter supplied the approved outer tracing wrapper at
`uo2-gate-runner-self-attribution/docs/unified-ops/results/0362ul29/tooling/bin/cargo`,
while the inventory admits the coordinator cap shim. The trace records only
the Cargo version probe. No raw profiles or inventory captures are produced.

Root verifies all three row log/trace hashes, unchanged C and all three
locks, unchanged initial binding and manifest hashes, no live-target guard
errors, zero target bytes and empty owned-process/builder lists in every
post-exit sample. The failed run remains at
`.coord/raw-census-Cf8cd398a-20261025/` under the worktree parent:

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `run/result.json` | 11,433 | `f35eb32901383b1f0f18253423e03f7ddd45b545720e2b5f0413d4a812a6aec5` |
| `root-terminal-verification.json` | 771 | `f3891030c09d5ae95523c6026f078b4d13914a56fab238e4f76f5a03e1ac98b0` |
| `run/state/02-final-api-inventory-attempt-f5.log` | 184 | `9a56b333f77445bee57027b5e53bb82041404a5df48a99a5808281cc6d0b6c51` |

Return v10 must correct the inventory command's admitted Cargo path while
preserving the outer F5 source, process, memory, target and binding guards.
It must exercise the real inventory admission seam in source-only checks and
use a fresh packet for its next actual run. Frozen candidate C
`f8cd398a770f7363f628be40edb3ce4219566c16` is unchanged. The inventory's
production admission rule is not waived or relaxed.

The build slot is idle after verified termination. The full gate's executable
v5 driver and native protocol fixes continue separately. Final capture,
callable review, strict finalization, all other gates, fidelity/performance
and trunk integration remain open. Nothing is pushed.
