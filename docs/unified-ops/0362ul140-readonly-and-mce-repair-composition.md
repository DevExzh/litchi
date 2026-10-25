# 0362ul140 — readonly and MCE repairs composed

The frozen C minimal PPTX diagnostic passes 1,395 library tests (one existing
ignored test) and integration chunks1–30. Chunk31 reproduces the exact
readonly foreign-base no-op assertion from record134; its other four
readonly tests and the revision test pass. Root verifies the failed terminal
capture before changing source. Its final target occupies 487,056,388 bytes.
Evidence under the worktree parent's
`.coord/ooxml-pptx-joint-c780cf4c7-f5-runtime-pptx-minimal1-v1/` is preserved:

| File | Bytes | SHA-256 |
|---|---:|---|
| `rows.json` | 10,937 | `45494903c7feaf96a68197b979105fc2804d4f2e9e31c8ae56ace4e9c71826b4` |
| `result.json` | 14,034 | `4899f16d03589c333ea7a06e99ce8632aba7a6a2a61fde8db80430744a912ccd` |

The capture is reproducible with its source-bound `run_f5.py --go` in fresh
state/target directories; `.coord/checkpoint96/verify_capture.py STATE`
checks the terminal evidence. The helper exits1, so the diagnostic adapter
correctly refuses continuation as an ordinary direct-Cargo exit101 failure.
No failed result is relabelled.

Root then merges the two independently reviewed repairs into private
integration without conflicts:

- `4078b4350b7348e080c6fec4c6a6a6770760fffb`: readonly tests construct the
  no-op patch from matching content and explicitly reject a foreign base.
- `d39e0adc0bb88d71dedc6405a47a185ce60a3236`: the shared rustdoc resolver
  proves the public MCE limits re-export through its audited defining path.

The composed commit is `e7841dd222f17a00abe79c5ae034ad18cde3f45d`, tree
`997c04a1e5b5014131c99d678508cb5617bdaa7f`. All five changed files exactly
match their reviewed leaves, and all three approved role locks are unchanged.
The external composition proof is
`.coord/runtime-repair-composition-0362ul140.json`, 2,081 bytes, SHA-256
`162a419604586af62732c67eb76e93d82fbded25e7e14f31539b7df939f65b92`.
The integrated glossary/raw-coverage suite passes all 75 Python tests.
Repaired-source Rust runtime qualification remains pending.

Final gate planning now avoids duplicated broad runs. The approved full53
includes the complete 69-profile glossary loop and all broad runtime checks.
It will run once on the final composed source, with six distinct compiler
profiles and five distinct runtime profiles added. Isolated wsdoc row34 and
core/ops/facade rustdoc row36 remain mandatory. Exact command coverage and
the unchanged resource guards are recorded in
`.coord/final-gate-coverage-780cf4c7/coverage-map.json`, 19,676 bytes, SHA-256
`059eb2c0f076dd5c6072482255aa950cd1e135325367afe773a8f7e8d40b0bde`.
This map is a plan, not an execution receipt. The C-only outer16 and minimal
tail packets remain inert; their required coverage will run on repaired
source in the final gate/supplement union.

Native execution work remains open: 135c must use genuine baseline and
candidate binaries, the existing five-block legacy workload and validated
route output; 136b must implement the actual v2 signed producer and runner;
137c must use the separately pinned protocol checkout and validated probe
inputs. Their prior parsing tests do not establish native execution.

The census matrix and final source bindings will be refreshed after these
required protocol changes compose. Reviewed merged leaves are queued for
cleanup. Root remains on `feat/unified-ops-api`; trunk integration, final
gates, API census and fidelity/performance are not complete. Nothing is pushed.
