# 0362ul139 — PPTX tail passes; native execution returns

The full-feature PPTX tail on frozen candidate
`780cf4c7eebd62784670e2ab14cc49088a4bc91c` passes all 26 rows: integration
chunks32–53, all six examples in three chunks, and doctests. Root verifies
the terminal capture, source and three lock identities, command/log/trace
bindings, process exit and target bounds. Final target usage is
1,460,761,808 bytes. This diagnostic result supplements the earlier runtime
capture; it does not erase its readonly test failure or grant final acceptance.

Evidence is outside the checkout at the worktree parent's
`(deleted-coord-artifact)`:

| File | Bytes | SHA-256 |
|---|---:|---|
| `rows.json` | 104,299 | `3adfb500d9107f648413b383ec36a0ffbc2307f935ea1ba2ba0c52596a9bf309` |
| `result.json` | 110,487 | `4d2ac97fb35ca2ebb9d674f8411a36576e4e9a328267d1e677542c8220476db8` |

The source-bound packet is reproduced with `python3 -B run_f5.py --go`
using fresh state/target directories and its recorded source, locks and
runner. Root verification uses `(deleted-coord-artifact) STATE`.
The separate minimal-feature helper has started after root preflight; its
result and remaining minimal tail are still pending. The sixteen outer
diagnostic groups are also pending on this source.

Return138 is resolved in source by
`d39e0adc0bb88d71dedc6405a47a185ce60a3236`. The shared rustdoc resolver maps
the audited public path `litchi_ooxml_common::mce::Limits` to its defining
path `litchi_ooxml_common::mce::model::Limits`. It still requires one nominal
identity; wrong defining paths and ambiguous identities fail closed. Both
checker consumers migrate in that commit. Independent review passes the
captured DOCX minimal check and all 75 glossary/raw-coverage Python tests.
This is a checker repair, with no production API change. It and readonly
return134a remain ready for composition after the frozen-source diagnostics.

Native protocol review requires actual execution and evidence consumption:

- 135a corrects baseline role-lock pins and binds lock validation to the
  source root. Follow-up135b must connect the PPT and eight legacy routes to
  their real runner and receipt writer; a check-only probe cannot produce
  the execution receipt or measurement streams that the validator requires.
- 136a must require route evidence, prove immediate signed-batch refusal
  through owner execution, reject wrapped v1 probe output, recompute trusted
  generated inputs, and compile the API branches it claims to cover.
- 137a repairs profile coverage, executable imports and v2 execution-policy
  boundaries. Its 43 focused Python tests pass; independent review remains
  pending before composition.

These are required engineering returns, not accepted exceptions or native
measurement passes. Production source stays frozen during diagnostics.
Final workspace/wsdoc gates, the raw API census, fidelity/performance,
trunk integration and remaining leaf cleanup stay open. Nothing is pushed.
