# 0362ul76 — raw-API census tooling-row registration

This source-only leaf is based on integration commit
`6b95d8955412c68821c7b961335ab93c79ebd06b`, which contains the reviewed raw
callable census from 0362ul71. It registers that census's Python unit suite in
the tracked safe-gate tooling row. The active approved runner worktree and the
external `.coord/gate.sh` were not modified; the coordinator applies the leaf
after review. The separate DrawingML-v2 registration remains out of scope
until its source is merged.

## Exact runner delta

Only the `tooling-unit` argv in
`docs/unified-ops/results/0362ul29/tooling/run_safe_gate.py` changed:

```text
added: tools.test_raw_api_coverage
position: after tools.test_final_api_inventory, before tools.test_w8_file_provenance
module count: 35 -> 36
normal plan steps: 53 -> 53
```

The prior runner is 52,641 bytes with SHA-256
`b99eed79df58011c2d2cdff817e914241e63089d304e186d24016d2677f40503`. The
leaf runner is 52,672 bytes with SHA-256
`5536aa2a0ef36c6152d7191b2c1fe2eb4c39cdbbbab9c4e482812c5260e09b3d`.
Removing the changed tooling-unit line from each version leaves the runner
byte-identical, including `run_step`, ownership, target guards and cleanup.

## Source-only verification

The leaf ran:

```text
python3 -m py_compile docs/unified-ops/results/0362ul29/tooling/run_safe_gate.py tools/raw_api_coverage.py tools/test_raw_api_coverage.py
python3 -m unittest -v tools.test_raw_api_coverage
Ran 11 tests ... OK
python3 -m unittest -v docs/unified-ops/results/0362ul29/tooling/test_run_safe_gate_ppt.py
Ran 3 tests ... OK
```

An AST/constructor check reports 53 normal steps, one `tooling-unit` step,
and 36 registered test modules with `tools.test_raw_api_coverage` present.
No Cargo, rustdoc, native, timing or full-gate command ran.

Coordinator integration also updates the external `.coord/gate.sh` tooling
row to the same 36 modules. Its prior bytes are retained at
`.coord/gate.sh.before-0362ul76` (SHA-256 `1ad5e255780c85a7bdc468414d64d4a288f2d484c80dee8fab5872862b732876`);
the updated script has SHA-256 `33b8dbad688f8a9a221bfb85eb668fddc1a789193a6096425182e7045628d179`.
The original approved f5 runner worktree and active captured runs remain unchanged.
