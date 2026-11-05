# 0362ul79f — raw-census tooling-row registration

This source-only leaf is based on integration commit
`88ca2fd38fdd23ca2ee4ad22449a83893804e4e1`. It registers the new
`tools.test_prepare_raw_api_census` module in the canonical tracked safe-gate
runner's existing `tooling-unit` row. The active approved runner worktree and
the external `.coord/gate.sh` were not modified; the coordinator applies the
tracked change and the external patch separately after review. Existing F5
runner hashes and all historical capture receipts remain immutable.

## Exact tracked runner delta

Only the `tooling-unit` argv changed. The new module is immediately after
`tools.test_raw_api_coverage` and appears exactly once:

```text
normal plan steps: 53 -> 53
tooling modules: 39 -> 40
```

The source runner before this leaf is 52,787 bytes with SHA-256
`e88893d497f7e7ef70bde4314c3e57747d20997f2ea9d82b4d4595edbbe8458`.
The prepared runner is 52,824 bytes with SHA-256
`5db2a14fb79ad987ff4c92fb0da4dfbdabb9ff1990f32d6d4f0205a54e6cf181`.
Removing the one module token leaves the runner byte-identical, including
the 53-step plan, target-guard setting, ownership, cleanup and cap policy.

The external coordinator patch is retained at
`docs/unified-ops/results/0362ul79f/gate-sh-tooling.patch`. It is a one-token
insertion into the current external `.coord/gate.sh` `toolunit` command. The
external file was not edited: its observed pre-patch size/hash are 6,249
bytes / `23170aedf5583e637b48eacb79f7976eef65c83e3fbe9f3c3749ddd807e1671c`,
and the expected post-patch size/hash are 6,283 bytes /
`8edbc9fc212f272c28ba33f3646812b2129625c4edd17692e76f3cfb99e813cc`.

The coordinator must review and bind the prepared runner hash before using it
for a later final gate. This leaf does not claim that the newly registered
module, the updated runner, or the external gate has passed a full gate.

## Build-free verification

The source-only checks were:

```text
python3 -B -m py_compile docs/unified-ops/results/0362ul29/tooling/run_safe_gate.py
python3 -B -m unittest tools.test_prepare_raw_api_census tools.test_raw_api_coverage
python3 -B docs/unified-ops/results/0362ul29/tooling/test_run_safe_gate_ppt.py
```

The runner AST check reports 53 normal steps, one `tooling-unit` step, and 40
unique registered modules with the new module at the required position. No
Cargo, rustdoc, glossary CLI, native, timing or full-gate command ran.
