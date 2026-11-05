# 0362ul41 — PPT tooling-row import and registration

This record binds the current tooling-row wiring to integration source
`30aaafbf98e097aeec8cf6fe3df6caaae62a37d6`.  The leaf is
`uo2/ppt-tooling-row`; its source change is commit
`49e7138361f9a701896d6d1be8bd87cec01d8989`.

## Change

`tools/test_ppt_lazy_native_qualification.py` now imports its sibling as
`from tools import ppt_lazy_native_qualification`, with a direct-script
fallback.  The `tooling-unit` argv in
`docs/unified-ops/results/0362ul29/tooling/run_safe_gate.py` now registers
`tools.test_ppt_lazy_build_receipt_adapter` and
`tools.test_ppt_lazy_cold_first_commit` before the existing PPT native suite.
No `PYTHONPATH` override was added.

The pre-fix reproduction was the exact three-module command, with
`PYTHONPATH` unset:

```text
env -u PYTHONPATH python3 -m unittest tools.test_ppt_lazy_build_receipt_adapter tools.test_ppt_lazy_cold_first_commit tools.test_ppt_lazy_native_qualification
```

It discovered 12 tests and failed at import with
`ModuleNotFoundError: No module named 'ppt_lazy_native_qualification'`.
The preserved log is
`/home/zhuhe/code/litchi-worktrees/.coord/ppt-tooling-row-30aa-three-before.log`
(923 bytes, SHA-256
`1cf9f5984dae8db8db232ef3501c5a392d6e9d30922c5d67def7ee19e955315d`).

The focused post-fix command discovered 26 tests and passed in 0.173 seconds;
the preserved log is
`/home/zhuhe/code/litchi-worktrees/.coord/ppt-tooling-row-49e7138361-focused.log`
(153 bytes, SHA-256
`d9ea5d32b930a9a1d6a78bda0b8e8fc83f22e6c7df267c6f720ddb4e352b59b5`).

## Full-row baseline receipt

Before the wrapper-environment repair, the complete current 33-module
tooling argv (the `tooling-unit` argv at the source commit) was run directly
with both `PYTHONPATH` and `CARGO_TARGET_DIR` unset.  It ran 1679 tests in
107.146 seconds and had exactly two failures; no Cargo or native build was
run.  The preserved log is
`/home/zhuhe/code/litchi-worktrees/.coord/ppt-tooling-row-49e7138361-direct.log`
(640391 bytes, SHA-256
`67c8bf98625a046d8e24adbd01f234b8522fc8765728f44e2122ffa47556fc98`).

The two failures are existing generated-data checks, not PPT import failures:

* `tools.test_officecli_capability_delta.ApplyingDispositions.test_the_committed_tables_pass_check`
  reports 11 stale generated OfficeCLI pages and asks to rerun
  `officecli_capability_delta.py`.
* `tools.test_unified_bench_matrix.Committed.test_committed_outputs_are_current_and_complete`
  reports 17 recent evidence directories with no benchmark rows.  These
  records are source/tooling or protocol evidence and must be classified
  accurately; this receipt does not invent performance rows.

The first run through `run_safe_gate.py --only tooling-unit` is retained for
environment diagnosis, not claimed as a pass:
`/home/zhuhe/code/litchi-worktrees/.coord/ppt-tooling-row-49e7138361-run.log`
(3797 bytes, SHA-256
`f55d8d67733aa152bc6d559a35eb34aeec9676155731a9756cd1c3af02f17fd7`), with
child output
`/home/zhuhe/code/litchi-worktrees/.coord/ppt-tooling-row-49e7138361-state/48-tooling-unit-attempt-1792884027013552574-2916606.log`
(651459 bytes, SHA-256
`d1b09681fd27ebbaeaaee160c78464b2ff958ad631514732e3fada8fc36048a6`).
That wrapper injected `CARGO_TARGET_DIR` for the Python row.  The native
qualification tests consequently raised
`RunnerError: manifest cannot override the explicit CARGO_TARGET_DIR` at
`tools/native_qualification_runner.py:2604` (assertions at
`tools/test_native_qualification_runner.py:340, 770, 1265, 1296, 1640`).
The wrapper environment must be repaired before a wrapper-row pass is
accepted; Cargo steps retain their explicit target guard.

The scoped snapshot/matrix and wrapper repairs were committed as
`02080ded637029036dd03764c9f86495a9a6e5fb`,
`1e5f80db71a08535a4a4b867e74e532683ba98ff`, and
`8d1184fc38ad8e5c81fb428c7f50613f210c0495`:

* the 17 non-benchmark result directories are explicitly classified in the
  curated matrix as source/protocol/tooling evidence with no invented rows;
* only the Python `tooling-unit` step omits `CARGO_TARGET_DIR` and
  `LITCHI_GATE_TARGET_ROOT`; Cargo-producing steps retain both the target
  environment and the runner's subreaper policy; and
* the tooling child is explicitly the nearest subreaper, so its lifecycle
  tests own and reap their adopted descendants.  The W8 cleanup retains
  identity-checked termination and fail-closed diagnostics.

## Final wrapper-row receipt

From clean commit
`8d1184fc38ad8e5c81fb428c7f50613f210c0495`, the exact runner command was:

```text
env -u PYTHONPATH -u CARGO_TARGET_DIR python3 docs/unified-ops/results/0362ul29/tooling/run_safe_gate.py --root /home/zhuhe/code/litchi-worktrees/uo2-ppt-tooling-row --target /home/zhuhe/code/litchi-worktrees/uo2-ppt-tooling-row/target --state /home/zhuhe/code/litchi-worktrees/.coord/ppt-tooling-row-8d1184fc38-state --only tooling-unit
```

The current 33-module `tooling-unit` row passed: 1679 tests in 105.38685154914856
seconds, exit 0.  The wrapper log is
`/home/zhuhe/code/litchi-worktrees/.coord/ppt-tooling-row-8d1184fc38-run.log`
(3674 bytes, SHA-256
`cf22d2ea0d98fc7cf377a954e88a96ea166bdc8a75166b4a09b5e1621062112e`); its
child log is
`/home/zhuhe/code/litchi-worktrees/.coord/ppt-tooling-row-8d1184fc38-state/48-tooling-unit-attempt-1792885613938334100-3447097.log`
(634981 bytes, SHA-256
`b70a9c2cefc35a27258df375aaaf2af10dbb0babd6b3851151b04e927de29357`).
The runner receipt records `cargo_target_trace=[]`, zero target bytes, and
exit 0; this was a source/fake-Python tooling run and did not invoke Cargo or
native qualification.  The earlier failed receipts above remain immutable
diagnostic history; this final receipt closes the tooling-row requirement.
