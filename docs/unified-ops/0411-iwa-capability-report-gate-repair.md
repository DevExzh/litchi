# 0411 — capability-report gate repair

- Status: generated-report repair on `uo2/iwa-capability-report-gate-repair`,
  based on trunk `3ce2a06774`.
- Ownership: generated OfficeCLI capability-delta output only. No generator,
  golden, production, corpus, or central-progress changes are included.

## Regeneration

The canonical command was run from the isolated worktree:

```text
python3 tools/officecli_capability_delta.py
python3 tools/officecli_capability_delta.py --check
```

The only generated diff is `docs/unified-ops/officecli-capability-delta.md`:
three current iWork registry input hashes changed (Keynote, Numbers and Pages),
the Numbers/Pages element counts increased by one, and the regenerated tables
add the current `chart` litchi-only rows. Registry goldens, capability goldens,
the generator, OfficeCLI disposition data and corpus files are unchanged.

## Gates

```text
python3 -m unittest tools.test_officecli_capability_delta — pass (44 tests)
python3 tools/officecli_capability_delta.py --check — pass
python3 tools/officecli_corpus.py --check — pass
tools/gate/run_gate.sh capability-report-repair 48 52 — pass
  row 48 tooling-unit: 1,800 tests, 0
  row 52 corpus-dispositions: 0
```

Row 42 (`goldens`) passed with the coordinator-approved terminal target and
the prescribed offline, four-job environment:
`CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/gate-trunk/target`
`CARGO_BUILD_JOBS=4 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0`
`CARGO_INCREMENTAL=0 CARGO_NET_OFFLINE=true` and the gate-tool path prefix.
The exact command was `bash tools/gate/check_goldens.sh`; capability/registry
legs (both feature profiles), the litchi-cli corpus leg, and regenerated
report/corpus checks passed. The full 53-row gate remains coordinator-owned.

No native, parity, performance, or feature-completeness claim is made by this
record.
