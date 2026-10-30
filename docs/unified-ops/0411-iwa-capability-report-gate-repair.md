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

Row 42 (`goldens`) was not run: it requires Cargo workspace execution and the
coordinator’s large terminal target cache. No duplicate cache was created.
The full 53-row gate remains coordinator-owned.

No native, parity, performance, or feature-completeness claim is made by this
record.
