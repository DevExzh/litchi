# 0893 producer full-cell guard qualification

This is harness correctness evidence, not a performance experiment. The
producer full-cell runner now forces every worksheet's tolerance report after
its timer, allocation region and process endpoint snapshots, then checks the
workbook aggregate and source freshness. The serialized guard contract is
`post-endpoint-per-worksheet-tolerance-v1`.

Measured guard vectors follow the report's elapsed sample order; warmup vectors
remain in execution order. The successful-guard iteration total includes both.
The fixed archive, cell oracle and default case selection remain unchanged.
The raw selected scanner's refusal tests are separate test-only observations;
`runtime_route_observed` remains false in benchmark reports.

## Reproduction

The exact Rust source copies and executed validation helper are retained per
attempt. `validate.py` uses Rust 1.95, a dedicated marked target, low-debug
settings, locked/offline Cargo, sequential gates and a create-only attempt.
Run from the repository root with a fresh attempt name:

```sh
python3 -B docs/performance/results/change-0893/validate.py --run --attempt quality-local
python3 -B docs/performance/results/change-0893/verify_guard.py \
  --log docs/performance/results/change-0893/attempts/quality-local/logs/producer_full_cell.log \
  --output-root docs/performance/results/change-0893/qualification-local
```

The eight gates are scoped formatting, the full-cell tests, selected-scanner
controls, forced-report tests, the existing producer tests, warning-denied
library Clippy, warning-denied library rustdoc, and crate boundaries. Every test
gate requires a positive passing count. The runner stops on failure and keeps
its logs and snapshots. Retained failed attempts must not be overwritten.

`verify_guard.py` checks the one serialized 1-warmup/3-sample test result,
including fixed corpus/oracle values, sample alignment, every guard vector,
the guard version and post-guard freshness. It rejects deliberately altered
reports and the historical aggregate-only 0892 report. Debug elapsed values in
this report are incidental test output, not a new baseline.

## Scope and cleanup

Forced materialization is outside the current timed region and its allocation
and process endpoints. It still affects whole-process RSS/profilers and later
iterations' allocator/high-water state. Any future paired experiment must
rebuild both lanes with the same guard contract and schedule. Historical
0885/0891/0892 captures remain unchanged and do not satisfy this new guard.

`exploratory-agent-report.json` records an agent's earlier exploratory dev
commands, including a failed dimension-error assumption. Those runs lack
sealed logs/source snapshots and are not qualification evidence. The final
malformed-prefix test uses invalid root XML and requires a typed error.

The dedicated `target/litchi-0893-quality` target is disposable only after its
exact ownership marker and terminal gate receipt have been checked. Shared
build targets and unrelated worktrees must remain untouched. Source copies,
commands, logs, reports and review receipts are retained as evidence.

Formal attempt `quality-v1` stopped at dependency compilation with tmpfs user
quota errors, before any test ran. Its source snapshots and log remain intact.
The marked `/tmp/litchi-0893-quality` target was removed; `quality-v2` uses
workspace storage and a private temporary directory inside its owned target.
The Rust source is identical across those attempts.

`quality-v2` passed all eight gates. The final `quality-v3` attempt extends the
existing dirty-store test to all four worksheet positions; no production or
guard implementation changes were made between these attempts.

Final `quality-v3` passed all eight gates. `qualification-v3` and
`qualification-after-cleanup` pass the literal guard/oracle checks, twelve
negative controls and historical-report rejection. `cleanup.json` records the
owned target removal. The source snapshots and logs remain available after
cleanup; no shared target or unrelated worktree was removed.
