# 0390 — W1-F: independent formula-owner admission

- Status: implementation slice on `uo2/iwa-formula-owner-admission`, based on
  `1076caabe7` (the accepted source-built duplication coverage checkpoint).
- Date: 2026-10-28.
- Scope: keep authored formula AST admission independent from the resolved
  formula-owner registry used by the Numbers formula author.
- Native evidence: none. The existing Numbers formula-authoring routes remain
  queued in `docs/unified-ops/0363-mac-queue.md`; this repair does not claim
  native open/save/close/reopen acceptance.

## Failure and bounded repair

The source-built multi-table duplication regression could duplicate a sheet,
but editing the cloned formula-bearing `Alpha` table refused before
publication with `FormulaWork { observed: 3, maximum: 1 }`. The authored
formula is a one-node numeric AST; the observed three is the resolved
owner-registry entry count, not AST work. The independent repro was run
against frozen coordinator source `9932c7d069` from
`/home/zhuhe/code/litchi-worktrees/.coord/w2d-formula-dup-repro-9932`:

- source `crates/litchi-numbers/tests/formula_dup_repro.rs`, SHA-256
  `25a45901ea564cc7bae6276867c5b5ec9ae308e181bbdd339b1367ce10df2602`;
- log `/home/zhuhe/code/litchi-worktrees/.coord/w2d-formula-dup-repro-9932.log`,
  18,472 bytes, SHA-256
  `1640144d7ff691b65c924ea54ee552255c83173bd412a5d9ef2a4d9387368c3b`.

`litchi-iwa-protos::numbers_formula_codec::DecodeOptions` now retains the
existing `max_nodes` AST ceiling and exposes a separate
`with_max_owner_entries` ceiling. Formula archive planning and resolved-owner
read planning use the owner ceiling only for owner-registry entries; AST
planning still rejects a three-node expression when `max_nodes == 1`. Owner
exhaustion is a distinct typed codec `DecodeLimit::Owners` and the Numbers
author maps it to the existing `LimitKind::References` budget. The author
handoff supplies the exact resolved `wire_owners.len()` after the surrounding
owner work/storage envelope has been admitted; it does not widen the AST
quota with `max(...)` or introduce a compatibility fallback.

## Correctness coverage

The shared codec unit test
`resolved_formula_owner_registry_has_independent_ceiling` proves all three
directions: one AST node with three owners succeeds under an owner ceiling of
three, three AST nodes still fail under an AST ceiling of one, and one AST node
fails with the typed owner limit when the owner ceiling is two.

The focused source-built target
`crates/litchi-numbers/tests/source_built_sheet_duplicate_tables.rs` now has
the formula-bearing multi-table case alongside the unchanged host regression.
It checks formula text and numeric caches for both source and cloned tables,
keeps the unrelated `Beta` table independent, applies the exact inverse to
restore source bytes, edits only the clone, and reopens the candidate to check
formula/cache readback. The existing invalid-selector transaction test and
same-name locality case remain in the target. The host target remains the
legacy provenance oracle; it was not weakened or replaced with a numeric
literal case.

No native evidence, fixture, ledger migration, or queue disposition was
invented by this repair. The existing host-vs-leaf formula-author parity rows
remain Tier B/queued until the coordinator obtains native evidence.

## Performance route prepared, not qualified

The affected existing routes are the three formula-authoring ABBA probes in
`crates/litchi-iwa/tests/parity/numbers_formula_author_cases.rs`, especially
`numbers_table_formula_author_graph_cost_abba`, which already exercises
resolved formula-owner state. A new ignored full-facade probe,
`numbers_table_formula_owner_admission_cost_abba`, prepares a source-built
two-table workbook, invokes the independent legacy `Draft` formula writer on
both owners, duplicates the sheet, and edits the cloned formula with its
cached value. The leaf side uses selector-first `Package` table-cell and sheet
lifecycle transactions. A normal test run of the probe passes in
host/leaf/leaf/host order; no wall-clock or callgrind result is claimed here. The
coordinator must run the prescribed release-mode pinned-CPU measurement on the
final integrated commit and qualify this full facade route separately from the
codec unit cost.

## Verification on this WP

All commands used offline with `CARGO_BUILD_JOBS=2`,
`CARGO_INCREMENTAL=0`, and development debug info disabled where applicable:

| Command | Result |
|---|---|
| `cargo test --offline -p litchi-iwa-protos numbers_formula_codec::tests::resolved_formula_owner_registry_has_independent_ceiling -- --nocapture` | pass |
| `cargo test --offline -p litchi-numbers --test source_built_sheet_duplicate_tables -- --nocapture` | 4 passed |
| `cargo test --offline -p litchi-iwa --test parity numbers_formula_author_cases -- --nocapture --test-threads=1` | 5 passed, 4 intentional cost probes ignored |
| `cargo test --offline -p litchi-iwa --test parity -- --ignored --exact numbers_formula_author_cases::numbers_table_formula_owner_admission_cost_abba --nocapture --test-threads=1` | pass; execution smoke only |
| `cargo fmt --all -- --check` | pass |
| `cargo test --offline -p litchi-iwa --test parity --no-run` | pass |
| `python3 -m unittest tools.test_no_3rdparty_references` | 4 passed |

The coordinator still owns the full integrated gate, final performance
qualification, native queue execution, and post-merge fuzz locks.
