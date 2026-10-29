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

## Performance qualification on frozen source

The affected existing routes are the three formula-authoring ABBA probes in
`crates/litchi-iwa/tests/parity/numbers_formula_author_cases.rs`, especially
`numbers_table_formula_author_graph_cost_abba`, which already exercises
resolved formula-owner state. A new ignored full-facade probe,
`numbers_table_formula_owner_admission_cost_abba`, prepares a source-built
two-table workbook outside the capture, invokes the independent legacy `Draft`
formula writer on both owners, duplicates the sheet, edits the cloned formula
with its cached value, and serializes. The leaf side uses selector-first
`Package` table-cell and sheet lifecycle transactions. The full logical route
is measured in host/leaf/leaf/host order; the setup is not counted as an
isolated edit or as a whole-file benchmark.

Frozen source: `2e1a7a3f6ad31e6546bfe7b0557246af949e7eb5` (the docs-only
current WP tip after merge `f6c8ec8d05c17c857717e3aabbd458c775f6679e`).
Release binary (kept outside the repository):
`/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1f-owner-admission/perf/rerun-2e1a7a3f6a/archive/parity-96c80ba0598fc9992029b4c8fbf395b9a45c9beb6f40abf35d6f234dbc1dcd0c`,
210,625,192 bytes, mode 0555, SHA-256
`96c80ba0598fc9992029b4c8fbf395b9a45c9beb6f40abf35d6f234dbc1dcd0c`.

| Route | Host Ir | Leaf Ir | Delta |
|---|---:|---:|---:|
| Basic no-cache authoring | 189,879,802 | 83,999,055 | **−55.762%** |
| Native-graph cached authoring | 159,284,086 | 50,061,400 | **−68.571%** |
| Basic explicit cached authoring | 189,859,438 | 84,013,752 | **−55.749%** |
| Owner-admission workflow | 500,405,208 | 51,610,047 | **−89.686%** |

All four routes are faster on the leaf side and meet the prescribed no-more-
than-5%-slower bar. Captures and exact-test exit logs are external under
`perf/rerun-2e1a7a3f6a/{callgrind,logs}`; their sizes and SHA-256 values are in
`perf/rerun-2e1a7a3f6a/receipt-2e1a7a3f6a.txt`. The earlier f6c8 receipt remains
historical only: it had a graph-leaf hash transcription typo and no persisted
per-capture logs, so it is not the current qualification receipt.

Regenerate the complete set with:

```sh
CARGO_BUILD_JOBS=4 CARGO_INCREMENTAL=0 \
  CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1f-owner-admission/perf/rerun-2e1a7a3f6a/target \
  CARGO_PROFILE_RELEASE_DEBUG=1 \
  cargo test --release --locked --offline -p litchi-iwa --test parity --no-run
taskset -c 3 valgrind --tool=callgrind \
  --callgrind-out-file=/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1f-owner-admission/perf/callgrind/basic.host \
  --toggle-collect='*numbers_formula_author_cases::host_route*' \
  /home/zhuhe/code/litchi-worktrees/.coord/iwa-w1f-owner-admission/perf/rerun-2e1a7a3f6a/archive/parity-96c80ba0598fc9992029b4c8fbf395b9a45c9beb6f40abf35d6f234dbc1dcd0c \
  --ignored --exact numbers_formula_author_cases::numbers_table_formula_author_cost_abba --test-threads=1
# Repeat with leaf_cost_route, the graph/cached exact test names, and the
# owner route toggles host_formula_owner_admission_route and
# leaf_formula_owner_admission_route.
```

## Review follow-up

Execution now rechecks the actual owner slices on every formula-write plan and
resolved-read-registry plan before any output or retained-vector allocation.
The regression controls cover a lowered owner ceiling after planning (with no
dependency visitor callback) and a valid zero-owner scalar plan; the existing
read-registry lowered-limit case now reports the typed owner dimension. This
prevents a plan admitted under a broader profile from bypassing a narrower
execution profile.

The merge wire reader keeps an owner-entry refusal in a scoped
`OwnerEntriesLimit` seam on `MergeReadError`; the Numbers table-merge facade
maps that exact dimension to its own `LimitKind::OwnerEntries`. It is not
coerced into protobuf fields. The other adapter mappings to existing reference
dimensions are documented at their mapping sites because those adapters count
resolved owner entries as retained formula references.

The formula-cache evaluator now pre-admits the actual table-owner registry
cardinality against `CacheLimits::wire_references`, charges it once, and passes
that exact cardinality to every resolved-reader decode. It no longer derives an
owner ceiling from `formula_work` or `graph_work`; those remain AST/evaluation
and graph-operation ceilings respectively.

## Verification on this WP

All commands used offline with `CARGO_BUILD_JOBS=2`,
`CARGO_INCREMENTAL=0`, and development debug info disabled where applicable:

| Command | Result |
|---|---|
| `cargo test --offline -p litchi-iwa-protos numbers_formula_codec::tests::resolved_formula_owner_registry_has_independent_ceiling -- --nocapture` | pass |
| `cargo test --offline -p litchi-iwa-protos numbers_formula_codec::tests --lib -- --nocapture --test-threads=1` | 39 passed |
| `cargo test --offline -p litchi-numbers-wire formula_cache --lib -- --nocapture --test-threads=1` | 49 passed |
| `cargo test --offline -p litchi-numbers-wire --test table_merges -- --nocapture --test-threads=1` | 16 passed |
| `cargo test --offline -p litchi-numbers --test source_built_sheet_duplicate_tables -- --nocapture` | 4 passed |
| `cargo test --offline -p litchi-numbers --test formula_authoring -- --nocapture --test-threads=1` | 16 passed |
| `cargo test --offline -p litchi-iwa --test parity numbers_formula_author_cases -- --nocapture --test-threads=1` | 5 passed, 4 intentional cost probes ignored |
| `cargo test --offline -p litchi-iwa --test parity -- --ignored --exact numbers_formula_author_cases::numbers_table_formula_owner_admission_cost_abba --nocapture --test-threads=1` | pass; execution smoke only |
| `cargo test --offline -p litchi-pages --test body_table_merges -- --nocapture --test-threads=1` | 11 passed |
| `cargo test --offline -p litchi-keynote --test slide_table_merges -- --nocapture --test-threads=1` | 4 passed |
| `cargo fmt --all -- --check` | pass |
| `cargo test --offline -p litchi-iwa --test parity --no-run` | pass |
| `python3 -m unittest tools.test_no_3rdparty_references` | 4 passed |

The coordinator still owns the full integrated gate, final performance
qualification, native queue execution, and post-merge fuzz locks.
