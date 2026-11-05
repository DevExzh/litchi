# 0429 — W3: delete the `litchi-iwa` host

- Status: merged to `feat/unified-ops-api` at `3001da641b` and to
  `feat/office-format-completeness` at `967e6ccce8`, on the owner's instruction,
  before the 53-row gate finished; see PROGRESS for its result. Not pushed.
- Date: 2026-11-05 (first prepared 2026-11-03).
- Plan: record [0363](0363-iwa-monolith-exit-plan.md), wave W3, and its
  deletion gate (§6, now §10); ADR [0028](../adr/0028-iwa-monolith-exit.md).
- Evidence: all 767 rows of the macOS queue have Mac evidence, 0 failed and 0
  pending (record [0434](0434-mac-run-3-results.md) §2). Three of those rows
  name an open part (0434 §10). The Pages identity row is closed. **Owner
  decision (2026-11-05):** the by-reference part of the two pivot rows is
  discarded, with no run 4 and no leaf test of that shape. The queue's cells are
  unchanged.
- Precondition: the exit ledger has 0 host-only rows (2,117 migrated, 119
  retired), and no workspace crate depends on `litchi-iwa`; only its own tests,
  its parity harness and its examples used it.
- `performance_claim`: none. No leaf or unified route changes; the deletion
  removes a crate nothing links.

## 1. Host-built macOS-queue inputs, kept and listed first

Many queue rows run on packages that the host built at test time ("host-added
Column2d chart", "two host `add_text_box`", "host-seeded title"). The host draws
random UUIDs, so a fresh run gives different bytes, and passes on such a package
rest on the bytes that went to the Mac. The set is therefore taken from the
generator output record 0434 folded against, not from a new host run.

- Inputs: 797 host-built inputs, 181,143,941 bytes, largest 803,216 bytes, in
  `.coord/iwa-w6/queue-out/inputs/`. They are too large to commit (coordinator
  decision) and stay outside git.
- Committed: `test-data/iwork/mac-queue/host-built-inputs.tsv`, a byte copy of
  that run's list (SHA-256, path, size, fixture description, cases; SHA-256
  `ee630a4f2a4897db2d273426c0452ad461a591a7c50adff689194ad0c1d68a7f`), and a
  README with provenance (manifest `.coord/iwa-w6/queue-out/manifest.json`,
  SHA-256 `6d59fef38bb0900568c337743bcde98aad0169d4bbf55e8a6a3370f7b98d045a`).
- Verified by SHA-256 against that manifest
  (`.coord/iwa-w3/scripts/verify_frozen.py`, exit 0, 0 errors): all 797 inputs
  and all 1,160 packages hash to their manifest entries, with no extra file; of
  the packages' 1,238 sources, 895 name a listed host-built input, 341 a tracked
  `test-data` package and 2 no input; every listed input is used; the queue and
  the manifest hold the same 767 row identifiers.
- Kept outside the repository: `.coord/iwa-w6/`, `.coord/mac-runbook/out*` and
  `.coord/iwa-w3/mac-inputs/` (the 798 inputs of the record 0430 run).
- **The generator cannot rebuild this manifest.** `queue-out/raw/` holds only
  `example/` (directory last changed at 12:40 on 2026-11-04, two minutes after
  the manifest), and that run's parity dumps are nowhere under `.coord`. I do not
  know who removed them. `inputs/`, `packages/` and the manifest remain, and they
  are what the evidence cites. `--reuse-dumps` is therefore proved only by the
  unit tests on a fake repository; a real run on a copy stops, naming 741 rows
  with no dump.
- `tools/mac_queue` without the host: `generate.py --check` maps all 767 rows
  (765 to packages, 2 grouped, 0 errors) from the 1,402 parity-test names frozen
  at `4d8a3e161f` in `tools/mac_queue/parity-tests.tsv`. A run without
  `--reuse-dumps` stops with an explanation, and `--out` no longer has a default
  that points at the evidence. Both examples run without the host: of the 28
  files they write, 26 equal `queue-out/raw/example/`, and the other 2 are the
  `pages-builder-identity-*` documents, which call `Identity::generate()`.
- The queue is unchanged: `git diff 4d8a3e161f -- docs/unified-ops/0363-mac-queue.md`
  is empty. W3's earlier note in its preamble was dropped, and `apply_results.py`
  was not run.

## 2. What is deleted

Against trunk `4d8a3e161f`:

| Item | Files | Lines |
|---|---:|---:|
| `crates/litchi-iwa` (whole crate) | 922 | 416,045 |
| of which `src/` | 564 | 305,169 |
| of which `tests/` (18 integration targets and the parity harness `tests/parity`, 143 files / 82,153 lines) | 160 | 91,806 |
| of which `examples/` | 196 | 15,508 |
| host audits in `tools/check_crate_boundaries.py` (319 → 191 `audit_*` functions; 78,326 → 57,320 lines) | 1 | 21,006 net |
| their tests in `tools/test_check_crate_boundaries.py` (51,309 → 23,986 lines) | 1 | 27,323 net |
| `iwa` profile of `tools/api_glossary.json` | 1 | 1,005 |
| host route checks in `crates/litchi-iwa-protos/build.rs` | 1 | 249 net |
| 12 `litchi-iwa -> *` debt entries and the host package entry in `tools/crate_boundaries.json` | 1 | 88 net |

Files other than the crate: 65 changed, 22343 lines added and 68,736 removed. The
additions are mostly the checker's test file re-laid out (17,671 lines by this
diff), the two frozen lists, this record and the new leaf tests.

## 3. What changes besides the deletion

- **Boundary policy and checker.** `migration_hosts` and `migration_debt` are
  empty; the checker reports 69 workspace packages, 303 internal dependency
  declarations and 0 debt items ("ordered migration debt: none"). Every audit
  that inspected host paths is deleted, with the constants and helpers only it
  used. Audits of shared crates that also looked at host readers keep their
  shared checks and lose the host part (chart data/metadata, formula render
  and envelope, table data list, cell value, drawable-container codec,
  image adjustments, media playback, table-cell borders, atomic publication,
  Keynote media creation, movie-geometry completion, slide-media core
  boundary, physical sort, neutral table model, chart-grid creation). Leaf
  forbidden-import patterns no longer list the bare `litchi_iwa` crate. The
  migration-host mechanism of the policy stays; its tests use a synthetic
  host (`litchi-legacy-host`).
- **Gate.** No `tools/gate/run_gate.sh` row named the host package: the
  workspace rows (`clippy-workspace`, `tests-workspace-excluding-ods-ops`,
  `wsdoc`) linted and tested it through `--workspace` and simply no longer
  see it. Row `tooling-unit` keeps `tools/iwa_exit_ledger.py --check` and
  `tools/mac_queue/generate.py --check`, which both run without the crate. The fuzz workspaces and their locks never named
  the host (only `crates/litchi/fuzz/README.md` did, reworded).
- **Exit ledger** (`tools/iwa_exit_ledger.py`). The ledger stays as the
  historical record. `--check` no longer scans host sources or parity-test
  declarations; it requires no `host-only` row, a resolvable destination, a
  parity test name and tier A/B for every migrated row, and existing native
  evidence or a queue row. `--update` is removed.
- **Other tools.** `non_iwork_gate.py` (18 iWork, 21 excluded, 69 packages),
  `check_iwork_leaf_public_api.py`, `api_debt_inventory.py` drop the host.
  `tools/check_feature_matrix.py` and `tools/unified_bench_matrix.py` never
  named it; the bench matrix's dated "leaf-versus-host" evidence reasons are
  history and stay.
- **Build script.** `litchi-iwa-protos/build.rs` checked Pages message-ID
  routes and the type registry in host sources and failed on a partial
  sibling checkout; the 12 host route declarations, 23 host production
  markers and the registry check are removed, the leaf `litchi-pages` routes
  stay checked.
- **Provenance comments** now cite the host by commit (`4d8a3e161f`), not by
  path: `litchi-numbers-wire` `table_create/fresh.rs`, `litchi-numbers`
  `create.rs`, `litchi-pages` `builder.rs`, the Keynote test
  `slide_drawable_comments_host_parity.rs`, and the other comments that
  called the host current (shared `litchi-iwa-*` engines, leaf bridges,
  tests, the fuzz README, the Keynote blank-asset regeneration recipe and the
  synthetic Numbers fixture README). The author string `"litchi-iwa"` that
  leaf comment writers store in documents is data and is unchanged.
- **Docs.** The three leaf feature matrices replace their legacy-host delta
  sections with a "Retired migration host" note and drop host links;
  `docs/FEATURE_MATRIX.md`, `docs/API_GLOSSARY.md` and the two dated iWork
  audits (now marked historical snapshots) are updated; ADR 0024 and ADR 0028
  have 2026-11-03 amendments; record 0363 has §10.
- **Leaf tests using the host as an oracle.** None link the host (no
  workspace crate depended on it); the parity harness that did is deleted
  with the crate, its inputs frozen as in §1 and the behaviours trunk added to
  it dispositioned in §4.

## 4. Coverage disposition

Before the deletion, at `4d8a3e161f`, the whole harness passed:
`cargo test -p litchi-iwa --all-features --test parity`, exit 0, 1,192 passed and
210 ignored cost probes, plus one temporary probe, not committed, that wrote the
expected values below (`.coord/iwa-w3/probe-out/`). Host and leaf gave the same
values in every case the probe compared. The leaf tests below pin those values.

**Owner decision (2026-11-05):** remove the crate without regard to historical
debt. The parity cases outside this table are dropped, not ported, and the
by-reference pivot tests are discarded.

The inline synthetic fixture is built in the test, and its bytes equal the
harness's inline fixture (SHA-256 `74ba2c6424a45a51…`, checked once with a
temporary dump).

| Host test or fix (record) | Behaviour | Now |
|---|---|---|
| `series_non_style`, `series_style` count (0431) | `count` is one past the highest index; a count unequal to the entries reads | Kept: `litchi-iwa-style` `sparse_count_is_one_past_the_highest_series_index`; `litchi-iwa-protos` `series_non_styles_match_the_generated_encoder`; `litchi-keynote/tests/native_pie_third_wedge.rs` on Apple's own file |
| `reference_line/archive` count (0431) | Same rule on a trim of the styles map | **New**: `litchi-iwa-style` `trimmed_style_count_is_one_past_the_highest_kept_index`. It had no leaf test. Still no Mac proof (0434 §12) |
| `image_caption`, `sheet_charts/caption` (0431) | The caption's paragraph style is listed across components | Kept: `litchi-numbers/tests/chart_caption_native_structure.rs`, `chart_caption_lists_its_paragraph_style_like_a_native_file`, and a check on native files |
| `numbers_pivot_categories_read_parity` | Categories of the plain workbook and of the inline tree | **New**: `litchi-numbers/tests/pivot_synthetic_group_bys.rs`, `plain_workbook_has_no_categories`, `reads_the_inline_tree` (11 exact rows). By-reference read: **discarded** by the owner |
| `numbers_pivot_formula_author_parity`, `_without_cache_`, `_inline_only_` | Formula archive and edges equal to the host's; inverse | **New**: `authors_a_category_of_the_inline_tree`, `authors_a_category_without_a_cache`. Whether a cache is absent is pinned only as "the output differs from the one with a cache" |
| `numbers_pivot_formula_author_referenced_tree_parity` | Formula on a by-reference tree | **Discarded** by the owner. `pivot_group_tree_bounds.rs` still proves that such a tree, cycle included, reads without overflowing the stack |
| `numbers_pivot_formula_missing_category_is_refused_like_host`, `_changed_category_` | Typed refusals | **New**: `a_category_absent_from_the_edited_package_is_refused`, `a_category_with_another_function_in_the_edited_package_is_refused`, on the inline fixture. The host's refusal half is retired: no host |
| `numbers_native_pivot_formula_host_and_leaf_both_write` | Leaf writes on the source table and a pivot body cell | Kept: `litchi-numbers/tests/native_pivot.rs`, `..._on_source_table_keeps_each_function`, `..._matches_what_numbers_wrote`. Retired: "the host reads 13", a host defect |
| `numbers_native_conditional_highlight_volatile_host_fails_leaf_registers` | Leaf half: first rule on the two Numbers-saved fixtures | **New**: `native_volatile_highlight.rs`, `numbers_saved_tables_register_their_first_volatile_rule`; the structure is in `first_volatile_rule_on_a_numbers_saved_table_builds_what_numbers_builds`. Retired: the host's failure |
| `numbers_conditional_highlight_volatile_parity`, `_lifecycle_parity` | Family registration and lifecycle on a built document | **New**: `built_document_volatile_lifecycle`, on the leaf `Builder`'s default table. Retired: comparison of known fields with the host, and the 8 x 6 and 12 x 40 host documents. Mac evidence stands on the frozen inputs |
| `leaf_allocator_avoids_the_host_owner_map_collision` (`exit_parity`, volatile fixture added by 0432) | Allocation above the map; the colliding identifier refused | **New**: `litchi-numbers/tests/owner_allocator_native.rs`. It scans every workbook with no clonable-family filter, so it pins 5 colliding workbooks where the host test listed 3. Byte equality with the host's engine rewrite on the others is retired |
| The `*_cost_abba` probes | Leaf-versus-host cost | Retired: they measured the host. Results stay in their records |
| `control_kit`, `variant-no-uuid-map`, `next-run`, `--evidence` | Tools | Kept, with their tests (`tools.test_mac_control_kit`, `tools.test_mac_queue`) |

**The by-reference branch of the leaf reader stays, and nothing pins its values.**
No writer makes that shape and no file Numbers wrote reaches the branch (0434 §10).

## 5. Gaps and follow-ups

- **Parity cases outside §4 are dropped by owner decision**, with no port planned.
  They were the Linux regression check for many routes: a word scan finds that for
  688 of 2,117 migrated ledger rows (194 distinct destinations), the destination's
  method name appears in no test of the leaves, the unified crates or the shared
  crates' test modules (`.coord/iwa-w3/scripts/ledger_gap.py`; coarse in both
  directions). A regression on those routes would not show on Linux.
- The other open items of 0434 §10 and §12.
- The leaf crates still carry hidden seams that only the host called (for
  example `__decode_movie_playback_payload`, the `internal-iwork-source`
  features of `litchi-pages`, `litchi-numbers` and `litchi-keynote`, the
  Numbers physical relocation and persisted-sort bridges). They are dead after
  the deletion; removing them changes leaf APIs and is a separate package.
- The follow-up caption-style change is separate work.
- The three leaf feature matrices predate W1 to W4 in many rows; re-auditing
  them against the exit ledger is a docs follow-up.
- The host-built inputs exist only on the Linux host (`.coord`), and cannot be
  rebuilt: a later Mac rerun of a host-built row depends on that directory.

## 6. Commits and verification

Commits on `uo2/iwa-w3` (not pushed):

1. `refactor(iwork)!: delete the litchi-iwa migration host`, with the boundary
   policy, checker, protos build script, ledger, glossary and gate tools.
2. `docs(iwork): retire litchi-iwa references`.
3. merge of trunk `4dbfe29241` (record 0430).
4. `tools(mac-queue): run without litchi-iwa`.
5. the first version of this record.
6. merge of trunk `4d8a3e161f`. The nine host files trunk changed resolve as
   deleted; the generated bench matrix is regenerated.
7. `test(iwork): pin in the leaves what only the host differential proved`.
8. `tools(mac-queue): freeze the inputs behind the Mac evidence`.
9. `docs(iwork): lift W3's held wording`.
10. this version of the record and the gate results.
11. `test(numbers): discard the by-reference pivot tests`, by owner decision.

Before the deletion, at `4d8a3e161f`: see §4. Verification head `c11ed2e6f1`
(this commit changes only this record). Logs: `.coord/iwa-w3/gates3/`. First-pass
failures: `owner_allocator_native` failed once, because it found 5 colliding
workbooks where I had copied the host test's 3 (§4; the test was corrected, no
leaf changed). `generate.py --reuse-dumps` on a copy of the frozen output exited 1
(§1). Nothing else failed. Goldens were not re-blessed: no capability, registry or
corpus output changed.

| Command | Exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo clippy` nine changed crates, `-p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy --workspace --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features pages, keynote, numbers] --all-targets -- -D warnings` | 0 (all four) |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features` nine changed crates | 0 |
| `tools/gate/chunked_cargo_tests.py` nine changed crates `--all-features`, new tests included | 0 |
| `tools/gate/chunked_cargo_tests.py -p litchi-ops --all-features` | 0 |
| `tools/gate/chunked_cargo_tests.py -p litchi --all-features` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `mac_queue_file_routes`, `mac_queue_native_routes` examples | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 (69 packages, 303 declarations, 0 debt items) |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `python3 -m unittest` the rules' tool modules plus `test_iwa_exit_ledger`, `test_mac_queue`, `test_mac_control_kit`, `test_gate_runner`, `test_check_iwork_leaf_public_api`, `test_check_api_glossary`, `test_api_debt_inventory`, `test_validate_api_disposition`, `test_unified_bench_matrix` | 0 |
| `officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check`, `iwa_exit_ledger.py --check`, `mac_queue/generate.py --check` | 0 |
| `.coord/iwa-w3/scripts/verify_frozen.py` | 0 |
| `git diff --exit-code 4d8a3e161f -- docs/unified-ops/0363-mac-queue.md` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir $CARGO_TARGET_DIR` | 0 (72 profiles) |
| `python3 tools/check_feature_matrix.py` | 0 |

Commit 11 came after this table. It removes test code and changes docs and one
tool comment, so only these were rerun on it, all exit 0: `cargo fmt --all --check`,
`cargo test` and `cargo clippy -D warnings` of `pivot_synthetic_group_bys`,
`check_crate_boundaries.py`, `unified_bench_matrix.py --check`,
`iwa_exit_ledger.py --check`, `mac_queue/generate.py --check`, and the unit tests
`test_mac_queue`, `test_mac_control_kit`, `test_no_3rdparty_references`,
`test_unified_bench_matrix`.

The full 53-row gate (`tools/gate/run_gate.sh`) was not run here; it is the
coordinator's.
