# 0429 — W3: delete the `litchi-iwa` host (held for the macOS queue)

- Status: prepared on branch `uo2/iwa-w3`, cut from trunk
  `feat/unified-ops-api` at `27b4840684`, trunk `4dbfe29241` (record 0430)
  merged in. **Held**: merge only after the
  macOS queue passes on a Mac (§5). Not pushed.
- Date: 2026-11-03.
- Plan: record [0363](0363-iwa-monolith-exit-plan.md), wave W3, and its
  deletion gate (§6, now §10); ADR [0028](../adr/0028-iwa-monolith-exit.md).
- Precondition (PROGRESS, "Exit ledger complete" and "Final candidate
  qualified", 2026-11-03): the exit ledger has 0 host-only rows (2,117
  migrated, 119 retired), and no workspace crate depends on `litchi-iwa`;
  only its own tests, its parity harness and its examples used it.
- `performance_claim`: none. No leaf or unified route changes; the deletion
  removes a crate nothing links.

## 1. Host-built macOS-queue inputs, kept and listed first

Many queue rows run on packages that the host built at test time ("host-added
Column2d chart", "two host `add_text_box`", "host-seeded title"). The Mac run
tests the packages the record [0430](0430-mac-queue-runbook.md) generator
built from exactly those inputs, and the host draws random UUIDs, so they must
not be regenerated: they are taken from that generator's output, not from a
new host run.

- Inputs: the 798 host-built inputs of the 0430 run (source commit
  `b9c6051c0a`), 181,281,065 bytes, largest 803,212 bytes, in
  `.coord/mac-runbook/out/inputs/`; a copy whose 798 SHA-256 values all match
  is in `/home/zhuhe/code/litchi-worktrees/.coord/iwa-w3/mac-inputs/`. They
  are too large to commit (coordinator decision) and stay outside git.
- Committed: `test-data/iwork/mac-queue/host-built-inputs.tsv`, a byte copy
  of the generator's manifest of them (SHA-256, path, size, fixture
  description, cases; SHA-256
  `138375655ec2f48bac7d500a62db74517923c0a668fdbf74f3f1971da59aa492`), and a
  README with provenance (generator manifest
  `.coord/mac-runbook/out/manifest.json`, SHA-256
  `56628e2d5ab64d32992953d114eb20a17519df08b01196f2b730533dc2f323e6`).
- `tools/mac_queue`: without the harness, `rows.parity_tests` reads the
  1,400 parity-test names frozen at trunk `4dbfe29241` in
  `tools/mac_queue/parity-tests.tsv`, so `generate.py --check` still maps all
  766 rows (764 to packages, 2 grouped, 0 errors); `generate.py` without
  `--reuse-dumps` now stops with an explanation instead of calling the host.

## 2. What is deleted

| Item | Files | Lines |
|---|---:|---:|
| `crates/litchi-iwa` (whole crate, at trunk `4dbfe29241`) | 922 | 415,964 |
| of which `src/` | 564 | 305,167 |
| of which `tests/` (18 integration targets and the parity harness `tests/parity`, 143 files / 82,094 lines, with the record 0430 dump hook) | 160 | 91,727 |
| of which `examples/` | 196 | 15,508 |
| host audits in `tools/check_crate_boundaries.py` (319 → 191 `audit_*` functions; 78,326 → 57,320 lines) | 1 | 21,055 net |
| their tests in `tools/test_check_crate_boundaries.py` (51,309 → 23,986 lines; 541 tests remain, all passing) | 1 | 27,416 net |
| `iwa` profile of `tools/api_glossary.json` | 1 | 1,005 |
| host route checks in `crates/litchi-iwa-protos/build.rs` | 1 | 259 |
| 12 `litchi-iwa -> *` debt entries and the host package entry in `tools/crate_boundaries.json` | 1 | 90 |

Against trunk `4dbfe29241`, files other than the crate lose 50,252 lines and
gain 2,712 (histogram diff): the two frozen lists (`parity-tests.tsv`, 1,402
lines; `host-built-inputs.tsv`, 799 lines), this record, the ADR amendments
and the reworded provenance comments.

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
- **Provenance comments** now cite the host by commit (`27b4840684`), not by
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
  with the crate, its inputs frozen as in §1.

## 4. Gaps and follow-ups

- The macOS queue (§5) is the only open deletion-gate item (ADR 0028 item 4).
- The leaf crates still carry hidden seams that only the host called (for
  example `__decode_movie_playback_payload`, the `internal-iwork-source`
  features of `litchi-pages`, `litchi-numbers` and `litchi-keynote`, the
  Numbers physical relocation and persisted-sort bridges). They are dead after
  the deletion; removing them changes leaf APIs and is a separate package.
- The three leaf feature matrices predate W1 to W4 in many rows; re-auditing
  them against the exit ledger is a docs follow-up.
- The host-built inputs exist only on the Linux host (`.coord`); losing both
  copies before the Mac run passes would require rerunning the 0430
  generator at a commit that still has the host, and then the Mac run.

## 5. How to merge once the Mac run passes

1. Run the record 0430 runbook (`tools/mac_queue/run_on_mac.sh`) on a Mac
   with the packages already generated in `.coord/mac-runbook/out` (do not
   regenerate them; after this branch merges they cannot be).
2. Fold the results into `0363-mac-queue.md` with
   `tools/mac_queue/apply_results.py` (fills `Result` and `Artifact SHA-256`),
   on trunk or on a commit on top of this branch.
3. If a row fails, fix the leaf route on a separate branch from trunk, merge
   it, and rebase this branch; the deletion itself needs no change.
4. When every row passes: rebase `uo2/iwa-w3` onto the trunk head (the only
   expected conflicts are queue rows or crosswalk rows added meanwhile; a new
   trunk dependent of `litchi-iwa` would be a regression to remove), rerun the
   full gate (`tools/gate/run_gate.sh`, 53 rows) on the rebased head, then
   merge. Keep `.coord/mac-runbook/out/` and `.coord/iwa-w3/mac-inputs/`
   until every queue row has passed.

## 6. Commits and verification

Commits on `uo2/iwa-w3` (not pushed):

1. `refactor(iwork)!: delete the litchi-iwa migration host` — the crate, the
   boundary policy and checker, the protos build script and its guard test,
   the ledger, glossary and gate tools.
2. `docs(iwork): retire litchi-iwa references` — comments, feature
   matrices, glossary, audits, ADR 0024/0028, record 0363 §10.
3. merge of trunk `4dbfe29241` (record 0430); the harness's dump hook is
   deleted with the crate.
4. `tools(mac-queue): run without litchi-iwa` — parity-test snapshot, the
   generator's host gate, the host-built input manifest.
5. this record and the regenerated `results/0289/bench-matrix.md`.

Before the deletion, on `27b4840684`: `cargo test -p litchi-iwa --all-features
--test parity` passed (1,190 passed, 210 ignored cost probes). On the
branch (verification head `0d58adb5f3`, commit 4; commit 5 adds only this
record and the regenerated bench matrix):

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-iwa-{protos,archive,style,table,text-wire} -p litchi-keynote -p litchi-numbers-wire -p litchi-numbers -p litchi-pages -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy --workspace --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings` | 0 (no format feature touched) |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p` (the nine changed crates) | 0 |
| `tools/gate/chunked_cargo_tests.py` over the nine changed crates `--all-features` | 0 (60 commands; run before the trunk merge, which does not touch them; protos `--lib` 1,063 passed after the guard fix) |
| `tools/gate/chunked_cargo_tests.py -p litchi-ops --all-features` | 0 (849 s; capabilities and registry goldens unchanged, nothing re-blessed) |
| `tools/gate/chunked_cargo_tests.py -p litchi --all-features` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 (69 packages, 303 declarations, 0 debt items) |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_iwa_exit_ledger tools.test_mac_queue tools.test_check_iwork_leaf_public_api tools.test_check_api_glossary tools.test_api_debt_inventory tools.test_validate_api_disposition tools.test_unified_bench_matrix` | 0 (827 tests, with this record present) |
| `python3 tools/officecli_capability_delta.py --check`, `tools/officecli_corpus.py --check`, `tools/unified_bench_matrix.py --check`, `tools/iwa_exit_ledger.py --check`, `tools/mac_queue/generate.py --check` | 0 (766 rows: 764 mapped, 2 grouped, 0 errors) |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir $CARGO_TARGET_DIR` | 0 (72 profiles) |
| `python3 tools/check_feature_matrix.py` | 0 |

The full 53-row gate (`tools/gate/run_gate.sh`) is to be run on the rebased
head before the merge (§5).
