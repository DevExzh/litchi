# 0365: Gate infrastructure (rows 41, 48–51 and the runner)

Branch `uo2/fin-gate-infra`, cut from trunk `55b5616c7d`. This record covers the
gate's own infrastructure: the runner, the fuzz locks, the tooling-unit
environment, crate boundaries, the non-iWork gate and the bench matrix. It
does not claim a full 53-row pass. The coordinator runs the whole gate.

## Canonical runner: `tools/gate/run_gate.sh`

Before this change there were two runners, both outside tracked tooling:

- `.coord/gate.sh` is short, but it has 52 rows. It has no `iwork-leaf-api`
  row and no RTF fuzz check, it depends on untracked `.coord/scripts/*`, and
  it inherits whatever PATH its caller has.
- `docs/unified-ops/results/0362ul29/tooling/run_safe_gate.py` defines the
  53-row table and the row numbering the records use. It is a 1,200-line
  qualification driver: it requires a clean source, uses one job, pins
  per-file hashes and tracks process ownership. That is the Codex
  qualification process, which is not an owner requirement.

The tracked runner `tools/gate/run_gate.sh` keeps the 53-row table and its
indices from the 0362ul29 runner, and the simple shell structure of
`.coord/gate.sh`:

- `GATE_DIR` selects the checkout. The default is the checkout that holds the
  script.
- `GATE_OUT` holds `summary.txt` (one `NN name=status` line per row, then
  `DONE failed=N`), one `NN-name.out` file per row, and `environment.txt`
  (commit, PATH, cargo and rustc versions, lock SHA-256).
- Optional arguments select rows by index or name. `--list` prints the table.
- The runner sets HOME (from passwd when it is unset), `CARGO_HOME`,
  `RUSTUP_HOME` and PATH itself:
  `tools/gate/bin:$CARGO_HOME/bin:/usr/local/bin:/usr/bin:/bin`. It verifies
  `rustc` and `cargo` before any row runs. A systemd scope with an empty
  environment therefore behaves like a login shell. Rows 41–52 were run that
  way: `systemd-run --user --scope … env -i …`.
- `tools/gate/bin/cargo` is the OOM-isolation shim, copied from
  `.coord/bin/cargo`. Each cargo command runs in its own user scope with
  `MemoryMax=${LITCHI_MEM_MAX:-26G}` and no swap. The real cargo comes from
  `$CARGO_HOME/bin`, never from PATH. If the shell has no `XDG_RUNTIME_DIR`,
  the shim derives it. If no user manager is reachable, the shim runs
  uncapped and prints a warning.
- `chunked_cargo_tests.py`, `target_headroom.py` and `check_goldens.sh` move
  from `.coord/scripts` into `tools/gate/` without behaviour changes.
- `tools/test_gate_runner.py` checks the following and joins the
  tooling-unit row:
  - the row table has 53 unique rows at their stable indices;
  - the runner starts with an empty environment;
  - every fuzz manifest has a tracked lock;
  - the shim's exec and missing-cargo paths.

Positional arguments to the new runner select rows. The extra clippy `-p`
arguments that `.coord/gate.sh` took positionally move to `GATE_EXTRA`. A
proposed thin `.coord/gate.sh` replacement is at
`.coord/fin-gate-infra/gate.sh.proposed`. It execs the tracked runner with
`GATE_OUT=.coord/gates/<name>` and the trunk lock as `GATE_LOCK_SEED`. The
coordinator owns `.coord/gate.sh` and should swap it in.
`run_safe_gate.py` stays in place as historical evidence, and its two unit
tests stay in the tooling row.

### Root `Cargo.lock`

The root lock is not tracked: `.gitignore` has `*.lock`. ADR 0028 states that
"this library workspace excludes `Cargo.lock` from version control and the
gate must work in a clean checkout". Tracking it would be a policy change,
which is left to the owner.

The runner now handles a checkout without a lock:

1. It copies `GATE_LOCK_SEED` when that is set. Otherwise it runs
   `cargo generate-lockfile`.
2. It runs `cargo metadata` to reconcile the lock with the manifests.
3. It records the lock's SHA-256 in `environment.txt`.

The `--locked` legs of `check_feature_matrix.py` (row 39) then work in a
fresh worktree. This worktree had no root lock and was seeded from the trunk
worktree. Row 39 passes.

## Row 41: fuzz locks

- All twenty `crates/*/fuzz` workspaces now have a tracked `Cargo.lock`. A
  `.gitignore` exception `!crates/*/fuzz/Cargo.lock` keeps them visible.
- Seeds:
  - core, iwa-protos, ops, pages and the facade start from their previously
    used ignored locks (from the unified worktree);
  - the other fifteen start from the trunk workspace lock.
- `cargo metadata` (online) reconciled each lock. `cargo fetch --locked`
  filled the registry cache, which now holds `js-sys 0.3.106` and the
  `wasm-bindgen 0.2.129` family. `cargo metadata --offline --locked` passes
  for all twenty.
- Nothing is vendored, and no version constraint changed.
- The row now checks every fuzz workspace with `--locked`: `--bins` for each,
  plus the ops `apply` feature and the two `smoke_batch_payload` runs. The
  old row checked five.
- Widening the row exposed one stale target: `litchi-pages` fuzz
  `pages_body_table_hidden_axes`, 21 errors. The raw-limits refactor added a
  `Limits` argument to `Catalog::from_bytes`, `to_bytes`, `insert_object`,
  `push_message` and `replace_message`. The fix passes `Default::default()`,
  as the crate's own tests do.
- Vendored `zlib-rs` prints dead-code warnings when it is built as a path
  dependency outside the workspace. These are warnings only, and `cargo
  check` does not deny them.

## Row 48: tooling-unit

- The PPT exception identity is already one class on trunk. 0362ul165uf made
  `tools.ppt_lazy_native_qualification` prefer package-qualified imports.
- The DrawingML pin already matches the current manifest:
  `tools/perf-baseline/Cargo.toml` is `f18a17d5…984a`. The DOCX-metadata
  lane is not merging, so `e98ad3e8…` is not applied.
- The runner supplies the missing `rustc` on PATH.
- One more cause of failure: with `CARGO_TARGET_DIR` exported, as
  `.coord/gate.sh` and every session `env.sh` do, the
  `tools.test_native_qualification_runner` tests fail with "manifest cannot
  override the explicit CARGO_TARGET_DIR". The runner's `_environment` copied
  `os.environ`, so it took an inherited variable for a manifest override.

### Decision: keep the native-qualification runner, fix the environment

The coordinator asked to remove `tools/native_qualification_runner.py` and
its tests unless something kept on trunk depends on them. Several things do:

- `tools/native_qualification/` (paired runners, validators and manifests
  for ODF raw codec, the ODF shared triad, the OLE owner, DrawingML raw and
  the legacy native lanes);
- the tooling-row modules `test_current_native_policy_v2`,
  `test_drawingml_raw_qualification(_v2)` and
  `test_legacy_native_qualification`;
- `run_safe_gate.py`.

Removing the runner means retiring that whole subsystem and its tests
together. That is a larger decision than this record should make, so it is
left to the coordinator or the owner. Instead:

- `_environment` now drops the inherited `CARGO_TARGET_DIR` along with
  `CARGO_HOME` and `CARGO_CONFIG`, then sets its own.
- A manifest `environment` entry that names a different target is still
  rejected.
- A new test covers both cases.
- The tooling row no longer needs `env -u CARGO_TARGET_DIR`.

## Rows 49–51

- 49 (crate boundaries) and 51 (bench matrix) pass with 0362ul165ug's
  registrations on trunk.
- 50 (non-iWork gate) also failed at source level: the `vendor/flate2` path
  dependency was rejected as an unknown workspace package. That was a stale
  rule, not a boundary violation, because the vendored crates are patched
  crates.io sources listed in `workspace.exclude`. This branch had its own
  fix, but trunk record 0366 landed the same fix first. The merge
  `718ba5c9bd` takes trunk's version, and this branch's change is dropped.
  The PATH part of row 50 is fixed by the runner.

## Results

Run `0365-b`: rows 39, 41, 43, 47–52 on this branch, launched with
`systemd-run --user --scope … env -i …`. Logs are under
`.coord/fin-gate-infra/gate-0365-b/`.

RESULTS_TABLE

The first run (`0365-a`, before the pages fix) had rows 47–52 at 0 and row 41
at 101. The 101 came from the pages fuzz target described above.

## Cleanup inventory

The coordinator deleted the obsolete `.coord` target directories with the
owner's approval, so item 6 was dropped. A pre-deletion inventory covering
1,911 entries and 1.10 TB on disk is in
`.coord/fin-gate-infra/cleanup-candidates.tsv.partial`. It is stale now and
kept for reference only.

## Follow-ups

- The coordinator should replace `.coord/gate.sh` with the proposed wrapper
  (or call `tools/gate/run_gate.sh` directly) and run the full 53 rows.
- Owner question: track the root `Cargo.lock`? That would reverse ADR 0028's
  stated policy. Until then the gate seeds or resolves the lock and records
  its SHA-256.
