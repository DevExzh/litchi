# 0334 — Facade feature gates, feature-matrix check, batch-payload fuzz targets

- Status: implemented on `uo2/facade-fuzz`, cut from `feat/unified-ops-api`
  at `3bec513e6a`. Trunk was merged at the merge that follows commit
  `4fbb2a54e9` (records 0329 and 0332).
- Date: 2026-10-17
- Input: the coordinator's DoD audit of 2026-10-17, §1 items 1(a)–(c) and 8,
  §6 (gates), WP6 and the code items of WP14.
- Owned:
  - `crates/litchi/{build.rs, build/, src/lib.rs}` and the facade's tests;
  - `crates/litchi-ops/fuzz/`;
  - `tools/check_feature_matrix.py`;
  - `docs/adr/` for the duplicate glossary ADR.
  - No hunk in the shared dispatch files or in `litchi-ops/Cargo.toml`. No
    format adapter, leaf crate or iWork crate is edited.
- `performance_claim`: none. No runtime path changed: the facade change is
  compile-time gating only, and the fuzz crates are outside the workspace.

## 1. Facade feature gates

**Bug.** `crates/litchi/src/lib.rs` gated `mod unified`, the `File`, `open`,
`create`, `Password` and `ToleranceReport` re-exports, `detection_smart` and
`tolerance` on ten copies of one hand-written list of 14 features. The list
left out `oth`, `odm`, `odg`, `odc`, `odi`, `odb` and `odf-formula`. `oth`
and `odm` enable `odt`, so the audit saw only the five object formats: a
facade built with only one of them had no `litchi::open` and no `File`.

**Fix.** The facade now has a build script.

- `build.rs` reads `Cargo.toml`. It takes every `[features]` entry with a
  plain (not weak) `"litchi-ops/<format>"` item, and emits
  `cargo::rustc-cfg=litchi_unified` when any of those features is enabled.
  It also declares the cfg through `rustc-check-cfg`.
- The parser is `build/format_features.rs`. It has no dependencies, and the
  facade's unit tests include the same file.
- All ten lists became `#[cfg(litchi_unified)]`, and so did the 14-feature
  gate of `tests/unified_open_facade.rs`.
- A new format feature now gates the entry point without anyone editing a
  cfg list.

**Regression check.** Four unit tests in `lib.rs` (`format_feature_tests`):

1. `every_claimed_format_feature_gates_the_unified_entry_point` asserts that
   the derived set is exactly the 21 claimed formats of
   `docs/FEATURE_MATRIX.md`.
2. `every_format_feature_names_a_litchi_ops_feature` asserts that each of
   them is a feature of `litchi-ops`.
3. `parser_reads_multi_line_arrays_and_skips_weak_items_and_comments` runs
   the parser on a synthetic manifest.
4. `the_unified_entry_point_is_present` is compiled under `litchi_unified`
   and calls `litchi::open`.

The unit tests ran with all features and with `--features odg` alone, and
passed in both builds. The feature matrix (§2) runs every single-feature
build.

**Test and example fixes found by the matrix.** `--all-targets` clippy with
one format feature failed in the facade's own targets. These failures were
already there before this change:

- `tests/tolerance_report_facade.rs` would not compile with no format
  feature. It is now gated on `litchi_unified`.
- `tests/catalog_first_detection.rs` had a `Some(other)` arm gated on a
  13-feature list, which breaks a DOCX-plus-object-format build. The arm is
  now always present, with `allow(unreachable_patterns)`.
- `tests/catalog_detection_arbitration.rs` (`odf_package`) and
  `src/detection_smart/{tests,prepared_index_tests}.rs` (`repack`, `Read`,
  `archive_constructions`, `let mut cases`) had helpers that were unused in
  single-format builds. Each is now gated to the features that use it.
- `src/detection_smart.rs` gated the probe-count re-export for the tests on
  `xlsb`, which no test reads. That feature is removed from the gate.
- The `core_props_office` example declared `required-features = ["opc",
  "ooxml-common"]` but imports `litchi::{docx, pptx, xlsx}`. It now requires
  `["docx", "pptx", "xlsx"]`.

## 2. Feature-matrix check

`tools/check_feature_matrix.py` runs one cargo command at a time, in a fixed
order, through the `cargo` first on PATH (the memory-capped shim on this
host). It never starts two cargo processes at once, and `CARGO_BUILD_JOBS`
comes from the environment.

| Group | Commands |
|---|---|
| `facade` | `litchi --no-default-features`, then the same with each of the 21 format features alone. The features are derived from the manifest the way `build.rs` derives them, and the script refuses to run if it does not find exactly 21. |
| `facade-aggregate` | `legacy`, `ooxml`, `odf`, `iwork`, `word`, `slides`, `spreadsheets` |
| `cli` | `litchi-cli` once. Its manifest pins all 21 formats on `litchi-ops`, so it has no per-format features. |
| `py` | `litchi-py` once. It pins `legacy, ooxml, odf, iwork, rtf, …` on the facade, so it has no per-format features either. |
| `ops` (`--ops`) | `litchi-ops --no-default-features`, then each of its 21 format features. This overlaps with the coordinator's existing ops single-format clippy step. |

Flags:

- `--all-targets` also covers tests, examples and benches.
- `--clippy` runs `cargo clippy … -- -D warnings` instead of `cargo check`.
- `--only GROUP,…` runs the listed groups only.
- `--fail-fast` stops at the first failure.
- `--dry-run` prints the commands without running them.
- `--list` prints the derived format features.

The cli and py groups are generated from those manifests, so they gain
per-format steps automatically if either crate grows per-format features.

Runtimes on this host, with `CARGO_BUILD_JOBS=6` and the shared
worktree target:

| Run | Commands | Result | Wall time |
|---|---:|---|---:|
| `python3 tools/check_feature_matrix.py` (lib `cargo check`, first run of these feature sets) | 31 | 31 ok | 437 s |
| `--only facade,facade-aggregate --all-targets --clippy` (first run) | 29 | 19 ok, 10 failed (the failures listed in §1) | 370 s |
| the same after the fixes (incremental) | 29 | 29 ok | 18 s |
| the default run after the trunk merge (warm) | 31 | 31 ok | 202 s |
| `--only facade,facade-aggregate --all-targets --clippy` after the merge | 29 | 29 ok | 191 s |

Most of a cold run is dependency builds. Incremental reruns take seconds.

## 3. Batch-payload ingress fuzz targets

**Existing targets.** `litchi-core/fuzz` has `detect_format` and
`selector_grammar`. `litchi-ops/fuzz` has `paths`, `values` and
`query_plan`, plus `smoke_*` runners for stable. The selector and path
grammar target is therefore not missing: `selector_grammar` and `paths`
cover it. The batch-payload (OfficeCLI JSON) ingress target was missing
(record 0248).

**New targets** in `crates/litchi-ops/fuzz`:

- `batch_payload` fuzzes `batch::from_bytes_with_limits`, with limits of
  64 KiB and 64 items. Every input must either fail with `corrupt_batch` or
  `limit_exceeded`, or decode to `BatchItem`s that pass all of these checks:
  - each item validates and has a non-blank command;
  - each item is unchanged by `into_owned`;
  - each path field (`path`, `path2`, `parent`, `from`, `to`, `after`,
    `before`) lowers through `PathArg::parse` to a path whose canonical
    spelling reparses to the same path, or fails with `invalid_path`,
    `limit_exceeded` or `max_depth_exceeded`;
  - the selector lowers through `Query::parse` to a display fixed point, or
    fails with an in-bounds offset;
  - property keys are found again by `Props::get`, and numbers are finite;
  - the whole list serializes to the canonical wire form, decodes back to
    equal items, and stays stable when serialized again. Numbers may drift
    up to 4 ULPs (finding F1).
- `batch_apply`, behind the fuzz crate's `apply` feature (DOCX, XLSX and
  PPTX adapters), runs every decoded batch of at most 16 items against
  blank created DOCX, XLSX and PPTX documents.
  - Each batch runs atomically with `Unsupported::Warn`, twice on the same
    handle, and both runs must agree (the handle is immutable). It then runs
    best effort.
  - Every result must be a commit or a typed refusal. A panic or an
    `internal_error` fails the target, and a commit must have one result per
    item.
- `smoke_batch_payload` is the stable runner. It replays every seed, every
  prefix and every single-bit flip through `batch_payload`. With
  `--features apply` it also sends the seeds, every seventh prefix and every
  61st flip through `batch_apply`.

**Corpus** (`corpus/batch_payload/`, 26 files, 7.1 KiB):

- 19 hand-written cases mirroring `src/batch/tests.rs`: aliases, case
  folding, nulls, array and repeated props, scalars up to `u64::MAX`,
  duplicate and unknown keys, a blank command, a non-integer index, and the
  structural, placement, raw and read commands;
- 5 `officecli-*.json` seeds: the first six items of one batch from each
  OfficeCLI example script under `test-data/officecli/examples/` (OfficeCLI,
  Apache-2.0; `test-data/officecli/PROVENANCE.md`).
  `harvest_batch_seeds.py` regenerates them. Nothing references
  `3rdparty/`;
- 2 `regression-*.json` libFuzzer finds (F1, F2).

**Stable builds.** All targets type-check on stable with
`cargo check --manifest-path`, with and without `--features apply`. The
existing `libfuzzer-sys` dependency compiles on stable, so nothing needs
gating for `check`. The README records the gate commands and the
seed-harvest command.

**Runs.** This host has no `cargo-fuzz` and no nightly toolchain. The
libFuzzer targets were built on stable 1.95.0 with LLVM's coverage pass
instead (`-Cpasses=sancov-module` plus the inline-8bit-counters, pc-table
and trace-compares `-Cllvm-args`, `--cfg fuzzing`, debug assertions and
overflow checks, release build, `--target x86_64-unknown-linux-gnu`). This
recipe is in the README. No sanitizer was linked.

| Run | Inputs | Result |
|---|---:|---|
| `smoke_batch_payload` (release, 24 seeds before the regressions were added) | 61,350 inputs, 21,271 decoded (108,182 items) | ok, 0.7 s |
| `smoke_batch_payload --features apply` (release, 26 seeds, after the trunk merge) | 65,627 inputs, 22,199 decoded (109,299 items), 2,047 executions × 3 documents × 3 batches | ok, 4 s run (the first invocation's 106 s was mostly the release build) |
| libFuzzer `batch_payload`, first run | about 9.3 M executions | **crash F2** |
| libFuzzer `batch_payload`, after F2 | | **crash F1** (a long-number variant) |
| libFuzzer `batch_payload`, 300 s, after both | 22,726,145 executions, cov 3374, ft 17402 | ok |
| libFuzzer `batch_apply`, 300 s | 1,389,601 executions, cov 18660, ft 58642 | ok, rss 102 MB |

### Findings (follow-ups, not fixed here: the code is outside this record's area)

- **F1 — JSON numbers drift across wire round trips.** The workspace's
  `serde_json` is built without `float_roundtrip`, so decimal-to-`f64`
  conversion is not always correctly rounded.
  - `98446744073709551615` decodes to `9.844674407370957e19`; the correctly
    rounded value, per Python, is `9.844674407370955e19`.
  - Re-serializing and decoding moves the value again: 1 ULP, then 2 ULPs on
    a 121-digit literal.
  - A `PropValue::Number` taken from batch JSON can therefore differ from
    the author's decimal by an ULP or more, and the canonical wire form is
    not a fixed point.
  - Follow-up for the coordinator or owner: enable `serde_json`'s
    `float_roundtrip` workspace-wide, or parse property numbers with
    `fast-float2`, which is already in the lock graph. Either change affects
    every crate's JSON parsing and should be measured.
  - The harness tolerates up to 4 ULPs per comparison, and
    `regression-number-drift.json` keeps the case.
- **F2 — `PathArg`'s canonical spelling is not a fixed point of
  `PathArg::parse`.**
  - The relative path `'bod! my'` is not sheet notation, because the quoted
    name is not followed by `!`. It parses through the core grammar, and its
    canonical display is `bod! my`.
  - Fed back through `PathArg::parse`, that display reads as Excel sheet
    notation: sheet `bod`, so `/bod/ my`.
  - A CLI or batch caller that echoes a displayed relative path, containing
    `!` and no `/`, `[`, `]` or `"`, therefore addresses a different node.
  - Follow-up for the path owner (`litchi-ops/src/path`): quote such
    segments in `Display`, or make `split_sheet` refuse a name the core
    grammar would print unquoted.
  - The harness checks reparsing through the core grammar, as the existing
    `paths` target does, and `regression-patharg-bang-name.json` keeps the
    case.

## 4. Leftover parallel entry points (WP14)

**Decision.** Under ADR 0034 §§1, 2 and 4 and the single-entry-point DoD
item:

1. **`litchi::{pages,keynote,numbers}::semantic::{open, from_bytes}`, and
   the `ReadResult` aliases that existed only for them, are deleted.**
   - They were facade-defined free functions: openers parallel to
     `litchi::open` (which returns `File::Pages`, `File::Keynote` and
     `File::Numbers` since records 0296 and 0306), and duplicates of the leaf
     openers.
   - ADR 0034 puts open verbs on the owner and projection types
     (`Package::open`, then `document()`, `Presentation::open`,
     `Workbook::open`), and those are already reachable through
     `litchi::pages::*`, `litchi::keynote::*` and `litchi::numbers::Workbook`.
   - Replaced APIs are removed, not deprecated: CI uses `-D deprecated`, and
     ADR 0004 forbids aliases.
   - The `semantic` namespaces stay. They are documented leaf re-exports of
     the archive-free value types (cell values, selectors, limits, errors),
     which `numbers_semantic_values`, `numbers_semantic_view` and
     `semantic_facade` use. They are not entry points, and their docs now
     say so and name the openers.
   - Callers: only `tests/semantic_reader_facade.rs`. It now reads each
     fixture through `litchi::open`, checking the format and text, and
     through the leaf openers, keeping the typed wrong-family assertions.
     No caller exists in pyo3-litchi, litchi-cli, perf-baseline, examples or
     docs. The stale "semantic readers" comment in `pyo3-litchi/Cargo.toml`
     is corrected.
2. **`litchi::sheet::text` stays, as documented non-Office tooling.**
   - Its CSV, TSV, PRN, SYLK, DIF and fixed-width readers and writers open
     none of the 21 claimed formats (`litchi::open` refuses text input), so
     it is not a parallel opener.
   - Design §16.4 keeps it until it is re-homed "behind `import`".
     `File::import` (record 0266) is now the unified route for CSV and TSV
     into a workbook, with its own parser. `sheet::text` has no counterpart
     for standalone text-workbook reading or writing, or for SYLK, DIF and
     fixed-width.
   - Deleting it would drop those capabilities without a replacement, and
     deprecating it is not allowed.
   - Follow-up (not done here): move it into a leaf crate (`litchi-sheet`),
     so the facade only re-exports it, and bring its constructors in line
     with ADR 0034 §4 (`from_path_with_config` → `open_with_options`,
     `TextConfig` → `OpenOptions`) in the same change.

## 5. Duplicate ADR 0033 file

`docs/adr/0033-unified-api-glossary-and-default-policies.md` was a tracked
pointer to ADR 0034, but under the number 0033, which already belongs to
the unified-operations ADR. Nothing links to it:

- the goal's mention names the *base checkout's* draft;
- record 0016 cites that draft's hash as provenance;
- the ADR index lists only `0033-unified-operations-api.md` and
  `0034-…`.

The tracked path would also collide with the owner's untracked draft at the
same path in `/home/zhuhe/code/litchi` when this branch merges back. It is
removed. ADR 0034's "Adoption and numbering" section records the removal,
so the goal's reference still resolves to an explanation.

## 6. Gate commands for the coordinator

Add these to `(deleted-coord-artifact)`, with the program's build env and one cargo
command at a time:

```sh
# feature combinations: facade, aggregates, CLI, Python (31 checks)
python3 tools/check_feature_matrix.py
# warning-denied, all targets, facade and aggregates (29 clippy runs)
python3 tools/check_feature_matrix.py --only facade,facade-aggregate --all-targets --clippy
# parser fuzz targets build on stable
cargo check --manifest-path crates/litchi-core/fuzz/Cargo.toml --bins
cargo check --manifest-path crates/litchi-ops/fuzz/Cargo.toml --bins
cargo check --manifest-path crates/litchi-ops/fuzz/Cargo.toml --bins --features apply
# deterministic fuzz smoke on stable (seconds each, after the release build)
cargo run --release --manifest-path crates/litchi-ops/fuzz/Cargo.toml --bin smoke_batch_payload
cargo run --release --manifest-path crates/litchi-ops/fuzz/Cargo.toml --features apply --bin smoke_batch_payload
```

The fuzz crates are their own workspaces with no tracked lockfile, so give
them a separate `CARGO_TARGET_DIR` (for example `<worktree>/target/fuzz`),
or they rebuild the main target's dependencies. The facade unit tests of §1
run inside the existing `cargo test -p litchi --all-features` step.

## 7. Verification

After the trunk merge, every command below exited 0:

| Gate | Command | Result |
|---|---|---|
| fmt | `cargo fmt --all --check` | 0 |
| clippy | `cargo clippy -p litchi -p litchi-ops -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| clippy, no features | `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings` | 0 |
| clippy, facade per feature | `python3 tools/check_feature_matrix.py --only facade,facade-aggregate --all-targets --clippy` | 0 (29/29) |
| feature matrix | `python3 tools/check_feature_matrix.py` | 0 (31/31) |
| rustdoc | `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi -p litchi-ops` | 0 |
| ops tests | `cargo test -p litchi-ops --all-features`, in chunks (`--lib`, `--doc`, 7 groups of `--test`) | 0; 2,512 passed |
| facade tests | `cargo test -p litchi --all-features` | 0; 234 passed |
| pytest | `crates/pyo3-litchi/tests/run_tests.sh` | 0; 285 tests |
| perf-baseline | `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| boundaries | `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0; 0 |
| glossary | `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 |
| no 3rdparty | `python3 tools/test_no_3rdparty_references.py` | 0 |
| fuzz build | the three `cargo check --manifest-path` commands of §6 | 0; 0; 0 |
| fuzz smoke | both `smoke_batch_payload` runs of §6 | 0; 0 |

The capability and registry goldens were not re-blessed: no capability
changed. The OfficeCLI corpus is unaffected, since no behaviour visible to
it changed.

## 8. Not claimed

- No libFuzzer campaign ran under a sanitizer or on nightly, and the runs
  were 300 s each.
- F1 and F2 are reported, not fixed.
- `sheet::text` naming is not brought in line with ADR 0034.
- The matrix does not cover combinations of two or more formats beyond the
  aggregates.
