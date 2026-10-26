# 0369 — W0-C: the `litchi-iwa` exit parity harness and checked ledger

- Status: landed on branch `uo2/iwa-w0c` (cut from trunk
  `feat/unified-ops-api` at `9c53797811`, trunk `67315868f5` merged).
- Date: 2026-10-26.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) wave W0-C (§0), parity
  strategy §4, deletion-gate items 2 and 3 (§6); owner decision D1 and the
  macOS queue [0363-mac-queue](0363-mac-queue.md).
- `performance_claim`: none. The harness is test-only; no route, hot path or
  public API changed.

Names differ from the plan's wave table: the harness is
`crates/litchi-iwa/tests/parity/` (test target `parity`) instead of
`tests/exit_parity.rs`, and the ledger tool is `tools/iwa_exit_ledger.py`
instead of `tools/check_iwa_exit_crosswalk.py` (coordinator brief). No
`Cargo.toml` change was needed: the harness uses only the host's existing
dependencies.

## 1. What landed

| Path | What |
|---|---|
| `crates/litchi-iwa/tests/parity/main.rs` | the `parity` test target (deleted with the host in W3) |
| `tests/parity/harness.rs` | `Case`, `Format`, `Tier`, `Evidence`, `Fixture`, `LeafRun`, `leaf_bytes!` |
| `tests/parity/archive_diff.rs` | member-level and object-level three-way diff (input, host, leaf) |
| `tests/parity/known_fields.rs` | test-only Prost oracle: per-format message-type → schema table, known-field canonicalisation, unknown-field report |
| `tests/parity/{numbers,pages,keynote}_cases.rs` | the first parity cases (§3) |
| `tests/parity/self_test.rs` | the harness must fail, readably, on each difference class |
| `tools/iwa_exit_ledger.py`, `tools/test_iwa_exit_ledger.py` | the checked ledger and its 13 unit tests |
| `docs/unified-ops/results/0363/iwa-exit-ledger.tsv` | the tracked ledger (2,228 rows) |
| `docs/unified-ops/0363-mac-queue.md` | three W0-C rows (the merge write routes) |

`results/0363/host-pub-fn-crosswalk.tsv` stays as the planning snapshot that
seeded the ledger; the ledger is now the authority and the crosswalk is not
updated further.

## 2. Using the harness in a wave

A parity case expresses the operation once and hands the harness two
closures over the same fixture bytes:

```rust
#[test]
fn numbers_table_merge_parity() -> TestResult {
    let input = fixture("iwork/numbers/table-merges-native.numbers")?;
    let region = Region::new(2, 0, 2, 2)?;                // the operation, once
    Case::new(
        "numbers_table_merge_parity",                      // = the test name; the ledger cites it
        Format::Numbers,
        Tier::A,                                           // or Tier::B
        Evidence::Queued,                                  // or Evidence::Native("test-data/…-receipt.json")
        &input,
    )
    .host(|source| {                                       // host route: native IDs resolved from the selector
        let mut editor = litchi_iwa::numbers::Draft::from_bytes(source)?;
        let id = host_table_model_id(&editor)?;
        editor.merge_cells(id, region)?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {                                       // leaf route: the 0351 family
        let package = litchi_numbers::Package::from_bytes(source)?;
        let mut edit = package.edit_table_merges("Sheet 1", "shared-model")?;
        edit.merge(region)?;
        let commit = edit.commit()?;
        let restored = commit.package().apply_table_merges(&commit.patch().inverse())?;
        Ok(LeafRun::with_inverse(leaf_bytes!(commit.package()), leaf_bytes!(restored.package())))
    })
    .read("host table_cell_merges", host_merges)           // any number of reads, host or leaf
    .read("leaf table_merges", leaf_merges)
    .run()?;
    Ok(())
}
```

In this harness the tier and evidence go in the `#[test]` function body,
either as literals or as arguments to a helper. The ledger tool reads them
from there (§4).

`Case::run` asserts, and lists every failure in one error:

1. both routes succeed, and the host output is not the input (a no-op
   operation proves nothing). `.read_only()` marks a read-route case: both
   routes open and save, both outputs must equal the input, and the reads
   carry the comparison;
2. each output reopens strictly through the format's leaf `Package`;
3. every read gives the same result on both outputs (shown as a trimmed
   `Debug` diff when it does not);
4. when the leaf route returns `LeafRun::with_inverse`, the inverse patch
   restores the input bytes exactly;
5. `Evidence::Native(path)` names a file that exists;
6. **Tier A**: the outputs are byte-identical, the whole ZIP included. On a
   mismatch the error carries the first differing offset and the full
   three-way diff;
7. **Tier B**: no violation in the three-way diff. The host output decides
   what the operation touched:
   - a ZIP member whose content the host left equal to the input must keep
     its exact physical bytes in the leaf output (local record, compressed
     payload and central metadata; the central record's offset is not
     compared because it moves with earlier members);
   - member sets must match (order differences are notes);
   - in a changed `.iwa` member, object identifiers must match (identity
     allocation, §4.1 item 4);
   - an object the host left byte-identical must stay byte-identical in the
     leaf output;
   - a changed object must have the same payload types and object/data
     references, and each differing payload must decode through the
     registered schema to the same known fields;
   - a differing non-IWA member fails.

   Reported, not failed: ZIP layout differences of changed members,
   unknown-field differences (top-level field numbers the schema does not
   know, per side), object-order differences, and objects the host
   re-encoded but the leaf kept byte-identical.

A passing case prints its report to stderr (`cargo test -p litchi-iwa
--test parity -- --nocapture`), so a wave can read the notes of a Tier B
case.

**Schemas.** `KnownFields::for_format` registers the shared types the
current cases touch (401, 2001, 6001, 6002, 6005, 6006, 6011, 11006) and
each format's document types (Numbers 1/2, Pages 10000/10011, Keynote
1/2/4/5). A Tier B case whose changed object has an unregistered type fails
and names the type. Register the type in `known_fields.rs`, or per case
with `.known_fields(|k| k.with(type, "TSX.Name", decode_as::<tsx::Name>))`.
Type numbers overlap between applications, so register them per format.

**Choosing the tier.** Use Tier A when the leaf family is the code path the
host already delegates to, which holds for every route the crosswalk marks
`Y` with "host delegates". Use Tier B when the host re-encodes with Prost
and the leaf rewrites through the bounded wire layer. A route that passes
neither tier is not migrated: pin the divergence with a test that asserts
the failure, as `numbers_table_insert_row_divergence_is_reported` does, and
list the route instead of shipping it.

## 3. The first cases

All run on tracked native fixtures under `test-data/iwork/`:

| Test | Format | Route (host → leaf) | Tier | Evidence |
|---|---|---|---|---|
| `numbers_table_merges_read_parity` | Numbers | `table_cell_merges` → `Package::table_merges` (read-only) | A | `numbers/table-merges-native-receipt.json` |
| `numbers_table_merge_parity` | Numbers | `merge_cells` → `edit_table_merges` merge A3:B4 | A | queued |
| `numbers_table_unmerge_parity` | Numbers | `unmerge_cells` → `edit_table_merges` unmerge B11:C12 | A | queued |
| `pages_body_table_merges_read_parity` | Pages | `table_cell_merges` → `body_table_merges` (read-only) | A | `pages/body-table-merges-native-receipt.json` |
| `pages_body_table_merge_parity` | Pages | `merge_table_cells` → `edit_body_table_merges` merge A2:A3 | A | queued |
| `pages_body_table_unmerge_parity` | Pages | `unmerge_table_cells` → `edit_body_table_merges` unmerge C4:D4 | A | queued |
| `keynote_slide_table_merges_read_parity` | Keynote | `slide_table_cell_merges` → `slide_table_merges` (read-only) | A | `keynote/slide-table-merges-native-receipt.json` |
| `keynote_slide_table_merge_parity` | Keynote | `merge_slide_table_cells` → `edit_slide_table_merges` merge A2:A3 | A | queued |
| `keynote_slide_table_unmerge_parity` | Keynote | `unmerge_slide_table_cells` → `edit_slide_table_merges` unmerge B4:C5 | A | queued |
| `numbers_table_insert_row_divergence_is_reported` | Numbers | `insert_table_row` → `edit_table_topology` insert body row 1, `basic.numbers` | B, **fails** (pinned) | queued |

Every merge write case is Tier A, which is expected because the host
delegates to the leaf. The native receipts cover merges authored and read
in the native applications, not merges written by Litchi. The three merge
write routes are therefore queued for a macOS run (owner decision D1), and
the read routes cite the receipts.

**Finding: row insertion is not at parity.** The crosswalk marks Numbers
`insert_table_row` as `Y` (`edit_table_topology`), but on `basic.numbers`
the harness reports three violations:

- the leaf drops `preview.jpg`, `preview-micro.jpg` and `preview-web.jpg`,
  which the host keeps;
- the leaf changes object 904712 (type 6000), which the host leaves
  untouched;
- the two routes write a different `TST.Tile.numrows`: 4 for the host,
  2 for the leaf.

The existing `numbers_topology_formula_differential` test compares only the
calculation-engine objects, so it did not catch this. The ledger keeps the
four topology rows `host-only`, and W1-A owns the fix.

Self-tests (`self_test.rs`) prove that the harness:

- fails Tier A but passes Tier B, with an unknown-field note, when only
  unknown fields differ;
- fails Tier B when the leaf changes an object the host left untouched;
- fails on a known-field difference, showing the field diff
  (`table_name`), and on a differing read;
- fails when a changed type has no schema;
- fails on missing native evidence, a wrong inverse, and a no-op host
  route.

## 4. Using the ledger in a wave

`docs/unified-ops/results/0363/iwa-exit-ledger.tsv` has one row per host
`pub fn`. It has 11 tab-separated columns:

| Column | Content |
|---|---|
| `path` | path under `crates/litchi-iwa/src` |
| `name` | the host function name |
| `occurrence` | the function's 1-based position among same-name functions in that file; it replaces the line number, so edits elsewhere in the file do not break the key |
| `format` | the iWork format |
| `family` | seeded from the crosswalk, a planning aid |
| `status` | `host-only`, `migrated` or `retired` |
| `destination` | for migrated rows, `crate::Item::item` |
| `parity_test` | the test name |
| `tier` | `A` or `B` |
| `native_evidence` | a repository path, or `-` |
| `reason` | required for retired rows, for example "merged into `edit_x` verb `set`" |

The tool regenerates the inventory from source with the record 0363 §1.4
scan (2,228 functions today, identical to the crosswalk) and checks the
ledger against it.

A wave, per family:

1. Write the parity case(s) in `crates/litchi-iwa/tests/parity/<format>_cases.rs`.
2. Flip the host rows to `migrated`. Fill in `destination` with the leaf API
   path, `parity_test` with the test name, `tier` with the tier the test
   declares, and `native_evidence` with the path the test cites (or `-`).
   Host functions that collapse into a verb of another family, or disappear,
   become `retired` with a reason.
3. With no native evidence, append one row per route (format + destination)
   to `docs/unified-ops/0363-mac-queue.md`.
   `python3 tools/iwa_exit_ledger.py --print-mac-queue` lists every such
   route as `queued` or `MISSING` and prints the missing rows ready to
   paste. Replace the `<wave>` and `<fixture>` placeholders in those rows.
4. After adding or removing host functions, run
   `python3 tools/iwa_exit_ledger.py --update`. It adds the new functions as
   `host-only` and drops `host-only` rows whose function is gone. Migrated
   and retired rows stay when the host function is deleted, so W3 keeps the
   record.
5. Run `python3 tools/iwa_exit_ledger.py --check`, the gate.

`--check` fails when:

- a host function has no row;
- a `host-only` row is stale;
- a status is unknown;
- a retired row has no reason;
- a migrated row has no destination;
- the destination does not resolve, meaning some segment after the crate is
  not a `pub` item or `pub use` re-export in `crates/<crate>/src`;
- the parity test is missing, or is not a `#[test]` under
  `tests/parity/`;
- the row's tier or evidence differs from the test's declaration;
- the cited evidence file is missing;
- a migrated row without native evidence has no queue row with the same
  format whose Route cell starts with the destination.

`--summary` prints the counts. Today 9 rows are `migrated`: 3 per format,
the merge read and write routes. The other 2,219 rows are `host-only`.

## 5. Gaps and follow-ups

- **Coverage.** The harness covers 10 routes. Each W1/W2 wave adds its
  families' cases. The plan's replay of the 1,697 host unit tests through
  leaf families (§4.1) is not started; each wave ports the vectors of the
  families it migrates.
- **Row insertion.** The divergence in §3 goes to W1-A, together with
  column insertion and removal, which the harness has not run yet.
- **Unknown fields.** Detection is top-level only. Nested unknown fields
  still show up when the known fields differ, but they are not counted in
  the notes.
- **Destination check.** The check is lexical (a `pub` item or `pub use`
  with that name somewhere in the crate). It does not prove the path
  resolves. Rustdoc JSON would be exact, but it needs a nightly toolchain;
  `check_api_glossary.py` already pays that cost and could take this over
  in W3.
- **Coordinator gate.** Add `tools.test_iwa_exit_ledger` to the tool unit
  tests and `python3 tools/iwa_exit_ledger.py --check` to the `--check`
  runs (§6).

## 6. Verification

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-iwa --all-targets --all-features -- -D warnings` | 0 |
| `cargo test -p litchi-iwa --all-features` (1,697 unit tests, 19 integration targets including `parity`: 16 tests) | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_iwa_exit_ledger` (1,288 tests) | 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
| `python3 tools/officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check`, `check_feature_matrix.py`, `check_crate_boundaries.py`, `check_example_targets.py` | 0 each |
| `python3 tools/non_iwork_gate.py verify` | 1, pre-existing on trunk and independent of this change: "package 'litchi-doc' dependency 'flate2' path vendor/flate2 is not a workspace package" (the root manifest's tracked `vendor/flate2` patch) |

No library, public API, manifest or ops dispatch code changed, so the
litchi-ops/litchi/pyo3 suites, rustdoc, the glossary gate and the
capability goldens are unaffected and were not rerun.

Gate command for the coordinator:
`python3 -m unittest tools.test_iwa_exit_ledger && python3 tools/iwa_exit_ledger.py --check`.
