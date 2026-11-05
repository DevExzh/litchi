# 0300 — iWork and tooling hygiene

- Status: implemented on `uo2/iwork-hygiene`, cut from
  `feat/unified-ops-api` at `ea935b7194`; trunk merged at `8a6e6968e1`
  (records 0295, 0299) and `928b19b4db` (record 0297) before the Numbers
  and Pages changes.
- Date: 2026-10-15
- Owner decision applied: **D2** (the iWork crates may be edited),
  2026-10-15. `litchi-keynote` is untouched (record 0298 is reworking it).
- Scope: the pre-existing failures that records 0296–0299 found outside
  their own work, so the iWork part of the workspace can be fully green.
- `performance_claim`: none. No hot path got slower (§7).

## 1. Release-only builder failure (record 0297 §4)

**Root cause.** `litchi-iwa-package`'s `EntryStoreState::rebuild_positions`
inserted each entry name into the name index *inside* `debug_assert!`:

```rust
debug_assert!(
    self.positions.insert(Arc::clone(&entry.name), position).is_none(),
    "entry names must be unique"
);
```

Release builds compile the assertion away, so after any structural change
(`try_insert_at`, `remove_at`) the index was empty and every by-name lookup
missed. `NumbersDocumentBuilder::build()` inserts its archives into a new
package and then reads `Index/Document.iwa` by name, hence
`Bundle("IWA package entry not found: Index/Document.iwa")` in release only.
Any other release path that inserted or removed a package entry and then
looked an entry up by name had the same bug.

**Fix** (`84c281f563`). The insertion runs unconditionally; only the
uniqueness result is asserted.

**Evidence.** Before the fix, `cargo test --release -p litchi-iwa-package`
failed two existing tests (`structural_mutations_rebuild_the_index`,
`patches_are_reversible_and_source_checked`); after it, 9/9 pass. The new
`crates/litchi-iwa/tests/release_builders.rs` (`NumbersDocumentBuilder::build`
plus `to_bytes`, and `NumbersEditor::create`) failed 0/2 in release before
and passes 2/2 after; run it with
`cargo test --release -p litchi-iwa --test release_builders`.

**Sweep.** A scan of every `debug_assert*!` in `crates/` for mutating calls
inside the macro found one other case: `litchi-ods/src/model/cell.rs`
(`debug_assert!(rich_text.remove_hyperlink(index))`). That file is not in
the module tree (`model/mod.rs` declares no `cell`, and neither `row.rs` nor
`sheet.rs`), so it is dead code and was left alone; see §8.

## 2. `litchi-iwa` chart examples (`971f06eb0c`)

`ChartData` moved to `litchi-iwa-common`, whose constructor returns the
typed `chart::DataError`, and `litchi_iwa::Error` had no conversion, so the
six `create_{pages,keynote,numbers}_{pie,donut}_chart` examples (helpers
returning `litchi_iwa::Result<ChartData>`) did not compile. Following the
crate's `#[from]` policy for leaf errors (ADR 0034 §6, no string
flattening), `Error` gains `ChartData(#[from] litchi_iwa_common::chart::DataError)`,
and the helpers use `Ok(ChartData::new(..)?)`. All six examples build and
write a document (7.8–12.0 KB each).

## 3. `litchi-iwa` clippy

`cargo clippy -p litchi-iwa --all-targets --all-features -- -D warnings`
failed on test-module, integration-test and example code; all fixed
without behaviour change (`ee67eb8e7c`, `e014333e0b`, `71496f6d93`):

- **Deprecated native IDs.** 232 reads of
  `KeynoteSlideTextInfo::drawable_object_id` (177 in test modules, 55 in
  examples), 23 of `KeynoteSlideInfo::{slide_id,node_id}` and two calls of
  the deprecated `PagesEditor::set_drawable_comment`. Examples that only *printed* the
  native IDs (`add_keynote_slide`, `duplicate_keynote_slide`,
  `inspect_reachable_media`, the slide line of `inspect_keynote_text_boxes`)
  no longer print them. Code that passes the IDs to the legacy raw-ID
  `KeynoteEditor`/`PagesEditor` API, which these tests and examples exist to
  exercise and which has no `litchi_keynote::Package` route for free text
  boxes, carries a per-function `#[expect(deprecated, reason = …)]`
  (narrowest scope; `expect` fails once the use disappears).
- **Lints:** `field_reassign_with_default` (10, struct-update syntax),
  `needless_question_mark` (3), `items_after_test_module`,
  `unneeded_struct_pattern`, `needless_update`, `needless_borrow` (2),
  `manual_contains` (2), five `collapsible_if` as let-chains, and a dead
  assignment in `create_pages_table`.
- **Shared fixtures.** `litchi-iwa` compiles two test-support files of
  other crates through `#[path]`. `litchi-iwa` has no `[lints]` table, so
  clippy's defaults (which deny `collapsible_if`) apply to them, while their
  owners use the workspace table, which allows it. The Numbers fixture's
  `collapsible_if` is fixed in place. The Keynote fixture belongs to
  `litchi-keynote`, so its `mod` declaration in
  `keynote_focused_drawable_comments.rs` carries
  `#[allow(clippy::collapsible_if, clippy::clone_on_copy,
  clippy::needless_question_mark, reason = …)]`; record 0298 already removes
  the last two in `litchi-keynote`.
- **Doctests** (`17a252c68f`). Two module examples still used the
  `litchi_numbers::Document` / `litchi_pages::Document::open` names removed
  by the W1 renames of 0297/0299.

`cargo clippy -p litchi-iwa -p litchi-numbers -p litchi-pages --all-targets
--all-features -- -D warnings` now passes.

## 4. `tools/test_check_crate_boundaries.py` (`c3273336f5`)

`test_iwa_package_edge_is_archive_owned_and_host_edge_cannot_return` pinned
270 policy edges. The 11 extra edges are all registered and intended:
seven `litchi-ops → {odb, odc, odf-formula, odg, odi, odm, oth}` from record
0282 and four `litchi-ops → {iwa-detect, keynote, numbers, pages}` from
record 0296 (owner decision D1), each with its ADR 0024 amendment. (The
brief named record 0266; 0266 set the 270, and 0282/0296 added the edges.)
The test now compares the parsed canonical and migration edge sets with the
edges the reviewed JSON lists, instead of a count, so registering an edge
does not break it while a parser that drops, invents or merges an edge
still does.

## 5. `tools/check_iwork_leaf_public_api.py` (45 findings → 0)

All 45 findings were in `litchi-numbers`. The checker stops at the first
failing crate, so Pages and Keynote were checked separately with `--crate`.
Both were clean before and after this record.

Triage:

| Findings | Verdict | Action |
|---|---|---|
| `table::Position`, `table::Range` (aliases of `CellPosition`, `CellRange`; also re-exported at the crate root and in `litchi::numbers`) | **Real ADR 0034 violation**: the aliases were documented "narrow migration names" (§1, renamed APIs are removed, not aliased), and a coordinate type named `Position` collides with §8, which reserves `Position` for positional selection and keeps grid coordinates out of the selector vocabulary | Deleted (`b661d68f22`); callers in `litchi-numbers`, `litchi-iwa` and the facade use `CellPosition`/`CellRange`; `numbers-minimal`/`numbers-full` ban the four names; the glossary's grid-coordinate row is updated |
| 43 references from the Numbers table facade (`Table`, `Builder`, `SharedTable`, `from_shared`/`into_shared`, re-exports) to `litchi_iwa_common::table::{model, coordinate, cell::value}` | **Intended exception**: these three modules are the archive-free table model that Pages, Keynote and Numbers share (each imports only `std` and the others); the checker forbade the whole `litchi_iwa_common` crate because it also holds wire helpers | Module-granular allowlist `ALLOWED_SEMANTIC_MODULES` with a reason per module (`5c6c512966`); the rest of `litchi_iwa_common` stays forbidden, and a physical name inside an allowed module still fails |

New checker tests: the allowlist is prefix-exact (sibling modules such as
`table::merge`, the bare module path and `wire::*` still fail; an allowed
module cannot launder a `WireView`), and a source test reads the three
modules and fails if they import anything but `std` and each other.
`python3 tools/check_iwork_leaf_public_api.py` exits 0.

## 6. Deterministic table merges (record 0297 §4)

**Cause.** The only ambient input was the merge-owner UUID. When a merge
creates a table's merge owner, `litchi-numbers`
(`package/table_merges/transaction.rs`) passed
`litchi_core::id::generate_guid_bytes()` to
`litchi_numbers_wire::table_merges::rewrite_table_merges`, so two identical
merges published different bytes. This went against ADR 0034 §9 (writers
never consult ambient randomness). `litchi-pages` body-table merges and
`litchi-keynote` slide-table merges call the same writer the same way.

**Fix.**

- `litchi-numbers-wire` (`c270453b89`) adds `rewrite_table_merges_derived`
  and `derive_merge_owner_uuid`. The owner UUID is a fixed bijection of the
  table model's own 128-bit `table_id`: a domain XOR, then odd-multiplier
  and xor-shift rounds, then the RFC 4122 version-4/variant bits, so it has
  the same shape as the random owners it replaces. Identical rewrites are
  byte-identical and distinct tables get distinct owners. The writer had
  already parsed `table_id`, so the change adds no parsing. The
  supplied-UUID entry point stays for the Keynote caller.
- `litchi-numbers` (`9f305a3508`) and `litchi-pages` (`5d793ea615`) call
  the derived entry point.

**Tests.** `litchi-numbers`: `identical_merges_are_byte_identical` merges
the same region into two fresh `Package::new()` workbooks and compares the
published bytes. `litchi-pages`: `merge_creating_an_owner_is_byte_deterministic`
strips the native fixture's merge owner, merges twice and compares the
bytes. Both fail when the transaction change is reverted. Wire unit test
`derived_owner_rewrites_are_deterministic_and_table_specific`.

## 7. Performance

No new or slower hot path, so no paired measurement was made:

- `rebuild_positions` does in release exactly what it already did in
  debug. Release builds used to skip the work and then return wrong
  results.
- The derived owner UUID (two 128-bit multiply/xor-shift rounds) replaces
  a call into the thread-local RNG, on a path that runs once per created
  merge owner.
- The other changes affect only names, tests, examples and tools.

## 8. Gaps and follow-ups

1. **Keynote merges are still random**: `litchi-keynote`
   `slide_table_merges/transaction.rs` passes `generate_guid_bytes()`.
   Record 0298 should switch to `rewrite_table_merges_derived`. Then delete
   the supplied-UUID `rewrite_table_merges` (no other caller).
2. **Other ambient randomness in Numbers edits**:
   `table_cell_display_format_native.rs::fresh_custom_uuid` draws a random
   custom-format UUID when a display-format edit registers a new custom
   format, so those edits are not byte-deterministic either. This record
   did not change it: deriving it needs a registry-aware scheme.
3. **Coordinates vs `At`/`Area`**: Numbers (and the shared
   `litchi-iwa-common` model) still name grid coordinates
   `CellPosition`/`CellRange`. ADR 0034 §8 names them `At`/`Area`. This is
   a cross-family rename for the next iWork wave.
4. **`litchi-iwa` lint table**: the crate has no `[lints] workspace = true`,
   which is why the shared fixtures are linted differently there. Adopting
   the workspace table is a separate review: it turns on the workspace's
   rustc `deny` set.
5. **Guard against side effects in `debug_assert!`**: Clippy's
   `debug_assert_with_mut_call` (nursery) would have caught §1. The sweep
   found no other live case, so it could be denied workspace-wide. That
   needs a full-workspace clippy run, which this record did not do.
6. **Dead `litchi-ods` files**: `src/model/{cell,row,sheet}.rs` are
   outside the module tree. They are either leftovers to delete or a missed
   `mod` declaration, and the ODF owner should decide which.

## 9. Verification

All with `CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`; every command below
exited 0 on the final tree. Test counts: `litchi-iwa` 1782,
`litchi-numbers` 1046, `litchi-pages` 683, `litchi-numbers-wire` 164,
`litchi-iwa-package` 9 (also 9 in release), `litchi-ops` 2295, `litchi`
260, `litchi-cli` 32; none failed.

Two gate failures surfaced on the merged tree and are fixed here:
`litchi-numbers-wire` did not document under `-Dwarnings` (two module-doc
intra-doc links, `7df6d71d01`), and the crate-boundary table-merge audit
required the literal `rewrite_table_merges` token in each focused
transaction (now also `rewrite_table_merges_derived`, `7b5b537cc5`).

- `cargo fmt --all --check`.
- `cargo clippy -p litchi-iwa -p litchi-numbers -p litchi-pages
  -p litchi-numbers-wire -p litchi-iwa-package -p litchi-ops -p litchi
  -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings`;
  `cargo clippy -p litchi-ops --no-default-features [--features numbers |
  pages] --all-targets -- -D warnings`.
- `cargo test --release -p litchi-iwa-package`; `cargo test --release -p
  litchi-iwa --test release_builders`; `cargo test` (all features) for
  `litchi-iwa-package`, `litchi-numbers-wire`, `litchi-iwa`,
  `litchi-numbers`, `litchi-pages`, `litchi-ops`, `litchi`, `litchi-cli`.
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features` for the
  changed crates.
- `python3 tools/check_crate_boundaries.py`,
  `python3 -m unittest tools.test_check_crate_boundaries`,
  `python3 tools/non_iwork_gate.py verify`,
  `python3 tools/check_iwork_leaf_public_api.py`,
  `python3 -m unittest tools.test_check_iwork_leaf_public_api`,
  `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile
  {numbers-minimal, numbers-full, pages, ops-iwork}`.
- `crates/pyo3-litchi/tests/run_tests.sh`;
  `cd tools/perf-baseline && cargo check --locked --all-targets`.
- No capability, registry, corpus or OfficeCLI-visible behaviour changed,
  so no golden was re-blessed. The capability and registry golden tests run
  inside `cargo test -p litchi-ops --all-features` without blessing.
