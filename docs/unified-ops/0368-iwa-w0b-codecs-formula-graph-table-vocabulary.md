# 0368 — W0-B: bounded codecs, the numbers-wire formula graph and the neutral table vocabulary

- Status: landed on branch `uo2/iwa-w0b` (cut from trunk
  `feat/unified-ops-api` at `9c53797811`; trunk `67315868f5` and
  `d7d0e1c857` merged).
- Date: 2026-10-26.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) wave W0-B (§0, §3.1, §7.3
  step 2), decision D4 (§8); owner decision D1 and the
  [macOS queue](0363-mac-queue.md); harness and ledger
  [0369](0369-iwa-exit-parity-harness-and-ledger.md).
- Donors (ported, never merged): `refactor/formula-graph-migration-20261023`
  and `test/formula-owner-budget-20261023` (`23737c7204`, `0f96331694`,
  `8743d9f106`, `accf4e036a`), the integration line
  `refactor/iwa-exit-current-20261023` (`fc29b16924`, `3f5d3cb497`),
  `repair/pages-body-shape-a408-20261024` (`ece66b17ff`),
  `refactor/keynote-shape-text-repair-20261024-r5` (`ecb83d14b9`),
  `refactor/iwa-attachment-reader-admission-20261023` (`5fcae790d4`).
- `performance_claim`: the formula-graph clone route is 17.8% cheaper than
  the host route (callgrind, §5); the table-value routes are unchanged
  (+0.02%); the theme decode is +3.1%. Two primitive micro-routes miss the
  5% bar and are listed in §5 and §6.

## 1. What landed

| Commit | Crate | What |
|---|---|---|
| `a9ffc79c85` | `litchi-iwa-protos` | the dependency codec refuses a CFUUID that carries both `uuid_bytes` and the four words (donor `0f96331694`) |
| `393b30bd5f` | `litchi-numbers-wire` | `formula_graph` (owner, cell-tile and range-tile remaps, calculation-engine owner registration) and `formula_owner` (owner identities, the empty table owner); glossary `required_signatures` |
| `831ca4e080` | `litchi-iwa` (tests only) | Tier A host-vs-leaf formula-graph parity on every Numbers fixture, the owner-map collision pin, the ABBA probe |
| `7a6b96ed14` | `litchi-iwa-common`, `litchi-numbers` | D4: cell data formats and topology coordinates move to `litchi_iwa_common::table`; Numbers re-exports them at their paths |
| `ebcf7aa0d6` | `litchi-iwa-protos` | theme, Keynote theme, theme presets, shape properties, Keynote shape creation and number attachment codecs with their Buffa projections and ratchets |
| `87e78bd7b3` | `litchi-numbers-wire` | `table_lifecycle` reference rewrites (donor `numbers_table_lifecycle_codec`) |
| `be8e4d164c`, `403ac199f2` | `litchi-iwa` (tests only) | read parity for the theme, drawable and number-attachment codecs on every Pages/Numbers/Keynote fixture; ABBA probes |
| this commit | docs, tools | this record, glossary rows, banned donor aliases |

No `Cargo.toml` changed and no crate edge was added:
`tools/crate_boundaries.json` is unchanged (`check_crate_boundaries.py` 0).
No public path changed. The host keeps its own implementations as the
parity oracle until W3; none of the host functions moved here is a `pub fn`,
so no row of the exit ledger (`results/0363/iwa-exit-ledger.tsv`) flips
(`iwa_exit_ledger.py --check` 0). No route is migrated in this wave — the
formula graph, the table-lifecycle rewrites and the codecs are primitives
that the W1 families consume — so no row is added to the macOS queue; each
W1 family that adopts them queues its own route.

## 2. The numbers-wire formula graph

`litchi_numbers_wire::formula_graph` is the host's
`numbers/editor/formula_clone/dependency_wire.rs` clone half (and
`numbers/formula_owner.rs`), ported from the donors and conformed to 0351:

- `remap_owner(source, Remap, &Limits)`,
  `remap_cell_record_tile`, `remap_range_precedents_tile` and
  `append_owners_to_engine(source, owner_object_ids, owner_map_entries,
  formula_count, &Limits)`; `Remap::new(object_ids, owner_ids, uuids)` takes
  sorted, unique maps and is validated by every call.
- One `Limits` profile per operation (`WireLimits`, references, scratch
  bytes, allocations) with checked `with_*` builders (`InvalidLimit` for 0 or
  above the hard ceiling; the donors silently clamped) and
  `Limits::DEFAULT` = hard ceilings. A private aggregate ledger charges
  fields, work, references, scratch and allocation events across every
  nested pass, so no nested payload resets a budget.
- Typed `Error::{InvalidSource, InvalidRemap, UnsupportedDependency, Limit
  { kind, observed, maximum }, InvalidLimit, Allocation}`; no `String`
  flattening, no `LimitExceeded`, no `*_with_limits` twin.
- The donors depended on twelve `*_with_limits` helpers that trunk's shared
  wire layer no longer has (record 0361a) and that W0-A owns
  (`litchi-iwa-common/src/wire.rs`). They are private to
  `formula_graph/wire.rs`, built on the public limit-taking scanners, with
  the shared layer's byte semantics (no-op rewrites stay byte-exact;
  rewritten repeated slots keep their keys and take canonical lengths).
- The donors rewrote an owner in twelve sequential path passes, each with a
  full preflight. The port walks the identity paths once
  (`rewrite_paths`); the result equals the sequential rewrite byte for byte
  (§3) and is what brings the route under the host (§5).

`litchi_numbers_wire::formula_owner::{Uuid, TABLE_OWNER_KIND,
encode_empty_table}` keeps the owner-UUID arithmetic (`for_table`,
`cfuuid_words`, `from_cfuuid_words`, `translate_from`). Unlike the donor
(which emitted the volatile and whole-owner envelopes empty and the
reference last), `encode_empty_table` emits the host's Prost bytes: the six
empty volatile sets, the whole-owner dependent-cell set and every envelope
in field order.

`litchi_numbers_wire::table_lifecycle` ports the donor's table
duplicate/remove reference rewrites (model UUID and name, the model,
table-info and data-list reference paths) with a required `&Limits` and a
typed `Error`. It lives beside the table engine its consumers already use
instead of in `litchi-iwa-protos`, which would have needed a new
`litchi-iwa-protos → litchi-iwa-common` edge.

## 3. Parity (Tier A)

`crates/litchi-iwa/src/numbers/editor/formula_clone/exit_parity.rs` (inside
the host, because the host routes are crate-private). For each of the 36
table formula-owner families on the tracked Numbers fixtures it builds the
remap `clone_table_formula_graph` builds and runs both rewrites on the same
source bytes:

| Object | Cases | Result |
|---|---:|---|
| `FormulaOwnerDependenciesArchive` | 396 | byte-identical |
| `CellRecordTileArchive` | 81 | byte-identical |
| `RangePrecedentsTileArchive` | 2 | byte-identical |
| `CalculationEngineArchive` registration | 36 | byte-identical |
| empty table owner (4 identity sets) | 4 | byte-identical to the host's Prost encode |

**Host defect found.** The host allocates internal owner identifiers one
above the largest owner *object* (`next_internal_owner_id`) and ignores
owner-map entries that have no owner object. On
`chart-arrangement-native.numbers` and `shared-model-native.numbers` the map
holds identifiers 20–24 beyond the owner objects, so a host duplicate writes
duplicate internal identifiers into the owner map. The leaf refuses that
remap (`Error::InvalidRemap`) instead of publishing a corrupt map;
`leaf_refuses_the_host_owner_map_collision` pins both fixtures. The parity
cases allocate above the map as well, which is the policy of
`table_create::next_internal_owner` and the one W1-A must adopt.

Read parity for the codecs (`theme/exit_parity.rs`,
`shapes/exit_parity.rs`), on every tracked Pages, Numbers and Keynote
fixture: 68 Pages/Numbers themes (identifier, both stylesheet references,
colour presets, preset UUIDs, and the five extensions, whose leaf bytes
decode to the host's typed extensions), 45 Keynote outer themes (UUID,
templates, default template, live video, the shared super), 315 drawables
(geometry and the four properties) and 86 number attachments (every field)
equal the host. The encoders (`theme_presets_codec`,
`keynote_shape_creation_codec`) and `table_lifecycle` have no host
primitive to compare against; their in-crate Prost differential tests came
with the donors, and the W1 families that use them prove parity through
their routes (§6).

## 4. D4: the neutral table vocabulary

The archive-free cell display and control formats (`DataFormat`, `Text`,
`Checkbox`, `StarRating`, and the `control`, `currency`, `custom`,
`date_time`, `duration`, `fraction`, `number`, `numeral_system`,
`percentage`, `pop_up_menu`, `scientific`, `text` modules) and the
section-relative `RowInsertion`, `ColumnInsertion`, `RowDeletion`,
`ColumnDeletion` move from `litchi-numbers` to
`litchi_iwa_common::table::{cell::data_format, topology}`. `Region`, `At`,
`Area` and the cell `Value` were already there. `litchi-numbers` keeps each
module at its path: a glob re-export of the common module beside its
`transaction` re-export, so `litchi_numbers::cell::data_format::currency::{
Currency, transaction::Edit}` resolve as before. The value unit tests moved
with the values. `crates/litchi-numbers/tests/neutral_table_vocabulary.rs`
proves every moved name is the same nominal type at both paths
(`TypeId`), which is what makes it a re-export and not an alias. The
glossary topology row now names the canonical owner. Pages and Keynote
gain no `litchi-numbers` edge; they re-export from `litchi-iwa-common` when
their W1 families first take these values (they have no data-format family
today, so no current path exists to preserve).

## 5. Performance

Callgrind, release, CPU 3 pinned, instruction counts (Ir), each route run
twice in ABBA order inside one probe (ignored tests in the host).

| Route (fixture corpus) | Host | Leaf | Δ |
|---|---:|---:|---:|
| formula-graph clone: 36 families × (owners + tiles + engine), host including its Prost decode of each object | 110,336,488 | 90,720,201 | **−17.8%** |
| · owners (792 calls) | 83,524,376 | 69,034,032 | −17.3% |
| · calculation-engine registration (72) | 17,821,424 | 18,376,322 | +3.1% |
| · cell tiles (162) | 2,349,699 | 3,182,678 | **+35.5%** (≈5k Ir, ≈2 µs per tile) |
| first leaf port, twelve sequential passes, for reference | 104,715,746 | 177,529,391 | +69.5% |
| D4 table values: 20 × (number-format set + row insert) on a created workbook, before/after the move | 981,534 / 981,309 | 981,738 / 981,505 | +0.02% |
| theme decode, 68 Pages/Numbers themes (host `IWorkThemeArchive::decode`) | 11,149,785 | 11,500,972 | +3.1% |
| drawable decode, 315 drawables (host Prost `TSD.DrawableArchive`) | 1,773,702 | 5,406,830 | **+205%** (≈5.7k Ir per drawable) |

The cell-tile primitive pays the strict dependency codec decode twice
(source and result), which the host replaces with one Prost decode. The
drawable codec cross-checks its strict parse against a Buffa lazy view
(`force_buffa`, 36% of its cost) and allocates its field table per message;
the host micro-route decodes only the `DrawableArchive` and validates
nothing. Neither primitive is a route yet: the host routes they serve
(table duplicate; shape geometry/property edits, which decode and re-encode
the whole shape archive with Prost) are measured end to end by W1-A and
W1-B/W1-C, which own the fix if the route misses.

Probe commands (binary from `CARGO_PROFILE_RELEASE_DEBUG=1 cargo test
--release -p litchi-iwa --lib --no-run`):

```sh
taskset -c 3 valgrind --tool=callgrind '--toggle-collect=*exit_parity::host_route' \
  target/release/deps/litchi_iwa-<hash> --ignored --exact \
  numbers::editor::formula_clone::exit_parity::formula_graph_route_cost_abba --test-threads=1
# likewise leaf_route; theme::exit_parity::theme_decode_cost_abba with
# {host,leaf}_theme_route; shapes::exit_parity::drawable_decode_cost_abba with
# {host,leaf}_drawable_route
```

The D4 probe was a scratch integration test of `litchi-numbers`, built at
`831ca4e080` (before) and `7a6b96ed14` (after); it is not committed.

## 6. Gaps and follow-ups

| Gap | Owner |
|---|---|
| Cell-tile remap +35.5% and drawable decode +205% as isolated primitives (§5) | W1-A (duplicate route), W1-B/W1-C (shape routes) |
| The host's internal-owner allocator collides with owner-map entries (§3); the leaf refuses | W1-A adopts the map-aware allocator |
| Keynote keeps its own `slide::table::number_format::{Number, DecimalPlaces, FixedDecimalPlaces, NegativeStyle, ThousandsSeparator, NumberError, MAX_DECIMAL_PLACES}` and `At`/`AtError`, duplicates of the common values. Re-exporting the common types would change Keynote's public error type (`NumberError` → `number::Error`, which has two more variants) | W1-C (owns Keynote `lib.rs`) |
| Pages/Keynote re-export of `table::cell::data_format`/`topology` when their families take them | W1-B, W1-C |
| Chart, 3D-chart, shape, media, colour and text value moves of the W0-B row (beyond the table vocabulary) | W1-E (shared chart/shape/media engine) |
| Host differential for the encoders (`theme_presets_codec`, `keynote_shape_creation_codec`) and `table_lifecycle` | the W1 families that call them |
| `theme_codec::ThemeReferenceSnapshot::encode` panics on a zero identifier (`encode_checked` is the fallible twin) | W1-E |
| The host's own formula-clone, owner, theme and codec paths stay as the oracle | W3 deletes them |

## 7. Verification

Commands on the merged branch (trunk `d7d0e1c857` merged), build env of the
coordinator rules (`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`).

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-iwa-protos -p litchi-numbers-wire -p litchi-iwa-common -p litchi-numbers -p litchi-pages -p litchi-keynote -p litchi-iwa -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`, and with `--features numbers`, `pages`, `keynote` | 0, 0, 0, 0 |
| `cargo test --all-features -p litchi-iwa-protos` (1,082 tests) | 0 |
| `cargo test --all-features -p litchi-numbers-wire` (208) / `-p litchi-iwa-common` (253) / `-p litchi-numbers` (1,073) | 0 / 0 / 0 |
| `cargo test --all-features -p litchi-pages` (716) / `-p litchi-keynote` (1,174) / `-p litchi-iwa` (1,809) / `-p litchi` (209) | 0 / 0 / 0 / 0 |
| `cargo test -p litchi-ops --all-features` in chunks: `--lib` (1,022), `--doc` (217), nine groups of 20 `--test` targets (1,485) | 0 each |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-iwa-common -p litchi-numbers -p litchi-numbers-wire -p litchi-iwa-protos` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (71 profiles) | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_iwa_exit_ledger` (1,289 tests) | 0 |
| `python3 tools/officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check` (after regenerating `results/0289/bench-matrix.md`) | 0, 0, 0 |
| `python3 tools/check_feature_matrix.py` (31 legs) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (301 tests) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

No capability or corpus-visible behaviour changed, so the capability and
registry goldens and the OfficeCLI corpus were not re-blessed (their tests
pass unchanged inside the litchi-ops chunks). Both pre-commit size checks
printed nothing for every commit.
