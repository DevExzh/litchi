# 0370 — W1-A: Numbers tables and cells, first slice

- Status: partial. Landed on branch `uo2/iwa-w1a`, which was cut from trunk
  `feat/unified-ops-api` at `f35f1c9f55`; trunk merged up to `31c87001d6`
  (records 0367, 0368a, 0371 cell writer, 0373).
- Date: 2026-10-26.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) wave W1-A (§0, §3, §4, §7).
  Harness and ledger: [0369](0369-iwa-exit-parity-harness-and-ledger.md).
  Owner decision D1 and the [macOS queue](0363-mac-queue.md).
- `performance_claim`: each migrated write route runs below its host route
  under callgrind (§4). The topology insertion route is 29.8% below the host
  route and the whole-cell text-style route 58.5% below.

This wave owns about 150 rows of the exit ledger. This slice migrates six
of them. It also lands the shared topology engine that Pages and Keynote
need, fixes the divergence 0369 pinned, and fixes three host defects. It
records the analysis that blocks the other families (§5), so the next
slice starts from evidence, not from the plan's parity guesses.

## 1. What landed

| Commit | Crate | What |
|---|---|---|
| `fa4aca99b0` | `litchi-numbers`, `litchi-iwa` | row and column insertion at parity with the host (§2); host defect fixes 1–3 |
| `80c9fe211d` | `litchi-iwa` (tests) | callgrind ABBA probe of the topology insertion routes |
| `74acb051e3` | `litchi-numbers-wire`, `litchi-numbers` | **shared engine**: `litchi_numbers_wire::table_topology` (§3) |
| `85f179159e` | docs | macOS queue rows for the routes whose host output the fixes changed |
| `17ae074f8a` | `litchi-iwa` | W1-B's host cell-writer fix (`e805eca0f5`, record 0371), cherry-picked so this branch's oracle matches trunk; trunk has it as `31c87001d6` |
| `df25e35046` | `litchi-numbers`, `litchi-iwa` (tests) | whole-cell text style at parity with the host (§2) |
| `c5c2d0ac59` | `litchi-iwa` (tests) | workbook catalog reads at parity |
| this commit | docs | this record |

Ledger: Numbers `migrated` rows went from 3 to 9. The six new rows are
`insert_table_row`, `insert_table_column`, `set_table_cell_text_style`,
`table_cell_text_style`, `sheets` and `tables`. `iwa_exit_ledger.py
--check` passes. None of these routes has a native receipt, and two of them
changed through host fixes, so each is queued for the macOS run.

## 2. Parity findings and fixes

### 2.1 Row and column insertion (the 0369 divergence)

0369 pinned three violations on `basic.numbers`. The harness then found more
on the native corpus. Each one is resolved on the side that native evidence
supports:

| Divergence | Native evidence | Resolution |
|---|---|---|
| leaf deleted the three root previews | the host and every cell-level leaf family keep them | the leaf topology family keeps them (`Working::keep_previews`) |
| leaf rewrote the drawable layout cache (`TableInfo` field 14) | the leaf's own code notes the application rebuilds a cache that does not describe the table | leaf leaves the cache as the host does |
| `TST.Tile.numrows`: host wrote the row extent, leaf the row count | all 23 tracked native fixtures where the two differ store the **count** | **host defect 1**: the host topology writes the count convention (`topology_tile_numrows`). W1-B applied the same convention to the host cell writer (record 0371) |
| column insertion: the host left the pre-BNC offset table naming the old columns | native rows keep both offset tables naming the same columns | **host defect 2**: the host splices the pre-BNC table too |
| a rewritten row's `has_wide_offsets`: host `Some(false)`, leaf absent | — | leaf states it, as the host encoder does |
| header or footer insertion: the host moved only the far edge of the formula owner's `body_range_for_table` | a body range must start after the headers | **host defect 3**: the host moves an extent that describes the table to the new geometry, and its wire check accepts a near-edge move |

Result: `numbers_table_insert_row_parity` and
`numbers_table_insert_column_parity` pass tier A on `basic.numbers`.
`numbers_table_topology_corpus_parity` runs 7 insertion positions (header,
body and footer rows; header and body columns) on the whole corpus. It
compares 140 routes; 107 are byte-identical and every one passes tier B. The
rest differ only in the Snappy framing of a changed member. Where the host
refuses, the leaf refuses too (formulas read from another table, volatile
dependency state, merges). On the fixtures whose table-appearance read the
host rejects, the leaf edits.

Removal is **not** at parity, and `numbers_table_removal_divergence_is_reported`
pins it. The leaf clears the removed cells through the cell writer first.
That deletes the previews, releases the cells' format keys and lowers the
perpendicular header counts. The host keeps the previews, leaks the format
keys, keeps stale header counts and leaves the pre-BNC tables naming the
removed column. Several of those are host defects, so the fix is a host
change plus the leaf's preview policy. See §5.

### 2.2 Whole-cell text style

The leaf `edit_table_cell_text_style` was ported from the host's
`cell_paragraph_style` (record 0309). The harness found six differences. The
leaf now matches the host on each:

- one variation per cell (the leaf used to share one across cells);
- a property equal to the inherited value stays inherited (unset emphasis
  reads as off);
- a new style lists only its parent in its message-info object references.
  Every tracked native style omits the stylesheet back-reference, and the
  leaf used to list it;
- the variation carries an empty `para_properties`;
- when the stylesheet's header cannot be transitioned (field-local
  references the aggregate list does not repeat), its aggregate list still
  names the new style;
- the root previews stay.

An empty cell now gets a style-only BNC cell, counted by its row and column
headers, as the host stores it. The leaf used to refuse it with
`Refusal::EmptyCell`; that variant is deleted. New style UUIDs are
deterministic on the leaf and random on the host. The harness's documented
`leaf_structural_metadata` policy (from W1-C, record 0372) compares them as
equal after checking their shape.

Tests: `numbers_cell_text_style_parity` and
`numbers_cell_text_style_shared_parity` pass tier B on `basic.numbers`.
`numbers_cell_text_style_corpus_parity` compares 36 corpus routes (B2 of the
first table), all tier B, including three empty cells.
`numbers_cell_text_style_read_parity` compares the host's
`table_cell_text_style` with the leaf's effective properties on 72 cells. An
unset size reads as the host's 12 pt default.

Only the point size, bold and italic setter migrates. The host's other 77
rich-text rows (font, colour with alpha and colour space, double and wavy
underline, double and triple strikethrough, baseline shift, capitalization,
character spacing, ligatures, outline, script, shadow, background, every
paragraph and list property, and their `reset_*` verbs) need properties that
`litchi_iwa_text_wire::CharacterProperties` cannot express. That is the
W1-D text engine (§5).

### 2.3 Workbook catalog reads

`numbers_workbook_catalog_read_parity` checks two reads on every fixture the
host opens. Sheet names equal `Package::sheets` on 37 fixtures. Table
names, dimensions and appearance equal `Sheet::tables` and
`Package::table_appearance` on 23 fixtures. The host's `tables` refuses the
other 14 when its appearance read fails.

## 3. The shared engine: `litchi_numbers_wire::table_topology`

The payload rewrites of a table-axis edit moved from `litchi-numbers`
(`package/table_topology/{storage,lists,calc}.rs`, history kept with
`git mv`) into the wire crate. Pages and Keynote can pair them with their
own object plumbing, as they already do with `table_create`:

- `storage::{AxisEdit, tile_rows, tile_columns, header_bucket, UidMap,
  strokes, StrokePlan, bnc_cells, replace_cell, HEADER_BUCKET_ROWS}`: tiles,
  both offset tables, header buckets, the row/column UID map, stroke layers
  and single-cell replacement;
- `lists::{release, release_with}`: `TST.TableDataList` reference release;
- `extents::{Extents, rewrite_extents}`: a table formula owner's
  total and body ranges.

They take a typed `table_topology::Error { Invalid, Unsupported,
Allocation }` and use the shared wire layer's default limits, as they did in
`litchi-numbers`. `litchi-numbers` maps that error into its structural
`Fault` (`From` impl). The formula-coordinate shift (`shift.rs`) stays in
`litchi-numbers` because it walks the package's calculation engine. No
crate edge changed. The coordinator merged `74acb051e3` to trunk as
`7faefb5783`.

## 4. Performance

Callgrind, release with debug information, CPU 3 pinned, instruction counts
(Ir). Each probe runs host route A, then leaf route B twice, then A again on
the same bytes: open, edit, commit, serialize
(`crates/litchi-iwa/tests/parity/cost_probes.rs`).

| Route | Host | Leaf | Δ |
|---|---:|---:|---:|
| topology insert: `basic.numbers` and `custom-number-native.numbers` × body row 1 and body column 0 | 754,419,162 | 529,471,532 | **−29.8%** |
| whole-cell text style: `basic.numbers`, one cell, then two | 810,771,098 | 336,488,941 | **−58.5%** |

```sh
CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa --test parity --no-run
taskset -c 3 valgrind --tool=callgrind '--toggle-collect=*cost_probes::host_topology*' \
  target/release/deps/parity-<hash> --ignored --exact \
  cost_probes::numbers_topology_insert_cost_abba --test-threads=1
# leaf_topology; host_text_style / leaf_text_style with numbers_text_style_cost_abba
```

## 5. Gaps, with the evidence that sets each one's order

| Family (ledger rows) | State | Blocking finding | Next step |
|---|---|---|---|
| row and column removal (2) | host-only, divergence pinned | §2.1: the leaf clears before it removes; the host leaks format keys, keeps stale perpendicular header counts and stale pre-BNC tables | fix those host defects (queue them), give the leaf removal path the preview-keeping policy, re-run the corpus sweep |
| conditional highlights (4) | not landed | The donor `refactor/numbers-conditional-highlight-admission-review-20261024` was ported and compiles on trunk (its popup, protos and core API drift resolved; Error conformed to `Refused`/`Limit { observed, maximum }`), but each route cannot read the other's output: the host rejects the leaf's predicates as an "unsupported predicate representation", and the leaf rejects the host's style set. The host's style set is 1,155 bytes against the leaf's 711; the leaf writes applied-rule 0 where the host writes 15 ("none"); the host draws random owner UUIDs. The donor graph is not the host's | port the host's `conditional_highlight/{native,write,dependencies}` graph instead of the donor's, or decide which graph the macOS run accepts. The port is kept as a patch outside the repository (`(deleted-coord-artifact)/ch-port-wip.patch`, 433 KB) |
| sheet and table add/remove (4) | host-only | an exploratory tier-B run on `basic.numbers` (add) and `table-relocation-native.numbers` (remove) found: the leaf structural workspace deletes previews, rewrites the sidebar tree node and allocates identifiers in another order; the host route refuses the remove fixture (appearance read) | align identifier allocation order and the sidebar node with the host, or record them as intentional differences in the harness policy |
| table duplicate and resize, sheet duplicate (3) | host-only | donors `refactor/iwa-table-lifecycle-native-tests-20261023` (4.2k lines) and the integration line's `sheet_duplication.rs` (7.6k) predate the 0351 contract and the `structure::Working` workspace; the map-aware owner allocator they need is now on trunk (`formula_graph::next_owner_id`, record 0368a) | build `duplicate`/`resize` as verbs of the existing `edit_table_lifecycle` on `Working`, reusing `table_lifecycle` and `formula_graph` from the wire crate |
| cell rich text, paragraph and list (77 of 79) | host-only | §2.2: `CharacterProperties` has 7 properties; the host vocabulary (`litchi_iwa_text::{character, appearance, paragraph}` values) needs W1-D's storage-scoped engines in `litchi-iwa-text-wire` | W1-D extends the property values; this wave adds the cell families on the same style-variation path as §2.2 |
| generic `DataFormat`, numeral system, list number format (8) | host-only | the leaf has 11 dedicated format families and no numeral-system family | one family taking the neutral `DataFormat` that delegates to the dedicated writers, plus a numeral-system writer |
| cell fill, border, layout (9) | host-only | not started; borders are stroke-sidecar layers (the shared `table_topology::storage::strokes` reads them) | — |
| hidden axes (2) | host-only | Pages has a 7.9k-line leaf family over `pages_hidden_state_codec`; Numbers and Keynote have none | lift Pages' payload rewrites into the wire crate (W1-B owns the Pages module) or port the host's 1.2k-line `table_hidden_axes.rs` |
| cell comment fallbacks (5), pivot categories (1), formulas and the formula AST (2 + `numbers/formula.rs`), Builder `table_dimensions`/`language`/`locale` (9), sort (5), io (4), topology oracle (1) | host-only | the Builder's host output is not deterministic (random identity and table UUIDs); the leaf Builder starts from a template | — |

## 6. Verification

The build env follows the coordinator rules (`CARGO_BUILD_JOBS=6`,
`RUST_TEST_THREADS=8`), on the branch with trunk `31c87001d6` merged. The
pre-commit size and binary checks printed nothing for every commit.

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-numbers-wire -p litchi-numbers -p litchi-iwa -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`, and with `--features numbers` | 0, 0 |
| `cargo test -p litchi-numbers-wire --all-features` / `-p litchi-numbers` / `-p litchi-iwa` (1,702 unit tests, every integration target, `parity`: 29 passed, 2 ignored probes) / `-p litchi` | 0 / 0 / 0 / 0 |
| `cargo test -p litchi-ops --all-features`: `--lib`, `--doc`, nine groups of `--test` targets | 0 each; `numbers_unified` first failed on an assertion that the empty-cell font set is refused, and passes since `ebf98b6baa` updated the assertion to the new behaviour |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers-wire -p litchi-numbers` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile` `numbers-wire`, `numbers-minimal`, `numbers-full` | 0, 0, 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0; 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_iwa_exit_ledger` (1,289 tests) | 0 |
| `python3 tools/iwa_exit_ledger.py --check`, `officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check` | 0 each |
| `cargo test -p litchi-cli --all-features --test officecli_corpus` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `python3 tools/check_feature_matrix.py` (31 legs) | 0 |

The capability and registry goldens pass unchanged inside the litchi-ops
chunks. This slice adds no unified capability (the empty-cell style is a
leaf behaviour the existing `font.*` set already reaches), so they were not
re-blessed.
