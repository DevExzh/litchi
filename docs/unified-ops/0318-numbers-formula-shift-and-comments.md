# 0318 — Numbers formula-coordinate shift, Numbers cell comments; iWork table creation not started

- Status: implemented on `uo2/iwork-numbers2`, cut from trunk
  `feat/unified-ops-api` at `58f4a443c1`.
- Date: 2026-10-16
- Inputs: records 0297–0314 and `IWORK_API_REQUESTS.md` (0312 §5 and 0314
  §5 give the diagnoses). The work list, in order: the Numbers
  formula-coordinate shift on row and column edits (records 0309, 0314);
  Numbers cell comments (request 19); Keynote and Pages table creation;
  Keynote speaker-notes run styles if time permitted.
- Owner decisions applied: **D1**, **D2** (2026-10-15): the iWork crates are
  edited directly.
- Commits: `7861bf1f94` (the shift, `litchi-numbers`), `fe6a08a91a` (host
  oracle and differential test, `litchi-iwa`), `c7f39b0d91` (comments on
  native packages), `64b40c6382` (`Package::table_cell_comments`),
  `0f87ea6dc7` (unified wiring, goldens, delta table), `66e555e1b2` (the
  empty-comment-list fast path), `f03d744dfc` (bench routes), the trunk
  merge, and this record.
- `performance_claim`: paired leaf/unified medians in §4 (release builds of
  `--features numbers`, one pinned core of a shared host, in-process ABBA,
  processes B A A B). Evidence: [results/0318/](results/0318/)
  (`measure.sh`, `summary.txt`, raw outputs, `binaries.sha256`; binaries not
  committed).

## 1. Summary

| Item | Status |
|---|---|
| Formula-coordinate shift on Numbers row and column edits | **done** (§2): own formulas, other tables' and sheets' formulas that read the table, ranges, absolute and relative flags, footer ranges; deleted references refused, typed |
| Differential test against the retired host editor | **added** (§2.4): 48 corpus and synthetic cases byte-equal on the calculation engine and formula lists, 318 plain ones, 53 packed-edge ones; the host editor's shared-entry defect found (§2.5) |
| Numbers cell comments (request 19) | **done** (§3): per-table census, FieldInfo-free lists, effective locators, a writer over several members; read, set, add and remove through `File` |
| New-member path for a table without a comment list | **not done**: a new list still goes beside the table model (§3.4) |
| Keynote and Pages table creation | **not started** (§5) |
| Keynote speaker-notes run styles | **not started** |

Existing read paths and existing write results are unchanged: the
query-corpus differential digests, corpus fidelity and the CRUD checklist
pass without re-blessing. The capability and registry goldens (both
feature sets) and the OfficeCLI delta table change only for the Numbers
`comment` key and the narrowed row and column gap.

## 2. The formula-coordinate shift (`litchi-numbers`, `table_topology/shift.rs`)

### 2.1 What it is

A wire-level port of the retired host editor's `formula_dependency_shift`
(`litchi-iwa/src/numbers/editor/formula_dependency_shift{,/ast,/wire}.rs`,
about 4,000 lines over Prost archives). `litchi-numbers` keeps exact
payloads and has no Prost, so the port reads the archives' fields itself
and rewrites them with the same bounded primitives the host editor's wire
layer calls (`litchi_iwa_common::wire::{patch_varint_field,
transform_length_delimited_field, rewrite_repeated_length_delimited_fields,
rewrite_repeated_varint_fields, patch_nested_varint_field}`), field by field
in the same order, so a payload it rewrites carries the host editor's
bytes. `table_topology::plan` calls it in place of the former
`calc::check`, which refused every table with formulas; `calc.rs` keeps the
table-extent rewrite.

What moves, for an inserted or removed physical row or column:

- **The table's own formulas** (the formula owner whose `formula_owner` is
  the table drawable): every AST of every formula cell the owner records
  (inline and tiled records), in its `TST.TableDataList` entry. Local cell
  references (`AST_row`/`AST_column`, relative to the host or absolute),
  colon tracts (relative and absolute ranges, both endpoints), nested
  thunks. A relative reference into another table compensates for its
  host's move. A footer formula whose range ends on the last body row grows
  when a body row is inserted after it and shrinks when that row goes (the
  host editor's `FooterRangeInsertion` rule).
- **The owner's dependency records**: record hosts, local precedent edges,
  owner-qualified edges into the table, the footer edge adjustments;
  dependency-record tiles (a record that would leave its 128-row or
  32-column tile is refused).
- **Range dependencies and range-precedent tiles**, **UUID-reference
  coordinate sets**, and the merge-proxy owners' ranges into the table.
- **The table extents** of the owner's spanning dependencies
  (`calc::rewrite_extents`): set from the new geometry when they describe
  the old one, else moved by one on the edited axis as the host editor
  moves them.
- **Beyond the host editor, which refuses these:** other tables' formulas
  that read the edited table (cross-table and cross-sheet), with their
  owner-qualified edges, range dependencies and range tiles; the reader's
  hosts stay. Packed edge arrays (this crate's formula writer packs them;
  native Numbers and the host editor do not) are rewritten in their own
  encoding. Whole-row and whole-column references keep their sentinel (the
  host editor would shift `0x7fff_ffff` to `0x8000_0000`).

A removal is planned twice: before the removed axis is cleared (formula
hosts on it are skipped, so a surviving formula that reads the axis is
refused before anything is cleared), and after.

### 2.2 Typed refusals

`table::topology::Unsupported` gains two reasons; two keep their names with
narrower meanings:

- `DeletedReference` (new): a surviving formula reads a cell of the removed
  row or column, or a range starts or ends there. Native Numbers writes
  `#REF!`; the host editor refuses (its `DependencyMutation::coordinate`),
  and so does this transaction: no corpus or reference output shows the
  native reference-error encoding to match, and a reference-error node
  without its UID tract would lose the formula's meaning.
- `SharedFormulas` (new): one formula-list entry shared by cells whose
  rewritten formulas would differ.
- `Formulas`: volatile, spanning, whole-owner, error or spill dependency
  state in any owner (the host editor's `validate_shiftable_formula_owner`,
  applied to every owner as it does), a category reference whose host
  moves, a packed coordinate, a record that would leave its tile.
- `ReferencedByOtherOwners`: a chart or another non-table owner reads the
  table, or a formula reads it through legacy UUID ranges or UUID
  references.

The unified layer maps all four to `unsupported_capability`; the operation
gap (`add row`, context "formulas the coordinate shift refuses, merges")
names them.

### 2.3 Shared formula entries

The host editor rewrites an entry in place when its reference count is one,
or when every cell holding it is rewritten to the same formula and no other
entry already holds that formula; otherwise it moves each cell onto a new
or an equal entry. This transaction rewrites in place whenever every cell
holding the entry gets the same formula (an equal entry elsewhere is
allowed, so two entries may hold one formula), and refuses
(`SharedFormulas`) when they would diverge. §2.5 is why.

### 2.4 The differential test (`litchi-iwa/tests/numbers_topology_formula_differential.rs`)

- **Oracle.** `NumbersEditor::__edit_table_topology_unverified` (new,
  `#[doc(hidden)]`, `topology_oracle.rs`) stages the host editor's row or
  column edit without the post-edit appearance reread of the public
  methods: 15 native corpus files fail that reread for reasons unrelated
  to the edit, which hid the shift.
- **Cases.** Every table of the 40 Numbers corpus packages, plus synthetic
  leaf-authored workbooks (same-table relative, absolute, mixed and range
  formulas; formulas shared by the formula writer; cross-table and
  cross-sheet readers of another table and of the edited one; a footer
  row's sum), each edited by inserting and removing a body row and a body
  column at the start, the middle and the end. For a removal the host
  editor gets the source with the removed axis already cleared by the
  leaf's cell writer (the leaf clears the same way first; the host
  editor's own clear leaves range dependencies behind and then refuses).
- **Compared.** Every calculation-engine payload (`4008`, `4009`, `4010`)
  and every formula list (`6005` of list type 3), by identifier, byte for
  byte; and every formula's rendered text in every table.
- **Packed edges.** The synthetic workbooks are compared with the host
  editor in an unpacked form (the test rewrites their packed edge arrays as
  unpacked fields); the leaf's result on the packed original is then
  compared with its result on the unpacked form, unpacked the same way.
- **Independent oracle.** Where the host editor cannot serve (incoming
  references, its defect below), the leaf's formula texts are checked
  against the source's texts with every A1 reference into the edited table
  shifted by the edit (`shift_text`, `shift_prefixed`).

Result (the committed test asserts the counts):

| Outcome | Cases |
|---|---:|
| Tables with formulas: host and leaf publish, calculation engine and formula lists byte-equal, texts equal | 48 |
| Tables without formulas: byte-equal | 318 |
| Packed originals: leaf result equal to the unpacked form's | 53 |
| Other tables read the edited one: host refuses, leaf publishes, texts match the oracle | 8 |
| Host formulas wrong (§2.5), leaf matches the oracle | 2 |
| Same formulas, formula list interned differently (§2.3) | 1 |
| Both refuse (merges, volatile/spanning state, charts, deleted references, created-table storage layout, cell-writer clears) | rest |
| Host publishes, leaf refuses | 21, every one in the leaf cell writer's clear of the removed axis (`HeaderNameIndex`, `FormulaCache`, the created workbook's reference budget) or a commented axis (`Comments`); none in the shift |

Every corpus table with formulas the host editor edits is matched; the host
editor refuses the other corpus formula tables too (their volatile or
spanning dependency state, sub-owners and charts that read them).

### 2.5 A host-editor defect the test found

On a table whose formula writer interned `=B2`/`=B3` in shared entries,
inserting a row between them made the host editor write `E8 = B1` where
`B3` was expected (and `=A2` for a column). It moves a cell whose shifted
formula equals another entry onto that entry, then rewrites that entry in
place from its stale reference count for the entry's own cells, so the
moved cell takes their formula. The leaf rewrites each entry in place
(§2.3); the test records the two cases and checks the leaf against the
oracle.

### 2.6 Tests

- `litchi-numbers/tests/topology.rs::formula_references_follow_row_and_column_edits`:
  a created workbook's formulas follow a row and a column insert, the
  inverse restores the bytes, removing them restores the texts, removing a
  read row is refused (`DeletedReference`).
- `litchi-ops/tests/numbers_unified.rs::rows_and_columns_move_formula_references`:
  the unified `add` of a row and a column on a table with a formula, the
  inverse, and the typed refusal.
- The corpus round trip (`corpus_tables_round_trip_or_refuse_typed`) and
  the unified corpus tests pass unchanged.

## 3. Numbers cell comments (request 19, `litchi-numbers`)

### 3.1 What native packages look like

Native Numbers gives every table its own comment list (an empty list too),
in a member of its own (`Index/Tables/DataList-<id>.iwa`, or the shared
`DataList.iwa`); comment storages beside the list; tiles in `Tile*.iwa`;
authors in `AnnotationAuthorStorage.iwa`. It keeps a comment list's
references in the ArchiveInfo aggregate without a FieldInfo mirror, and
names many components by one preferred locator (`Tables/DataList`) with
distinct effective ones (`Tables/DataList-905013`), the first of which is
also a member.

### 3.2 What changed

- **Census** (`comments.rs::validate_comment_message_metadata`): a mutation
  no longer requires the FieldInfo mirror. The aggregate is still checked
  against the payload's references in both modes; a present mirror is
  still checked in full. This unblocked every edit and clear of a native
  comment (6 → 0 refusals of record 0314's metadata rule).
- **Locators** (`comments_metadata.rs::find_physical_descriptor`): the
  effective locator names the member, as `structure::metadata` already
  resolves it. The former rule (refuse when the preferred member also
  exists) refused every native package in the registry.
- **Per-table census** (`comments_create.rs::creation_graph`): creation
  owns only the target table's list (the table model's
  `commentStorageTable`); other lists are proven well formed by the census
  but no longer refuse (31 → 0 refusals of record 0314's census rule).
- **Several members** (`comments_create.rs::Archives`): the list's member
  takes the new storage; the tile's member, and an existing author
  storage's member for a new author, are rewritten too, each archive once;
  objects are re-sorted only in a member that gained objects. The storage's
  author edge across members is registered as an external reference unless
  one already covers it (`MetadataRegistry::external_edge_present`, new).
  An existing author in another member is not UUID-checked against the
  list's member (its edge is).
- **Representation**: an appended list entry gains a FieldInfo only when
  the list already mirrors its references (or has none yet).
- **Clear of a UUID-registered storage** (one this crate created): the
  registration is removed with it (`prepare_metadata_uuid_removal_entry`,
  a removal-only metadata batch); a storage metadata names otherwise still
  refuses.
- **`Package::table_cell_comments(sheet, table)`** (new): a table's
  comments in row-major order. A positional selection whose comment list
  is empty costs one read of the table model's data store and list fields
  (`comment_list_populated`).

### 3.3 Corpus (`litchi-numbers/tests/comment_corpus.rs`)

Every corpus table (35) takes a comment on a stored cell, reopens with it,
gives it back through the inverse byte for byte, takes a second edit on the
published package and a clear, and the clear's inverse restores the
comment; the 9 native comments are rewritten and cleared the same way;
`table_cell_comments` lists exactly the cells that read a comment. A
created workbook and a table added to it take comments too.

Two leaf tests changed with the intended behaviour:
`table_cell_comment_reply.rs` (a FieldInfo-free reply graph is now edited,
exactly reversibly: `graphs_without_field_info_mirrors_are_edited`) and
`table_data_list_reader_integration.rs` (the clear of a UUID-registered
storage now publishes and inverts exactly).

### 3.4 Not done

- **A table without a comment list.** Creation still adds the new list
  beside the table model, in the model's member; native packages would
  give it a member of its own (with its component descriptor and the
  model's external reference). Every corpus table has a list, and created
  workbooks and added tables take comments, so this path is reached only by
  packages from other producers.
- **A comment on an empty cell** materialises the cell where the tile
  stores its row (7 corpus cases) and is refused where the tile writer
  would need a sparse row (`NeedSparse`, 29 cases); the row-header cell
  counts are not rewritten for a new comment-only cell (pre-existing
  behaviour of the same-member writer).
- **Several authors**: creation still refuses a package with more than one
  author.

### 3.5 Unified (`litchi-ops`, `format/iwork/numbers.rs`)

- `comment` on a cell is read and set: `set` with text creates or replaces,
  `null` clears (`set_table_cell_comment`, `clear_table_cell_comment`).
  Each table's comments are read once, on its first cell item.
- `add` of `comment` (with `text`) under a cell creates one (refused,
  `invalid_argument`, when the cell already has one); `remove` of
  `<cell>/comment` clears it.
- The `comment` property gap is gone; capability and registry goldens are
  re-blessed in both feature sets, and the OfficeCLI delta table
  regenerated.
- Test: `numbers_unified.rs::cell_comments_are_read_set_added_and_removed`
  (a native comment replaced, cleared and removed; `basic.numbers` and a
  created workbook: add, inverse, refused second add, edit, remove).

## 4. Measurements

Release builds of `-p litchi-ops --features numbers --example
numbers_unified_bench` at `58f4a443c1` (base) and at the final tree
(after), `taskset -c 7`, `LITCHI_NUMBERS_BENCH_SAMPLES=8`, processes
B A A B (`results/0318/measure.sh`). Ratio is unified/leaf; the bar is 5 %
or 20 µs, whichever is larger. Every pair's outputs are byte-identical
before timing.

| Scenario | Base 1 / 4 | After 2 / 3 | Leaf / unified µs (after 2) | Bar |
|---|---:|---:|---:|---|
| `row_insert_formulas` (new) | — | 1.011 / 1.011 | 370.7 / 374.9 | **met** |
| `comment_set` (new) | — | 1.003 / 1.002 | 4320.6 / 4334.4 | **met** |
| `open_get` | 1.060 / 1.043 | 1.063 / 1.069 | 98.1 / 104.2 | met (+6 µs; the 20 µs arm) |
| `batch_1000` | 1.301 / 1.314 | 1.320 / 1.313 | 960.6 / 1267.9 | miss (pre-existing, 0314) |
| the 17 other scenarios | 0.975–1.034 | 0.976–1.034 | — | met |

- `open_get` cost: the unified cell item now reads its table's comments;
  for an empty list (every table without comments) that is
  `comment_list_populated`, about 1–3 µs over base. The first version
  resolved the table fully and missed (1.29); the fast path fixed it.
- Leaf cost of the shift on tables without formulas: the leaf topology legs
  moved by about 1–2 % (`row_insert` 521 → 529–531 µs, `row_remove`
  1023–1024 → 1037–1040 µs, `col_insert` 486–488 → 497–498 µs,
  `col_remove` 1184–1187 → 1198–1202 µs): the shift enumerates every
  formula owner of the package twice per operation (the check before the
  clear and the plan). Under the 5 % regression rule; not optimized.

## 5. Keynote and Pages table creation: not started

The retired host editor's `KeynoteEditor::add_slide_table` copies an
existing table's structure as a template (`table_template_from_catalog`),
builds a fresh TST graph (`create_empty_table_graph_in_package`: model,
data store, tiles, every data list, styles, UID map, CalculationEngine
owner), sets geometry and uniform dimensions, and links the drawable into
the slide. In the leaf world the pieces exist separately: Numbers
`table::lifecycle::add` (record 0301) builds a fresh table from the first
table's structure over `litchi-numbers`' structural toolkit; Keynote has
drawable creation (`slide_movie_creation`, `slide_audio_creation`) and the
shared TST cell writer (`litchi_numbers_wire::table_write`, records 0309,
0312); Pages has the 0312 structural toolkit (`package/structure.rs`). A
port needs the table-graph copy ported onto Keynote's and Pages' structure
toolkits (a decks' or documents' first table, or a built-in template where
none exists), the drawable insertion (slide drawable lists; Pages body
attachments and their anchors), and the unified `add` of `tbl`. Not
started in this record; the capability gaps stay.

Keynote speaker-notes run styles (records 0312 §3.2): not started.

## 6. Gates (final tree, exit codes)

Build environment per the rules (`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`,
one cargo command at a time, the memory-capped cargo shim, test executables
pruned between chunks).

| Gate | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-numbers -p litchi-iwa -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features numbers\|pages\|keynote] --all-targets -- -D warnings` (4 runs) | 0 |
| `cargo clippy -p litchi-numbers -p litchi-iwa --all-targets -- -D warnings` (default features) | 0 |
| `cargo test -p litchi-numbers --all-features` (1076); `cargo test -p litchi-iwa --all-features` (1783) | 0 |
| `cargo test -p litchi-ops --all-features`, chunked: `--lib` 922, `--doc` 210, 139 test targets in 4 chunks (1258) | 0 |
| `cargo test -p litchi-ops --no-default-features --features` every feature but `encryption` `--test capabilities --test registry` (re-blessed in `0f87ea6dc7`, both sets) | 0 |
| `cargo test -p litchi-ops --all-features --test crud_matrix --test corpus_fidelity` (not re-blessed: unchanged) | 0 |
| `cargo test -p litchi --all-features` (230) | 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus`; `python3 tools/officecli_corpus.py --check` | 0 |
| `python3 tools/officecli_capability_delta.py --officecli …/3rdparty/OfficeCLI` (regenerated in `0f87ea6dc7`) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers -p litchi-iwa -p litchi-ops -p litchi` | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify`; `python3 tools/check_iwork_leaf_public_api.py` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir target/glossary` (40 profiles, 2,306 bans, 228 families) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (279) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

## 7. Not claimed

- Native-app acceptance: this host has no macOS. In particular, that a
  shifted formula, a cross-member comment or a formula list holding two
  equal entries opens in Numbers.
- That the shift matches native Numbers where the host editor refuses
  (incoming references, packed edges): those cases are checked against the
  text oracle and by reopening, not against a native or reference output.
- That `#REF!` semantics are implemented: deleted references are refused.
