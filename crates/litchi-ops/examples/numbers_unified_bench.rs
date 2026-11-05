//! Paired leaf and unified Numbers timings (records 0297 and 0301).
//!
//! Scenarios, each on the same input bytes for both legs:
//!
//! - `open_get`: open from bytes and read one cell (leaf
//!   `Package::from_bytes` + the projection's typed cell; unified
//!   `open_bytes` + `get`);
//! - `set`: one cell value (leaf `edit_table_cells` + one change + commit;
//!   unified `File::set`);
//! - `batch_N`: N distinct cell values (leaf: one `edit_table_cells` with N
//!   changes; unified: one atomic `File::batch` of N `set` items, coalesced
//!   into one leaf commit);
//! - record 0314: `batch_N_owned` (the batch items copied outside the
//!   clock) and `batch_N_parsed` (one `File::edit` of the same assignments,
//!   each path parsed once outside the clock);
//! - `rename_sheet`: leaf `edit_names` + `rename_sheet`; unified `set` of the
//!   sheet `name`;
//! - `merge_add`: leaf `edit_table_merges` + `merge`; unified `add` of a
//!   `mergeCell`;
//! - `sheet_move`: leaf `edit_sheet_order` + `move_sheet`; unified
//!   `move_to(.., Placement::Start)` (on the two-sheet fixture).
//! - record 0301: `sheet_add` (leaf `edit_sheet_lifecycle().add`; unified
//!   `add("/", sheet)`), `sheet_remove` (an empty sheet), `table_add` (leaf
//!   `edit_table_lifecycle(..).add`; unified `add("/sheet[1]", table)`),
//!   `table_remove` (the grid's table) and `table_move` (leaf
//!   `edit_table_order`; unified in-sheet `move_to`), on the grid fixture.
//! - record 0309: `row_insert` and `col_insert` (leaf `edit_table_topology`
//!   with `insert_row`/`insert_column` before physical row 51 / column 6;
//!   unified `add(table, row|col)` at that position), `row_remove` and
//!   `col_remove` (the same row or column), and `cell_font` (leaf
//!   `edit_table_cell_text_style` + bold on A3; unified `set` of
//!   `font.bold`), on the grid fixture.
//! - record 0347: `sheet_swap` (three sheets: leaf `move_sheet` 0 → 2, then
//!   1 → 0; unified `swap` of sheets 1 and 3), `table_swap` (three tables of
//!   one sheet: two `edit_table_order` moves; unified `swap`) and
//!   `table_swap_sheets` (one table in each of two sheets: two
//!   `edit_table_relocation`s and two `edit_table_order` moves; unified
//!   `swap`), on packages grown from the grid.
//! - record 0318: `row_insert_formulas` (the `row_insert` pair on a
//!   created workbook whose table holds formulas, so the leaf shifts their
//!   coordinates; inserted before physical row 3) and `comment_set` (leaf
//!   `set_table_cell_comment` on a stored cell of `basic.numbers`; unified
//!   `set` of the cell's `comment`).
//! - record 0329: `row_remove_ref` (the row-3 removal on the created
//!   workbook with formulas, which every formula reads: `#REF!`, a shrunk
//!   range; leaf `remove_row`, unified `remove` of the row) and
//!   `comment_empty` (a comment on the last, unstored cell of
//!   `basic.numbers`).
//!
//! Legs run in ABBA order per sample after a warm-up; outputs of every
//! write scenario are compared byte for byte before timing (the merge by
//! its regions: a leaf merge draws fresh object identifiers). Path and
//! payload building and verification are untimed; neither leg serializes
//! inside the clock.
//!
//! ```text
//! cargo build --release -p litchi-ops --features numbers --example numbers_unified_bench
//! taskset -c 7 target/release/examples/numbers_unified_bench > numbers.json
//! ```
//!
//! `LITCHI_NUMBERS_BENCH_SAMPLES` sets the ABBA rounds (default 5; each
//! round times each leg twice). `LITCHI_NUMBERS_BENCH_ONLY`, a
//! comma-separated list of scenario names, runs only those scenarios.

#[cfg(feature = "numbers")]
mod enabled {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_numbers::cell::text_style::CharacterProperties;
    use litchi_numbers::formula::Expression;
    use litchi_numbers::table::At;
    use litchi_numbers::table::cells::{Change, Input};
    use litchi_numbers::table::merge::Region;
    use litchi_numbers::table::topology::{
        ColumnDeletion, ColumnInsertion, RowDeletion, RowInsertion,
    };
    use litchi_numbers::{Package, SheetSelector, TableSelector};
    use litchi_ops::{BatchItem, BatchMode, File, NewNode, Placement, Props, props};
    use serde::Serialize;

    #[derive(Serialize)]
    struct Report {
        schema: &'static str,
        rounds: usize,
        order: &'static str,
        fixtures: Vec<Fixture>,
        scenarios: Vec<Scenario>,
    }

    #[derive(Serialize)]
    struct Fixture {
        name: &'static str,
        bytes: usize,
    }

    #[derive(Serialize)]
    struct Scenario {
        name: String,
        fixture: &'static str,
        outputs_identical: Option<bool>,
        leaf_ns: Vec<u128>,
        unified_ns: Vec<u128>,
        leaf_median_ns: u128,
        unified_median_ns: u128,
        ratio: f64,
    }

    /// One 101 x 11 table: the corpus has no table of 1000 cells
    /// (`test-data/iwork/synthetic/numbers/README.md`).
    const GRID: &str = "../synthetic/numbers/bench-grid.numbers";
    /// A Numbers-authored package with two sheets, for the sheet move.
    const SHEETS: &str = "table-relocation-native.numbers";
    /// The formula scenario's package (the synthetic grid has no formula
    /// owner).
    const FORMULAS: &str = "basic.numbers";
    const BATCH: usize = 1000;

    fn fixture(name: &str) -> Vec<u8> {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/numbers")
            .join(name);
        std::fs::read(&path).unwrap_or_else(|error| panic!("{}: {error}", path.display()))
    }

    fn leaf_bytes(package: &Package) -> Vec<u8> {
        let mut bytes = Vec::new();
        package.write_to(&mut bytes).expect("leaf bytes");
        bytes
    }

    fn median(values: &[u128]) -> u128 {
        let mut sorted = values.to_vec();
        sorted.sort_unstable();
        sorted[sorted.len() / 2]
    }

    fn time<T>(operation: impl FnOnce() -> T) -> u128 {
        let start = Instant::now();
        black_box(operation());
        start.elapsed().as_nanos()
    }

    fn selected(scenario: &str) -> bool {
        std::env::var("LITCHI_NUMBERS_BENCH_ONLY").map_or(true, |only| {
            only.split(',').any(|wanted| wanted == scenario)
        })
    }

    /// Run `leaf` and `unified` in ABBA order `rounds` times.
    fn abba(
        name: impl Into<String>,
        fixture: &'static str,
        rounds: usize,
        outputs_identical: Option<bool>,
        mut leaf: impl FnMut() -> u128,
        mut unified: impl FnMut() -> u128,
    ) -> Option<Scenario> {
        let name = name.into();
        if !selected(&name) {
            return None;
        }
        leaf();
        unified();
        let (mut leaf_ns, mut unified_ns) = (Vec::new(), Vec::new());
        for _ in 0..rounds {
            leaf_ns.push(leaf());
            unified_ns.push(unified());
            unified_ns.push(unified());
            leaf_ns.push(leaf());
        }
        let (leaf_median_ns, unified_median_ns) = (median(&leaf_ns), median(&unified_ns));
        Some(Scenario {
            name,
            fixture,
            outputs_identical,
            ratio: unified_median_ns as f64 / leaf_median_ns as f64,
            leaf_ns,
            unified_ns,
            leaf_median_ns,
            unified_median_ns,
        })
    }

    /// `package` with a small table named `name` appended to `sheet`.
    fn package_with_table(package: &Package, sheet: usize, name: &str) -> Package {
        package
            .edit_table_lifecycle(sheet)
            .expect("edit")
            .add(name, 4, 3)
            .expect("add")
            .commit()
            .expect("commit")
            .into_package()
    }

    fn table() -> (SheetSelector<'static>, TableSelector<'static>) {
        (SheetSelector::index(0), TableSelector::index(0))
    }

    /// The `BATCH` cells the batch scenario writes: rows 2.. of columns B..K.
    fn batch_cells() -> Vec<(At, String, f64)> {
        (0..BATCH)
            .map(|index| {
                let row = (index / 10 + 1) as u32;
                let column = (index % 10 + 1) as u32;
                let position = At::new(row, column);
                let a1 = format!("{}{}", char::from(b'A' + column as u8), row + 1);
                (position, a1, index as f64)
            })
            .collect()
    }

    #[inline(never)]
    fn leaf_batch(package: &Package, cells: &[(At, String, f64)]) -> Package {
        let (sheet, table) = table();
        let changes = cells.iter().map(|(position, _, value)| {
            Change::set(*position, Input::number(*value).expect("finite"))
        });
        package
            .edit_table_cells(sheet, table)
            .expect("edit")
            .extend(changes)
            .expect("stage")
            .commit()
            .expect("commit")
            .into_package()
    }

    /// Kept out of line so a profiler can select it
    /// (`--toggle-collect=*unified_batch*`).
    #[inline(never)]
    fn unified_batch(file: &File, items: &[BatchItem<'static>]) -> File {
        file.batch(items.to_vec(), BatchMode::Atomic)
            .expect("batch")
            .file()
            .clone()
    }

    /// Print each corpus table's dimensions (`LITCHI_NUMBERS_BENCH_SURVEY`).
    fn survey() {
        let directory =
            std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers");
        let mut names: Vec<_> = std::fs::read_dir(&directory)
            .expect("corpus")
            .filter_map(|entry| entry.ok().map(|entry| entry.path()))
            .filter(|path| path.extension().is_some_and(|ext| ext == "numbers"))
            .collect();
        names.sort();
        for path in names {
            let Ok(package) = Package::open(&path) else {
                continue;
            };
            let workbook = package.workbook_ref();
            for (sheet_index, sheet) in workbook.sheets().iter().enumerate() {
                for (table_index, table) in sheet.tables().enumerate() {
                    eprintln!(
                        "{} {sheet_index}/{table_index} {}x{}",
                        path.display(),
                        table.row_count(),
                        table.column_count()
                    );
                }
            }
        }
    }

    pub(crate) fn main() {
        if std::env::var_os("LITCHI_NUMBERS_BENCH_SURVEY").is_some() {
            survey();
            return;
        }
        let rounds = std::env::var("LITCHI_NUMBERS_BENCH_SAMPLES")
            .ok()
            .and_then(|value| value.parse().ok())
            .unwrap_or(5);
        let grid = fixture(GRID);
        let sheets = fixture(SHEETS);
        let leaf = Package::from_bytes(&grid).expect("leaf open");
        let file = litchi_ops::open_bytes(grid.clone()).expect("unified open");
        let mut scenarios = Vec::new();

        scenarios.extend(abba(
            "open_get",
            GRID,
            rounds,
            None,
            || {
                time(|| {
                    let package = Package::from_bytes(&grid).expect("open");
                    let workbook = package.workbook_ref();
                    let table = workbook.sheets()[0].tables().next().expect("table");
                    table.get_a1("A3").expect("cell").cloned()
                })
            },
            || {
                time(|| {
                    let file = litchi_ops::open_bytes(grid.clone()).expect("open");
                    let node = file.get("/sheet[1]/table[1]/A3").expect("get");
                    node.view()
                        .text()
                        .map(|text| text.map(|text| text.into_owned()))
                })
            },
        ));

        // One cell value.
        let leaf_set = |package: &Package| {
            let (sheet, table) = table();
            package
                .edit_table_cells(sheet, table)
                .expect("edit")
                .extend([Change::set(
                    At::new(2, 1),
                    Input::number(42.0).expect("finite"),
                )])
                .expect("stage")
                .commit()
                .expect("commit")
                .into_package()
        };
        let unified_set = |file: &File| {
            file.set("/sheet[1]/table[1]/B3", props! { "value" => 42 })
                .expect("set")
        };
        let same = leaf_bytes(&leaf_set(&leaf)) == unified_set(&file).to_bytes().expect("bytes");
        scenarios.extend(abba(
            "set",
            GRID,
            rounds,
            Some(same),
            || time(|| leaf_set(&leaf)),
            || time(|| unified_set(&file)),
        ));

        // `BATCH` cell values.
        let cells = batch_cells();
        let items: Vec<BatchItem<'static>> = cells
            .iter()
            .map(|(_, a1, value)| {
                let mut item = BatchItem::new("set");
                item.path = Some(format!("/sheet[1]/table[1]/{a1}").into());
                item.props = Some(Props::new().with("value", *value));
                item
            })
            .collect();
        let same = leaf_bytes(&leaf_batch(&leaf, &cells))
            == unified_batch(&file, &items).to_bytes().expect("bytes");
        scenarios.extend(abba(
            format!("batch_{BATCH}"),
            GRID,
            rounds,
            Some(same),
            || time(|| leaf_batch(&leaf, &cells)),
            || time(|| unified_batch(&file, &items)),
        ));
        // Record 0314: the batch with its items copied outside the clock
        // (`File::batch` takes them by value), and the same assignments in
        // one `File::edit` with each path parsed once outside the clock.
        scenarios.extend(abba(
            format!("batch_{BATCH}_owned"),
            GRID,
            rounds,
            Some(same),
            || time(|| leaf_batch(&leaf, &cells)),
            || {
                let (file, owned) = (&file, items.to_vec());
                time(move || {
                    file.batch(owned, BatchMode::Atomic)
                        .expect("batch")
                        .into_file()
                })
            },
        ));
        let targets: Vec<litchi_ops::TargetArg<'_>> = items
            .iter()
            .map(|item| {
                litchi_ops::TargetArg::parse(item.path.as_deref().expect("path")).expect("target")
            })
            .collect();
        let unified_parsed = |file: &File| {
            file.edit(|tx| {
                for (target, (_, _, value)) in targets.iter().zip(&cells) {
                    tx.set(target, props! { "value" => *value })?;
                }
                Ok(())
            })
            .expect("edit")
            .into_file()
        };
        let same_parsed = leaf_bytes(&leaf_batch(&leaf, &cells))
            == unified_parsed(&file).to_bytes().expect("bytes");
        scenarios.extend(abba(
            format!("batch_{BATCH}_parsed"),
            GRID,
            rounds,
            Some(same_parsed),
            || time(|| leaf_batch(&leaf, &cells)),
            || time(|| unified_parsed(&file)),
        ));

        let formulas = fixture(FORMULAS);
        let basic = Package::from_bytes(&formulas).expect("leaf open");
        let basic_file = litchi_ops::open_bytes(formulas.clone()).expect("unified open");
        // One formula.
        let leaf_formula = |package: &Package| {
            let (sheet, table) = table();
            let expression = Expression::parse("B3*2").expect("parse");
            package
                .edit_table_cells(sheet, table)
                .expect("edit")
                .extend([Change::set_formula(At::new(2, 3), expression)])
                .expect("stage")
                .commit()
                .expect("commit")
                .into_package()
        };
        let unified_formula = |file: &File| {
            file.set("/sheet[1]/table[1]/D3", props! { "formula" => "B3*2" })
                .expect("formula")
        };
        let same = leaf_bytes(&leaf_formula(&basic))
            == unified_formula(&basic_file).to_bytes().expect("bytes");
        scenarios.extend(abba(
            "formula",
            FORMULAS,
            rounds,
            Some(same),
            || time(|| leaf_formula(&basic)),
            || time(|| unified_formula(&basic_file)),
        ));

        // Sheet rename.
        let leaf_rename = |package: &Package| {
            package
                .edit_names()
                .rename_sheet(SheetSelector::index(0), "Renamed")
                .expect("rename")
                .commit()
                .expect("commit")
                .into_package()
        };
        let unified_rename = |file: &File| {
            file.set("/sheet[1]", props! { "name" => "Renamed" })
                .expect("rename")
        };
        let same =
            leaf_bytes(&leaf_rename(&leaf)) == unified_rename(&file).to_bytes().expect("bytes");
        scenarios.extend(abba(
            "rename_sheet",
            GRID,
            rounds,
            Some(same),
            || time(|| leaf_rename(&leaf)),
            || time(|| unified_rename(&file)),
        ));

        // Merge add.
        let region = Region::new(40, 2, 2, 2).expect("region");
        let leaf_merge = |package: &Package| {
            let (sheet, table) = table();
            let mut edit = package.edit_table_merges(sheet, table).expect("merges");
            edit.merge(region).expect("merge");
            edit.commit().expect("commit").into_package()
        };
        let unified_merge = |file: &File| {
            file.add(
                "/sheet[1]/table[1]",
                NewNode::new("mergeCell").prop("ref", "C41:D42"),
            )
            .expect("merge")
            .into_file()
        };
        // A merge draws fresh object identifiers, so two leaf merges already
        // differ byte for byte; compare the merged regions instead.
        let regions = |bytes: Vec<u8>| {
            let (sheet, table) = table();
            Package::from_bytes(&bytes)
                .expect("reopen")
                .table_merges(sheet, table)
                .expect("merges")
                .to_vec()
        };
        let same = regions(leaf_bytes(&leaf_merge(&leaf)))
            == regions(unified_merge(&file).to_bytes().expect("bytes"));
        scenarios.extend(abba(
            "merge_add",
            GRID,
            rounds,
            Some(same),
            || time(|| leaf_merge(&leaf)),
            || time(|| unified_merge(&file)),
        ));

        // Sheet move.
        let native = Package::from_bytes(&sheets).expect("leaf open");
        let native_file = litchi_ops::open_bytes(sheets.clone()).expect("unified open");
        let leaf_move = |package: &Package| {
            package
                .edit_sheet_order()
                .move_sheet(SheetSelector::index(1), 0)
                .expect("move")
                .commit()
                .expect("commit")
                .into_package()
        };
        let unified_move = |file: &File| file.move_to("/sheet[2]", Placement::Start).expect("move");
        let same = leaf_bytes(&leaf_move(&native))
            == unified_move(&native_file).to_bytes().expect("bytes");
        scenarios.extend(abba(
            "sheet_move",
            SHEETS,
            rounds,
            Some(same),
            || time(|| leaf_move(&native)),
            || time(|| unified_move(&native_file)),
        ));

        // Record 0301: structural operations on the grid.
        let leaf_sheet_add = |package: &Package| {
            package
                .edit_sheet_lifecycle()
                .add("Added")
                .expect("add")
                .commit()
                .expect("commit")
                .into_package()
        };
        let unified_sheet_add = |file: &File| {
            file.add("/", NewNode::new("sheet").prop("name", "Added"))
                .expect("add sheet")
                .into_file()
        };
        let same = leaf_bytes(&leaf_sheet_add(&leaf))
            == unified_sheet_add(&file).to_bytes().expect("bytes");
        scenarios.extend(abba(
            "sheet_add",
            GRID,
            rounds,
            Some(same),
            || time(|| leaf_sheet_add(&leaf)),
            || time(|| unified_sheet_add(&file)),
        ));
        let two = leaf_sheet_add(&leaf);
        let two_file = litchi_ops::open_bytes(leaf_bytes(&two)).expect("open two");
        let leaf_sheet_remove = |package: &Package| {
            package
                .edit_sheet_lifecycle()
                .remove(1_usize)
                .expect("remove")
                .commit()
                .expect("commit")
                .into_package()
        };
        let unified_sheet_remove = |file: &File| file.remove("/sheet[2]").expect("remove sheet");
        let same = leaf_bytes(&leaf_sheet_remove(&two))
            == unified_sheet_remove(&two_file).to_bytes().expect("bytes");
        scenarios.extend(abba(
            "sheet_remove",
            GRID,
            rounds,
            Some(same),
            || time(|| leaf_sheet_remove(&two)),
            || time(|| unified_sheet_remove(&two_file)),
        ));
        let leaf_table_add = |package: &Package| {
            package
                .edit_table_lifecycle(0_usize)
                .expect("edit")
                .add("Added", 10, 5)
                .expect("add")
                .commit()
                .expect("commit")
                .into_package()
        };
        let unified_table_add = |file: &File| {
            file.add("/sheet[1]", NewNode::new("table").prop("name", "Added"))
                .expect("add table")
                .into_file()
        };
        let same = leaf_bytes(&leaf_table_add(&leaf))
            == unified_table_add(&file).to_bytes().expect("bytes");
        scenarios.extend(abba(
            "table_add",
            GRID,
            rounds,
            Some(same),
            || time(|| leaf_table_add(&leaf)),
            || time(|| unified_table_add(&file)),
        ));
        let leaf_table_remove = |package: &Package| {
            package
                .edit_table_lifecycle(0_usize)
                .expect("edit")
                .remove(0_usize)
                .expect("remove")
                .commit()
                .expect("commit")
                .into_package()
        };
        let unified_table_remove =
            |file: &File| file.remove("/sheet[1]/table[1]").expect("remove table");
        let same = leaf_bytes(&leaf_table_remove(&leaf))
            == unified_table_remove(&file).to_bytes().expect("bytes");
        scenarios.extend(abba(
            "table_remove",
            GRID,
            rounds,
            Some(same),
            || time(|| leaf_table_remove(&leaf)),
            || time(|| unified_table_remove(&file)),
        ));
        let tables = leaf_table_add(&leaf);
        let tables_file = litchi_ops::open_bytes(leaf_bytes(&tables)).expect("open tables");
        let leaf_table_move = |package: &Package| {
            package
                .edit_table_order(0_usize)
                .expect("edit")
                .move_table(1_usize, 0)
                .expect("move")
                .commit()
                .expect("commit")
                .into_package()
        };
        let unified_table_move = |file: &File| {
            file.move_to("/sheet[1]/table[2]", Placement::Start)
                .expect("move table")
        };
        let same = leaf_bytes(&leaf_table_move(&tables))
            == unified_table_move(&tables_file).to_bytes().expect("bytes");
        scenarios.extend(abba(
            "table_move",
            GRID,
            rounds,
            Some(same),
            || time(|| leaf_table_move(&tables)),
            || time(|| unified_table_move(&tables_file)),
        ));

        // Record 0347: swaps, each the leaf's two moves (four steps across
        // sheets) against one unified `swap`.
        let leaf_add_sheet = |package: &Package, name: &str| {
            package
                .edit_sheet_lifecycle()
                .add(name)
                .expect("add")
                .commit()
                .expect("commit")
                .into_package()
        };
        let leaf_sheet_order = |package: &Package, from: usize, to: usize| {
            package
                .edit_sheet_order()
                .move_sheet(SheetSelector::index(from), to)
                .expect("move")
                .commit()
                .expect("commit")
                .into_package()
        };
        let leaf_table_order = |package: &Package, sheet: usize, from: usize, to: usize| {
            package
                .edit_table_order(sheet)
                .expect("edit")
                .move_table(from, to)
                .expect("move")
                .commit()
                .expect("commit")
                .into_package()
        };
        let three_sheets = leaf_add_sheet(&two, "Third");
        let three_sheets_file =
            litchi_ops::open_bytes(leaf_bytes(&three_sheets)).expect("open three sheets");
        let leaf_sheet_swap =
            |package: &Package| leaf_sheet_order(&leaf_sheet_order(package, 0, 2), 1, 0);
        let unified_sheet_swap =
            |file: &File| file.swap("/sheet[1]", "/sheet[3]").expect("swap sheets");
        let same = leaf_bytes(&leaf_sheet_swap(&three_sheets))
            == unified_sheet_swap(&three_sheets_file)
                .to_bytes()
                .expect("bytes");
        scenarios.extend(abba(
            "sheet_swap",
            GRID,
            rounds,
            Some(same),
            || time(|| leaf_sheet_swap(&three_sheets)),
            || time(|| unified_sheet_swap(&three_sheets_file)),
        ));
        let three_tables = package_with_table(&tables, 0, "Third");
        let three_tables_file =
            litchi_ops::open_bytes(leaf_bytes(&three_tables)).expect("open three tables");
        let leaf_table_swap =
            |package: &Package| leaf_table_order(&leaf_table_order(package, 0, 0, 2), 0, 1, 0);
        let unified_table_swap = |file: &File| {
            file.swap("/sheet[1]/table[1]", "/sheet[1]/table[3]")
                .expect("swap tables")
        };
        let same = leaf_bytes(&leaf_table_swap(&three_tables))
            == unified_table_swap(&three_tables_file)
                .to_bytes()
                .expect("bytes");
        scenarios.extend(abba(
            "table_swap",
            GRID,
            rounds,
            Some(same),
            || time(|| leaf_table_swap(&three_tables)),
            || time(|| unified_table_swap(&three_tables_file)),
        ));
        let split = package_with_table(&two, 1, "Other");
        let split_file = litchi_ops::open_bytes(leaf_bytes(&split)).expect("open split");
        let leaf_relocate = |package: &Package, sheet: usize, to: usize| {
            package
                .edit_table_relocation(
                    SheetSelector::index(sheet),
                    TableSelector::index(0),
                    SheetSelector::index(to),
                )
                .expect("edit")
                .commit()
                .expect("commit")
                .into_package()
        };
        let leaf_cross_swap = |package: &Package| {
            let relocated = leaf_relocate(&leaf_relocate(package, 0, 1), 1, 0);
            leaf_table_order(&leaf_table_order(&relocated, 0, 0, 0), 1, 0, 0)
        };
        let unified_cross_swap = |file: &File| {
            file.swap("/sheet[1]/table[1]", "/sheet[2]/table[1]")
                .expect("swap across sheets")
        };
        let same = leaf_bytes(&leaf_cross_swap(&split))
            == unified_cross_swap(&split_file).to_bytes().expect("bytes");
        scenarios.extend(abba(
            "table_swap_sheets",
            GRID,
            rounds,
            Some(same),
            || time(|| leaf_cross_swap(&split)),
            || time(|| unified_cross_swap(&split_file)),
        ));

        // Record 0309: rows and columns, and cell character properties.
        let (sheet, table_selector) = table();
        let settings = leaf
            .table_header_settings(sheet, table_selector)
            .expect("header settings");
        let (header_rows, header_columns) =
            (settings.header_row_count(), settings.header_column_count());
        const ROW: usize = 50;
        const COLUMN: usize = 5;
        let leaf_row_insert = |package: &Package| {
            let (sheet, table) = table();
            package
                .edit_table_topology(sheet, table)
                .expect("edit")
                .insert_row(RowInsertion::body(ROW - header_rows))
                .expect("insert")
                .commit()
                .expect("commit")
                .into_package()
        };
        let unified_row_insert = |file: &File| {
            file.add(
                "/sheet[1]/table[1]",
                NewNode::new("row").at(Placement::Position(litchi_ops::path::Position::new(ROW))),
            )
            .expect("add row")
            .into_file()
        };
        let same = leaf_bytes(&leaf_row_insert(&leaf))
            == unified_row_insert(&file).to_bytes().expect("bytes");
        scenarios.extend(abba(
            "row_insert",
            GRID,
            rounds,
            Some(same),
            || time(|| leaf_row_insert(&leaf)),
            || time(|| unified_row_insert(&file)),
        ));
        let leaf_row_remove = |package: &Package| {
            let (sheet, table) = table();
            package
                .edit_table_topology(sheet, table)
                .expect("edit")
                .remove_row(RowDeletion::body(ROW - header_rows))
                .expect("remove")
                .commit()
                .expect("commit")
                .into_package()
        };
        let row_path = format!("/sheet[1]/table[1]/row[{}]", ROW + 1);
        let unified_row_remove = |file: &File| file.remove(row_path.as_str()).expect("remove row");
        let same = leaf_bytes(&leaf_row_remove(&leaf))
            == unified_row_remove(&file).to_bytes().expect("bytes");
        scenarios.extend(abba(
            "row_remove",
            GRID,
            rounds,
            Some(same),
            || time(|| leaf_row_remove(&leaf)),
            || time(|| unified_row_remove(&file)),
        ));
        let leaf_col_insert = |package: &Package| {
            let (sheet, table) = table();
            package
                .edit_table_topology(sheet, table)
                .expect("edit")
                .insert_column(ColumnInsertion::body(COLUMN - header_columns))
                .expect("insert")
                .commit()
                .expect("commit")
                .into_package()
        };
        let unified_col_insert = |file: &File| {
            file.add(
                "/sheet[1]/table[1]",
                NewNode::new("col")
                    .at(Placement::Position(litchi_ops::path::Position::new(COLUMN))),
            )
            .expect("add col")
            .into_file()
        };
        let same = leaf_bytes(&leaf_col_insert(&leaf))
            == unified_col_insert(&file).to_bytes().expect("bytes");
        scenarios.extend(abba(
            "col_insert",
            GRID,
            rounds,
            Some(same),
            || time(|| leaf_col_insert(&leaf)),
            || time(|| unified_col_insert(&file)),
        ));
        let leaf_col_remove = |package: &Package| {
            let (sheet, table) = table();
            package
                .edit_table_topology(sheet, table)
                .expect("edit")
                .remove_column(ColumnDeletion::body(COLUMN - header_columns))
                .expect("remove")
                .commit()
                .expect("commit")
                .into_package()
        };
        let col_path = format!("/sheet[1]/table[1]/col[{}]", COLUMN + 1);
        let unified_col_remove = |file: &File| file.remove(col_path.as_str()).expect("remove col");
        let same = leaf_bytes(&leaf_col_remove(&leaf))
            == unified_col_remove(&file).to_bytes().expect("bytes");
        scenarios.extend(abba(
            "col_remove",
            GRID,
            rounds,
            Some(same),
            || time(|| leaf_col_remove(&leaf)),
            || time(|| unified_col_remove(&file)),
        ));
        let leaf_font = |package: &Package| {
            let (sheet, table) = table();
            let mut bold = CharacterProperties::new();
            bold.bold = Some(true);
            package
                .edit_table_cell_text_style(sheet, table)
                .expect("edit")
                .set(At::new(2, 0), bold)
                .expect("stage")
                .commit()
                .expect("commit")
                .into_package()
        };
        let unified_font = |file: &File| {
            file.set("/sheet[1]/table[1]/A3", props! { "font.bold" => true })
                .expect("font")
        };
        let same = leaf_bytes(&leaf_font(&leaf)) == unified_font(&file).to_bytes().expect("bytes");
        scenarios.extend(abba(
            "cell_font",
            GRID,
            rounds,
            Some(same),
            || time(|| leaf_font(&leaf)),
            || time(|| unified_font(&file)),
        ));

        // Record 0318: a row under formulas, and a cell comment.
        let formula_package = {
            let base = Package::new();
            let mut edit = base.edit_table_cells(0, 0).expect("cells");
            for (index, address) in ["B2", "C2", "B3", "C3", "B4", "C4"].into_iter().enumerate() {
                edit = edit
                    .set_a1(address, Input::number(index as f64).expect("finite"))
                    .expect("value");
            }
            for (address, text) in [("D2", "B2+C3"), ("D3", "SUM($B$2:B4)"), ("D4", "B$3*$C3")] {
                edit = edit
                    .set_formula_a1(address, Expression::parse(text).expect("parse"))
                    .expect("formula");
            }
            edit.commit().expect("formulas").into_package()
        };
        let formula_bytes = leaf_bytes(&formula_package);
        let formula_file = litchi_ops::open_bytes(formula_bytes.clone()).expect("open formulas");
        let formula_leaf = Package::from_bytes(&formula_bytes).expect("leaf formulas");
        let leaf_formula_row = |package: &Package| {
            let (sheet, table) = table();
            package
                .edit_table_topology(sheet, table)
                .expect("edit")
                .insert_row(RowInsertion::body(1))
                .expect("insert")
                .commit()
                .expect("commit")
                .into_package()
        };
        let unified_formula_row = |file: &File| {
            file.add(
                "/sheet[1]/table[1]",
                NewNode::new("row").at(Placement::Position(litchi_ops::path::Position::new(2))),
            )
            .expect("add row")
            .into_file()
        };
        let same = leaf_bytes(&leaf_formula_row(&formula_leaf))
            == unified_formula_row(&formula_file)
                .to_bytes()
                .expect("bytes");
        scenarios.extend(abba(
            "row_insert_formulas",
            "created workbook with formulas",
            rounds,
            Some(same),
            || time(|| leaf_formula_row(&formula_leaf)),
            || time(|| unified_formula_row(&formula_file)),
        ));
        let comment_leaf = Package::from_bytes(&formulas).expect("leaf basic");
        let comment_file = litchi_ops::open_bytes(formulas.clone()).expect("open basic");
        let stored = comment_leaf.sheets()[0]
            .tables()
            .next()
            .and_then(|table| table.iter_cells().next())
            .map(|cell| cell.position())
            .expect("a stored cell");
        let stored_path = format!(
            "/sheet[1]/table[1]/{}{}",
            char::from(b'A' + stored.column() as u8),
            stored.row() + 1
        );
        let leaf_comment = |package: &Package| {
            let (sheet, table) = table();
            package
                .set_table_cell_comment(sheet, table, stored, "bench")
                .expect("comment")
                .into_package()
        };
        let unified_comment = |file: &File| {
            file.set(stored_path.as_str(), props! { "comment" => "bench" })
                .expect("comment")
        };
        let same = leaf_bytes(&leaf_comment(&comment_leaf))
            == unified_comment(&comment_file).to_bytes().expect("bytes");
        scenarios.extend(abba(
            "comment_set",
            FORMULAS,
            rounds,
            Some(same),
            || time(|| leaf_comment(&comment_leaf)),
            || time(|| unified_comment(&comment_file)),
        ));

        // Record 0329: a removal formulas read, and a comment on an empty
        // cell in a row its tile does not store.
        let leaf_formula_remove = |package: &Package| {
            let (sheet, table) = table();
            package
                .edit_table_topology(sheet, table)
                .expect("edit")
                .remove_row(RowDeletion::body(1))
                .expect("remove")
                .commit()
                .expect("commit")
                .into_package()
        };
        let unified_formula_remove = |file: &File| {
            file.remove("/sheet[1]/table[1]/row[3]")
                .expect("remove a read row")
        };
        let same = leaf_bytes(&leaf_formula_remove(&formula_leaf))
            == unified_formula_remove(&formula_file)
                .to_bytes()
                .expect("bytes");
        scenarios.extend(abba(
            "row_remove_ref",
            "created workbook with formulas",
            rounds,
            Some(same),
            || time(|| leaf_formula_remove(&formula_leaf)),
            || time(|| unified_formula_remove(&formula_file)),
        ));
        let empty = {
            let table = comment_leaf.sheets()[0].tables().next().expect("table");
            At::new(table.row_count() - 1, table.column_count() - 1)
        };
        let empty_path = format!(
            "/sheet[1]/table[1]/{}{}",
            char::from(b'A' + empty.column() as u8),
            empty.row() + 1
        );
        let leaf_comment_empty = |package: &Package| {
            let (sheet, table) = table();
            package
                .set_table_cell_comment(sheet, table, empty, "bench")
                .expect("comment")
                .into_package()
        };
        let unified_comment_empty = |file: &File| {
            file.set(empty_path.as_str(), props! { "comment" => "bench" })
                .expect("comment")
        };
        let same = leaf_bytes(&leaf_comment_empty(&comment_leaf))
            == unified_comment_empty(&comment_file)
                .to_bytes()
                .expect("bytes");
        scenarios.extend(abba(
            "comment_empty",
            FORMULAS,
            rounds,
            Some(same),
            || time(|| leaf_comment_empty(&comment_leaf)),
            || time(|| unified_comment_empty(&comment_file)),
        ));

        let report = Report {
            schema: "litchi.numbers-unified-bench.v1",
            rounds,
            order: "ABBA",
            fixtures: vec![
                Fixture {
                    name: GRID,
                    bytes: grid.len(),
                },
                Fixture {
                    name: SHEETS,
                    bytes: sheets.len(),
                },
                Fixture {
                    name: FORMULAS,
                    bytes: formulas.len(),
                },
            ],
            scenarios,
        };
        println!("{}", serde_json::to_string(&report).expect("json"));
    }
}

#[cfg(feature = "numbers")]
fn main() {
    enabled::main();
}

#[cfg(not(feature = "numbers"))]
fn main() {
    eprintln!("numbers_unified_bench needs the `numbers` feature");
}
