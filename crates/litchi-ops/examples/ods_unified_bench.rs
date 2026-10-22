//! Paired leaf and unified ODS timings (record 0277).
//!
//! Scenarios, each on the same input bytes for both legs:
//!
//! - `open_get`: open from bytes and read one cell's displayed text (leaf
//!   `Package::from_bytes` + the projection's typed cell; unified
//!   `open_bytes` + `get("/Data/C5")`);
//! - `query`: enumerate every content cell's text (leaf: the projection's
//!   typed runs; unified: `query("cell")`), on an opened handle;
//! - `query_warm` and `query_cold` (record 0280): the same enumeration with
//!   both legs warm, and with both legs building their index on a freshly
//!   opened handle inside the clock;
//! - `query_warm_owned` (record 0315): the warm query against a leaf leg
//!   that expands repeated rows and cells as the walk does and keeps each
//!   content cell as an owned handle (the shared workbook and its
//!   coordinates) while reading its text;
//! - `set`: one cell value (leaf `Package::edit` + one `set_cells` + commit;
//!   unified `File::set`);
//! - `batch_N`: N distinct cell values in one edit (leaf: one `set_cells`
//!   per 4096 changes inside one worksheet edit; unified: one atomic
//!   `File::batch` of N `set` items, coalesced into one leaf commit).
//!
//! Legs run in ABBA order per sample after a warm-up; outputs of every
//! write scenario are compared byte for byte before timing. Fixture
//! construction, path and payload building and verification are untimed.
//!
//! ```text
//! cargo build --release -p litchi-ops --features ods --example ods_unified_bench
//! taskset -c 7 target/release/examples/ods_unified_bench > ods.json
//! ```
//!
//! `LITCHI_ODS_BENCH_SAMPLES` sets the ABBA rounds (default 5; each round
//! times each leg twice). `LITCHI_ODS_BENCH_ROWS` sets the fixture rows
//! (default 1000; ten columns). `LITCHI_ODS_BENCH_ONLY` (record 0287), a
//! comma-separated list of scenario names, runs only those scenarios.

#[cfg(feature = "ods")]
mod enabled {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::Position;
    use litchi_ods::worksheet::{CellChange, MAX_CELL_CHANGES, Selector};
    use litchi_ods::{Builder, Cell, CellValue, Package, Row, Sheet};
    use litchi_ops::{BatchItem, BatchMode, File, Props};
    use serde::Serialize;

    #[derive(Serialize)]
    struct Report {
        schema: &'static str,
        rows: usize,
        columns: usize,
        input_bytes: usize,
        rounds: usize,
        order: &'static str,
        scenarios: Vec<Scenario>,
    }

    #[derive(Serialize)]
    struct Scenario {
        name: String,
        outputs_identical: Option<bool>,
        leaf_ns: Vec<u128>,
        unified_ns: Vec<u128>,
        leaf_median_ns: u128,
        unified_median_ns: u128,
        ratio: f64,
    }

    const COLUMNS: usize = 10;

    fn fixture(rows: usize) -> Vec<u8> {
        let mut sheet = Sheet::new("Data").expect("sheet");
        for row in 0..rows {
            let mut cells = Row::new();
            for column in 0..COLUMNS {
                let cell = if column % 2 == 0 {
                    let value = (row * COLUMNS + column) as f64;
                    Cell::new(CellValue::Number(value), value.to_string())
                } else {
                    let text = format!("r{row}c{column}");
                    Cell::new(CellValue::Text(text.clone()), text)
                };
                cells.push_cell(cell).expect("cell");
            }
            sheet.push_row(cells).expect("row");
        }
        let mut builder = Builder::new();
        builder.add_sheet(sheet).expect("add sheet");
        builder.build().expect("build").to_bytes().expect("bytes")
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

    /// Run `leaf` and `unified` in ABBA order `rounds` times.
    fn abba(
        name: impl Into<String>,
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
            outputs_identical,
            ratio: unified_median_ns as f64 / leaf_median_ns as f64,
            leaf_ns,
            unified_ns,
            leaf_median_ns,
            unified_median_ns,
        })
    }

    /// Whether `LITCHI_ODS_BENCH_ONLY` (when set) names `scenario`.
    fn selected(scenario: &str) -> bool {
        std::env::var("LITCHI_ODS_BENCH_ONLY").map_or(true, |only| {
            only.split(',').any(|wanted| wanted == scenario)
        })
    }

    /// The unified leg of the query scenarios, kept out of line so a
    /// profiler can select it (`--toggle-collect=*unified_query*`).
    #[inline(never)]
    fn unified_query(file: &File) -> usize {
        file.query("cell")
            .expect("query")
            .map(|node| {
                node.expect("node")
                    .view()
                    .text()
                    .expect("text")
                    .map_or(0, |text| text.len())
            })
            .sum::<usize>()
    }

    /// Record 0315's fair leaf leg of `query_warm`, as record 0280's owned
    /// legs are for XLSX and PPTX: every content cell instance in document
    /// order (repeated rows and cells expanded, as the unified walk yields
    /// them), each kept as an owned, retainable handle (the shared workbook
    /// projection and the cell's coordinates, which is what a unified node
    /// holds), and its text read. A byte source has no source to re-check,
    /// so there is no fenced variant.
    #[inline(never)]
    fn leaf_owned(package: &Package) -> usize {
        let workbook = package.workbook().expect("projection");
        let mut total = 0_usize;
        for (sheet, data) in workbook.sheets().iter().enumerate() {
            let mut row_index = 0_usize;
            for row in &data.rows {
                for _ in 0..row.repeat() {
                    let mut column = 0_usize;
                    for cell in &row.cells {
                        for _ in 0..cell.repeat() {
                            if !cell.text.is_empty() {
                                black_box((workbook.clone(), sheet, row_index, column));
                                total += black_box(cell.text.as_str()).len();
                            }
                            column += 1;
                        }
                    }
                    row_index += 1;
                }
            }
        }
        total
    }

    fn cell(index: usize) -> (usize, usize, String) {
        let row = index / COLUMNS;
        let column = index % COLUMNS;
        let mut path = String::from("/Data/");
        let mut letters = Vec::new();
        let mut number = column + 1;
        while number > 0 {
            letters.push(b'A' + ((number - 1) % 26) as u8);
            number = (number - 1) / 26;
        }
        letters.reverse();
        path.push_str(std::str::from_utf8(&letters).expect("letters"));
        path.push_str(&(row + 1).to_string());
        (row, column, path)
    }

    #[inline(never)]
    fn leaf_set(package: &Package, count: usize) -> Package {
        let mut edit = package.edit().expect("leaf edit");
        edit.worksheets(|worksheets| {
            let mut batch = Vec::new();
            for index in 0..count {
                let (row, column, _) = cell(index);
                batch.push(CellChange::new(
                    row,
                    column,
                    Cell::new(CellValue::Number(-1.0), "-1"),
                ));
                if batch.len() == MAX_CELL_CHANGES {
                    worksheets.set_cells(
                        Selector::Position(Position::new(0)),
                        std::mem::take(&mut batch),
                    )?;
                }
            }
            if !batch.is_empty() {
                worksheets.set_cells(Selector::Position(Position::new(0)), batch)?;
            }
            Ok(())
        })
        .expect("leaf worksheet edit");
        edit.commit().expect("leaf commit").into_package()
    }

    #[inline(never)]
    fn unified_batch(file: &File, items: &[BatchItem<'static>]) -> File {
        let commit = file
            .batch(items.iter().cloned(), BatchMode::Atomic)
            .expect("unified batch");
        assert_eq!(
            commit.diagnostics().commits(),
            1,
            "one coalesced leaf commit"
        );
        commit.into_file()
    }

    pub(super) fn run() {
        let rounds = std::env::var("LITCHI_ODS_BENCH_SAMPLES")
            .ok()
            .and_then(|value| value.parse().ok())
            .filter(|value| *value > 0)
            .unwrap_or(5);
        let rows = std::env::var("LITCHI_ODS_BENCH_ROWS")
            .ok()
            .and_then(|value| value.parse().ok())
            .filter(|value| *value > 0)
            .unwrap_or(1_000);
        let bytes = fixture(rows);
        let package = Package::from_bytes(&bytes).expect("leaf open");
        let file = litchi_ops::open_bytes(bytes.clone()).expect("unified open");
        let mut scenarios = Vec::new();

        scenarios.extend(abba(
            "open_get",
            rounds,
            None,
            || {
                time(|| {
                    let package = Package::from_bytes(&bytes).expect("leaf open");
                    let workbook = package.workbook().expect("projection");
                    workbook.sheets()[0].cell(4, 2).expect("cell").text.len()
                })
            },
            || {
                time(|| {
                    let file = litchi_ops::open_bytes(bytes.clone()).expect("unified open");
                    let node = file.get("/Data/C5").expect("get");
                    node.view()
                        .text()
                        .expect("text")
                        .map_or(0, |text| text.len())
                })
            },
        ));

        scenarios.extend(abba(
            "query",
            rounds,
            None,
            || {
                time(|| {
                    let workbook = package.workbook().expect("projection");
                    let mut total = 0_usize;
                    for sheet in workbook.sheets() {
                        for row in &sheet.rows {
                            for _ in 0..row.repeat() {
                                for cell in &row.cells {
                                    if !cell.text.is_empty() {
                                        total += cell.repeat() * cell.text.len();
                                    }
                                }
                            }
                        }
                    }
                    total
                })
            },
            || {
                // A fresh handle per sample, so the content index is built
                // inside the timed region as a first query builds it.
                let fresh = litchi_ops::open_bytes(bytes.clone()).expect("unified open");
                time(|| unified_query(&fresh))
            },
        ));

        // Record 0280: fair pairings of the same query. `query_warm` reads a
        // warm handle on both sides (the leaf's projection and the unified
        // content index already built); `query_cold` builds both inside the
        // clock on a freshly opened handle.
        let warm = litchi_ops::open_bytes(bytes.clone()).expect("unified open");
        let _ = warm.query("cell").expect("query").count();
        scenarios.extend(abba(
            "query_warm",
            rounds,
            None,
            || {
                time(|| {
                    let workbook = package.workbook().expect("projection");
                    let mut total = 0_usize;
                    for sheet in workbook.sheets() {
                        for row in &sheet.rows {
                            for _ in 0..row.repeat() {
                                for cell in &row.cells {
                                    if !cell.text.is_empty() {
                                        total += cell.repeat() * cell.text.len();
                                    }
                                }
                            }
                        }
                    }
                    total
                })
            },
            || time(|| unified_query(&warm)),
        ));
        // Record 0315: the same warm query against the fair leaf leg, under
        // a new scenario name; both legs read the same text, checked once
        // outside the clock.
        assert_eq!(
            leaf_owned(&package),
            unified_query(&warm),
            "owned leaf text"
        );
        scenarios.extend(abba(
            "query_warm_owned",
            rounds,
            None,
            || time(|| leaf_owned(&package)),
            || time(|| unified_query(&warm)),
        ));
        scenarios.extend(abba(
            "query_cold",
            rounds,
            None,
            || {
                let fresh = Package::from_bytes(&bytes).expect("leaf open");
                time(|| {
                    let workbook = fresh.workbook().expect("projection");
                    let mut total = 0_usize;
                    for sheet in workbook.sheets() {
                        for row in &sheet.rows {
                            for _ in 0..row.repeat() {
                                for cell in &row.cells {
                                    if !cell.text.is_empty() {
                                        total += cell.repeat() * cell.text.len();
                                    }
                                }
                            }
                        }
                    }
                    total
                })
            },
            || {
                let fresh = litchi_ops::open_bytes(bytes.clone()).expect("unified open");
                time(|| unified_query(&fresh))
            },
        ));

        let leaf_output = if selected("set") {
            leaf_set(&package, 1).to_bytes().expect("bytes")
        } else {
            Vec::new()
        };
        let unified_output = if selected("set") {
            file.set("/Data/A1", Props::new().with("value", -1))
                .expect("unified set")
                .to_bytes()
                .expect("bytes")
        } else {
            Vec::new()
        };
        scenarios.extend(abba(
            "set",
            rounds,
            Some(leaf_output == unified_output),
            || time(|| leaf_set(&package, 1)),
            || {
                time(|| {
                    file.set("/Data/A1", Props::new().with("value", -1))
                        .expect("unified set")
                })
            },
        ));

        for count in [1, 10, 100, 1_000, 10_000] {
            if count > rows * COLUMNS || !selected(&format!("batch_{count}")) {
                continue;
            }
            let items: Vec<BatchItem<'static>> = (0..count)
                .map(|index| {
                    let (_, _, path) = cell(index);
                    let mut item = BatchItem::new("set");
                    item.path = Some(path.into());
                    item.props = Some(Props::new().with("value", -1));
                    item
                })
                .collect();
            let leaf_output = leaf_set(&package, count).to_bytes().expect("bytes");
            let unified_output = unified_batch(&file, &items).to_bytes().expect("bytes");
            scenarios.extend(abba(
                format!("batch_{count}"),
                rounds,
                Some(leaf_output == unified_output),
                || time(|| leaf_set(&package, count)),
                || time(|| unified_batch(&file, &items)),
            ));
        }

        let report = Report {
            schema: "litchi.ods_unified_bench.v1",
            rows,
            columns: COLUMNS,
            input_bytes: bytes.len(),
            rounds,
            order: "ABBA",
            scenarios,
        };
        println!(
            "{}",
            serde_json::to_string_pretty(&report).expect("serialize report")
        );
    }
}

fn main() {
    #[cfg(feature = "ods")]
    enabled::run();
}
