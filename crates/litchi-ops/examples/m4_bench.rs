//! Paired leaf-vs-unified measurements for record 0266: CSV import
//! throughput and formula-evaluation cost.
//!
//! ```text
//! cargo build --release -p litchi-ops --features xlsx,xlsb,xls,eval --example m4_bench
//! taskset -c 2 target/release/examples/m4_bench [rows] [rounds] > m4.json
//! ```
//!
//! Each scenario runs its two legs in ABBA order for `rounds` rounds (default
//! 7) and reports every sample, the median and the minimum in microseconds,
//! and the unified/leaf ratio of the medians.
//!
//! * `import/<format>`: the leaf leg stages the already-typed values of a
//!   `rows` x 5 table (integer, text, decimal, boolean, ISO date) in one
//!   native edit of the opened package and commits it; the unified leg is
//!   `File::import` of the same table as CSV text, so it also decodes,
//!   parses and infers. XLS stores the date column as text in both legs.
//! * `eval/xlsx`: a `rows` x 4 sheet with three uncached formulas per row.
//!   The leaf leg evaluates every formula cell with `litchi-eval` directly
//!   over the handle's generic `WorkbookTrait`; the unified leg is
//!   `File::recalculate`, which loads a snapshot, evaluates and compares.

use std::future::Future;
use std::time::Instant;

use litchi_core::FileFormat;
use litchi_core::sheet::WorkbookTrait;
use litchi_ops::File;
use litchi_ops::calc::Options as CalcOptions;
use litchi_ops::import::Options;
use litchi_ops::transaction::SemanticLimits;
use serde_json::{Value, json};

fn block_on<F: Future>(future: F) -> F::Output {
    let mut future = std::pin::pin!(future);
    let mut context = std::task::Context::from_waker(std::task::Waker::noop());
    match future.as_mut().poll(&mut context) {
        std::task::Poll::Ready(output) => output,
        std::task::Poll::Pending => panic!("the engine suspended"),
    }
}

/// One table row: id, name, score, active, joined (as days of 2026).
struct Row {
    id: u32,
    name: String,
    score: f64,
    active: bool,
    day: u32,
}

fn rows(count: u32) -> Vec<Row> {
    (0..count)
        .map(|index| Row {
            id: index + 1,
            name: format!("name {index}"),
            score: f64::from(index % 1000) / 8.0 + 0.25,
            active: index % 3 == 0,
            day: index % 28 + 1,
        })
        .collect()
}

fn csv(rows: &[Row]) -> Vec<u8> {
    let mut text = String::new();
    for row in rows {
        text.push_str(&format!(
            "{},{},{},{},2026-01-{:02}\n",
            row.id,
            row.name,
            row.score,
            if row.active { "TRUE" } else { "FALSE" },
            row.day
        ));
    }
    text.into_bytes()
}

/// 2026-01-dd as an Excel 1900-system serial.
fn serial(day: u32) -> f64 {
    46_022.0 + f64::from(day) - 1.0
}

fn time(mut leg: impl FnMut()) -> f64 {
    let start = Instant::now();
    leg();
    start.elapsed().as_secs_f64() * 1e6
}

fn median(samples: &mut [f64]) -> f64 {
    samples.sort_by(f64::total_cmp);
    let middle = samples.len() / 2;
    if samples.len() % 2 == 0 {
        f64::midpoint(samples[middle - 1], samples[middle])
    } else {
        samples[middle]
    }
}

fn abba(rounds: usize, mut leaf: impl FnMut(), mut unified: impl FnMut()) -> Value {
    // One warm-up of each leg.
    leaf();
    unified();
    let (mut a, mut b) = (Vec::new(), Vec::new());
    for _ in 0..rounds {
        a.push(time(&mut leaf));
        b.push(time(&mut unified));
        b.push(time(&mut unified));
        a.push(time(&mut leaf));
    }
    let (leaf_samples, unified_samples) = (a.clone(), b.clone());
    let (leaf_median, unified_median) = (median(&mut a), median(&mut b));
    json!({
        "leaf_us": {"median": leaf_median, "min": a[0], "samples": leaf_samples},
        "unified_us": {"median": unified_median, "min": b[0], "samples": unified_samples},
        "ratio_of_medians": unified_median / leaf_median,
    })
}

fn limits() -> SemanticLimits {
    SemanticLimits::builder()
        .max_expanded_targets(1 << 22)
        .and_then(|builder| builder.max_staged_bytes(1 << 30))
        .and_then(|builder| builder.max_snapshot_bytes(1 << 32))
        .and_then(litchi_ops::transaction::SemanticLimitsBuilder::build)
        .expect("bench limits")
}

fn import_xlsx(bytes: &[u8], table: &[Row], source: &[u8], rounds: usize) -> Value {
    use litchi_xlsx::cell::{Content, Number, Value as Cell};
    let package = litchi_xlsx::Package::from_bytes(bytes).expect("leaf open");
    let file = litchi_ops::open_bytes(bytes.to_vec()).expect("unified open");
    let limits = limits();
    abba(
        rounds,
        || {
            let mut edit = package.edit().expect("edit");
            let mut sheet = edit.sheet(0usize).expect("sheet").expect("first sheet");
            for (index, row) in table.iter().enumerate() {
                let r = u32::try_from(index).expect("row");
                let at = |column| litchi_xlsx::Address::at(r, column).expect("address");
                sheet
                    .set(
                        at(0),
                        Content::Value(Cell::from(Number::new(row.id.to_string()).expect("n"))),
                    )
                    .expect("id");
                sheet
                    .set(at(1), Content::Value(Cell::text(row.name.as_str())))
                    .expect("name");
                sheet
                    .set(
                        at(2),
                        Content::Value(Cell::from(Number::try_from(row.score).expect("n"))),
                    )
                    .expect("score");
                sheet.set(at(3), row.active).expect("active");
                sheet
                    .set(
                        at(4),
                        Content::Value(Cell::from(Number::try_from(serial(row.day)).expect("n"))),
                    )
                    .expect("joined");
                sheet
                    .number_format(at(4), Some("yyyy-mm-dd"))
                    .expect("format");
            }
            drop(edit.commit().expect("commit"));
        },
        || {
            let commit = file
                .import_with_options(
                    "/sheet[1]",
                    source,
                    &Options::csv(),
                    litchi_ops::EditOptions::new().semantic_limits(limits),
                )
                .expect("import");
            assert_eq!(commit.value().cells(), table.len() * 5);
        },
    )
}

fn import_xlsb(bytes: &[u8], table: &[Row], source: &[u8], rounds: usize) -> Value {
    use litchi_xlsb::cell_values::{Format, Reference, Scalar};
    let package = litchi_xlsb::Package::from_bytes(bytes).expect("leaf open");
    let file = litchi_ops::open_bytes(bytes.to_vec()).expect("unified open");
    let limits = limits();
    abba(
        rounds,
        || {
            let mut edit = package
                .edit_cell_values(0usize)
                .expect("edit")
                .resolve()
                .expect("worksheet resolution")
                .expect("first sheet");
            let mut date = Format::default();
            date.number_format = Some("yyyy-mm-dd".to_owned());
            for (index, row) in table.iter().enumerate() {
                let r = u32::try_from(index).expect("row");
                let at = |column| Reference::new(r, column).expect("reference");
                edit.set_scalar(at(0), Scalar::Number(f64::from(row.id)))
                    .expect("id");
                edit.set_scalar(at(1), Scalar::Text(row.name.clone()))
                    .expect("name");
                edit.set_scalar(at(2), Scalar::Number(row.score))
                    .expect("score");
                edit.set_scalar(at(3), Scalar::Boolean(row.active))
                    .expect("active");
                edit.set_scalar(at(4), Scalar::Number(serial(row.day)))
                    .expect("joined");
                edit.set_format(at(4), &date).expect("format");
            }
            drop(edit.commit().expect("commit"));
        },
        || {
            let commit = file
                .import_with_options(
                    "/sheet[1]",
                    source,
                    &Options::csv(),
                    litchi_ops::EditOptions::new().semantic_limits(limits),
                )
                .expect("import");
            assert_eq!(commit.value().cells(), table.len() * 5);
        },
    )
}

fn import_xls(bytes: &[u8], table: &[Row], source: &[u8], rounds: usize) -> Value {
    use litchi_xls::cell_values::{Reference, Selector, Value as Cell};
    let package = litchi_xls::Package::from_bytes(bytes).expect("leaf open");
    let file = litchi_ops::open_bytes(bytes.to_vec()).expect("unified open");
    let limits = limits();
    abba(
        rounds,
        || {
            // The leaf admits 4,096 staged changes per edit: commit every
            // 1,024 cells (5 per row), as the unified import does.
            let sheet = || Selector::Position(litchi_core::Position::new(0));
            let mut current = package.clone();
            for (chunk_index, chunk) in table.chunks(1_024 / 5).enumerate() {
                let mut edit = current.edit().expect("edit");
                for (offset, row) in chunk.iter().enumerate() {
                    let r = u32::try_from(chunk_index * (1_024 / 5) + offset).expect("row");
                    let at = |column| Reference::new(r, column).expect("reference");
                    edit.insert_cell(sheet(), at(0), Cell::Number(f64::from(row.id)))
                        .expect("id");
                    edit.insert_cell(sheet(), at(1), Cell::Text(row.name.clone()))
                        .expect("name");
                    edit.insert_cell(sheet(), at(2), Cell::Number(row.score))
                        .expect("score");
                    edit.insert_cell(sheet(), at(3), Cell::Boolean(row.active))
                        .expect("active");
                    edit.insert_cell(
                        sheet(),
                        at(4),
                        Cell::Text(format!("2026-01-{:02}", row.day)),
                    )
                    .expect("joined");
                }
                current = edit.commit().expect("commit").into_package();
            }
            drop(current);
        },
        || {
            let commit = file
                .import_with_options(
                    "/sheet[1]",
                    source,
                    &Options::csv(),
                    litchi_ops::EditOptions::new().semantic_limits(limits),
                )
                .expect("import");
            assert_eq!(commit.value().cells(), table.len() * 5);
        },
    )
}

fn eval_xlsx(count: u32, rounds: usize) -> Value {
    let mut source = String::new();
    for row in 1..=count {
        source.push_str(&format!(
            "{row},=A{row}*2+1,\"=IF(B{row}>10,B{row},0)\",=SUM(A{row}:C{row})\n"
        ));
    }
    let file = litchi_ops::create(FileFormat::Xlsx)
        .expect("create")
        .import_with_options(
            "/Sheet1",
            source.as_bytes(),
            &Options::csv().formulas(true),
            litchi_ops::EditOptions::new().semantic_limits(limits()),
        )
        .expect("formulas")
        .into_file();
    let file = litchi_ops::open_bytes(file.to_bytes().expect("publish")).expect("reopen");
    let File::Xlsx(handle) = &file else {
        unreachable!("an XLSX file")
    };
    let options = CalcOptions::new().max_cells(1 << 24);
    abba(
        rounds,
        || {
            let evaluator = litchi_eval::FormulaEvaluator::new(handle);
            let name = &WorkbookTrait::worksheet_names(handle)[0];
            for row in 1..=count {
                for column in 2..=4 {
                    let value =
                        block_on(evaluator.evaluate_cell(name, row, column)).expect("value");
                    std::hint::black_box(value);
                }
            }
        },
        || {
            let report = file.recalculate(&options).expect("recalculate");
            assert_eq!(report.uncached(), (count * 3) as usize, "{report:?}");
        },
    )
}

fn main() {
    let count: u32 = std::env::args()
        .nth(1)
        .map_or(10_000, |value| value.parse().expect("rows"));
    let rounds: usize = std::env::args()
        .nth(2)
        .map_or(7, |value| value.parse().expect("rounds"));
    let table = rows(count);
    let source = csv(&table);
    let mut scenarios = serde_json::Map::new();
    type Leg = fn(&[u8], &[Row], &[u8], usize) -> Value;
    let legs: [(FileFormat, Leg); 3] = [
        (FileFormat::Xlsx, import_xlsx),
        (FileFormat::Xlsb, import_xlsb),
        (FileFormat::Xls, import_xls),
    ];
    // The BIFF8 editor's cell insertion grows with the worksheet, so XLS
    // runs a smaller table (record 0266 reports this leaf cost).
    let xls_count = count.min(1_000);
    let xls_table = rows(xls_count);
    let xls_source = csv(&xls_table);
    for (format, run) in legs {
        let (count, table, source) = if format == FileFormat::Xls {
            (xls_count, &xls_table, &xls_source)
        } else {
            (count, &table, &source)
        };
        let bytes = litchi_ops::create(format)
            .expect("create")
            .to_bytes()
            .expect("publish");
        let start = Instant::now();
        let mut result = run(&bytes, table, source, rounds);
        result["rows"] = json!(count);
        eprintln!("import/{format:?}: {:.1} s", start.elapsed().as_secs_f64());
        let unified = result["unified_us"]["median"].as_f64().unwrap_or(f64::NAN);
        let leaf = result["leaf_us"]["median"].as_f64().unwrap_or(f64::NAN);
        result["unified_rows_per_s"] = json!(f64::from(count) / (unified / 1e6));
        result["leaf_rows_per_s"] = json!(f64::from(count) / (leaf / 1e6));
        scenarios.insert(format!("import/{format:?}").to_lowercase(), result);
    }
    let eval_rows = (count / 5).max(1);
    let mut eval = eval_xlsx(eval_rows, rounds);
    eval["formulas"] = json!(eval_rows * 3);
    scenarios.insert("eval/xlsx".to_owned(), eval);
    let report = json!({
        "schema": "litchi.m4-bench.v1",
        "record": "0266",
        "rows": count,
        "columns": 5,
        "csv_bytes": source.len(),
        "rounds": rounds,
        "order": "ABBA (leaf, unified, unified, leaf) per round after one warm-up of each leg",
        "scenarios": scenarios,
    });
    println!(
        "{}",
        serde_json::to_string_pretty(&report).expect("serialize the bench report")
    );
}
