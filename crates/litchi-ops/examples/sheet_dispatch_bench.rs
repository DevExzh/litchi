//! Before/after measurements for record 0353: the formula-evaluation path
//! and per-cell sheet reads through the generic `litchi-core::sheet` traits
//! on XLSX, XLSB and XLS.
//!
//! ```text
//! cargo build --release -p litchi-ops --features xlsx,xlsb,xls,eval --example sheet_dispatch_bench
//! taskset -c 2 target/release/examples/sheet_dispatch_bench [rows] [samples] > run.json
//! ```
//!
//! One process measures every scenario `samples` times (after one warm-up)
//! and prints the median and minimum in microseconds. The before/after
//! comparison runs the baseline and candidate binaries in ABBA process
//! order (`tools`-free: see record 0353 for the driver loop).
//!
//! Per format, the workbook is a `rows` x 4 sheet: a number in column A
//! and three uncached formulas per row (`=A*2+1`, an `IF`, a `SUM` over the
//! row).
//!
//! * `recalculate/<format>`: `File::recalculate`, the unified evaluation
//!   path (snapshot grid, the engine over the grid's sheet traits).
//! * `evaluate/<format>`: `litchi_eval::FormulaEvaluator` directly over the
//!   `WorkbookView` of the file, evaluating every formula cell.
//! * `cell/<format>`: `Worksheet::cell(row, column)` and `Cell::value` for
//!   every cell of the sheet.
//! * `cell_value/<format>`: `Worksheet::cell_value(row, column)` for every
//!   cell.
//! * `cells/<format>`: one pass of `Worksheet::cells()`.
//! * `recalculate_heavy/<format>`: `File::recalculate` of a formula-heavy
//!   sheet of `rows / 8` rows: running `SUM`s over growing ranges, a
//!   conditional with a division or a `MAX` over them and an `AVERAGE` plus `COUNT`
//!   over every row so far (record 0353a).

use std::future::Future;
use std::time::Instant;

use litchi_core::FileFormat;
use litchi_core::sheet::{Cell as _, WorkbookTrait, Worksheet as _};
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

fn median(samples: &mut [f64]) -> f64 {
    samples.sort_by(f64::total_cmp);
    let middle = samples.len() / 2;
    if samples.len() % 2 == 0 {
        f64::midpoint(samples[middle - 1], samples[middle])
    } else {
        samples[middle]
    }
}

/// Measure `leg` as scenario `name`, unless `SHEET_DISPATCH_ONLY` names
/// another scenario (for instruction-count profiling of one scenario).
fn record(
    out: &mut serde_json::Map<String, Value>,
    name: String,
    samples: usize,
    leg: impl FnMut(),
) {
    if std::env::var("SHEET_DISPATCH_ONLY").is_ok_and(|only| only != name) {
        return;
    }
    out.insert(name, measure(samples, leg));
}

fn measure(samples: usize, mut leg: impl FnMut()) -> Value {
    leg();
    let mut times = Vec::with_capacity(samples);
    for _ in 0..samples {
        let start = Instant::now();
        leg();
        times.push(start.elapsed().as_secs_f64() * 1e6);
    }
    let all = times.clone();
    let median = median(&mut times);
    json!({"median_us": median, "min_us": times[0], "samples_us": all})
}

fn limits() -> SemanticLimits {
    SemanticLimits::builder()
        .max_expanded_targets(1 << 22)
        .and_then(|builder| builder.max_staged_bytes(1 << 30))
        .and_then(|builder| builder.max_snapshot_bytes(1 << 32))
        .and_then(litchi_ops::transaction::SemanticLimitsBuilder::build)
        .expect("bench limits")
}

fn workbook(format: FileFormat, count: u32) -> File {
    let mut source = String::new();
    for row in 1..=count {
        source.push_str(&format!(
            "{row},=A{row}*2+1,\"=IF(B{row}>10,B{row},0)\",=SUM(A{row}:C{row})\n"
        ));
    }
    import(format, &source)
}

/// A formula-heavy sheet: running totals over growing ranges, a nested
/// conditional over them and an average over every row so far, so each
/// formula reads many cells and references other formulas.
fn heavy_workbook(format: FileFormat, count: u32) -> File {
    let mut source = String::new();
    for row in 1..=count {
        source.push_str(&format!(
            "{row},=SUM(A1:A{row}),\"=IF(B{row}>A{row}*A{row}/4,B{row}/A{row},MAX(A1:B{row}))\",=AVERAGE(A1:C{row})+COUNT(B1:B{row})\n"
        ));
    }
    import(format, &source)
}

fn import(format: FileFormat, source: &str) -> File {
    let sheet = if format == FileFormat::Xlsx {
        "/Sheet1"
    } else {
        "/sheet[1]"
    };
    let file = litchi_ops::create(format)
        .expect("create")
        .import_with_options(
            sheet,
            source.as_bytes(),
            &Options::csv().formulas(true),
            litchi_ops::EditOptions::new().semantic_limits(limits()),
        )
        .expect("formulas")
        .into_file();
    litchi_ops::open_bytes(file.to_bytes().expect("publish")).expect("reopen")
}

fn scenarios(
    format: FileFormat,
    count: u32,
    samples: usize,
    out: &mut serde_json::Map<String, Value>,
) {
    let file = workbook(format, count);
    let tag = format!("{format:?}").to_lowercase();
    let options = CalcOptions::new().max_cells(1 << 24);
    let report = file.recalculate(&options).expect("recalculate");
    assert_eq!(report.formulas().len(), (count * 3) as usize, "{report:?}");
    assert_eq!(report.unsupported(), 0, "{report:?}");
    record(out, format!("recalculate/{tag}"), samples, || {
        let report = file.recalculate(&options).expect("recalculate");
        std::hint::black_box(report);
    });
    let heavy_rows = (count / 8).max(1);
    let heavy = heavy_workbook(format, heavy_rows);
    let report = heavy.recalculate(&options).expect("recalculate");
    assert_eq!(
        report.formulas().len(),
        (heavy_rows * 3) as usize,
        "{report:?}"
    );
    assert_eq!(report.unsupported(), 0, "{report:?}");
    record(out, format!("recalculate_heavy/{tag}"), samples, || {
        let report = heavy.recalculate(&options).expect("recalculate");
        std::hint::black_box(report);
    });
    // XLS addresses the generic traits zero-based, XLSX and XLSB one-based.
    let base = u32::from(format != FileFormat::Xls);
    let (first, last) = (base, count - 1 + base);
    let view = file.workbook().expect("workbook view");
    let name = WorkbookTrait::worksheet_names(&view)[0].clone();
    record(out, format!("evaluate/{tag}"), samples, || {
        let evaluator = litchi_eval::FormulaEvaluator::new(&view);
        for row in first..=last {
            for column in base + 1..base + 4 {
                let value = block_on(evaluator.evaluate_cell(&name, row, column)).expect("value");
                std::hint::black_box(value);
            }
        }
    });
    let sheet = WorkbookTrait::worksheet_by_index(&view, 0).expect("first sheet");
    record(out, format!("cell/{tag}"), samples, || {
        for row in first..=last {
            for column in base..base + 4 {
                let cell = sheet.cell(row, column).expect("cell");
                std::hint::black_box(cell.value());
            }
        }
    });
    record(out, format!("cell_value/{tag}"), samples, || {
        for row in first..=last {
            for column in base..base + 4 {
                let value = sheet.cell_value(row, column).expect("cell value");
                std::hint::black_box(&value);
            }
        }
    });
    record(out, format!("cells/{tag}"), samples, || {
        let mut seen = 0usize;
        for cell in sheet.cells() {
            std::hint::black_box(cell.expect("cell").value());
            seen += 1;
        }
        assert!(seen >= count as usize * 4, "{seen}");
    });
}

fn main() {
    let count: u32 = std::env::args()
        .nth(1)
        .map_or(2_000, |value| value.parse().expect("rows"));
    let samples: usize = std::env::args()
        .nth(2)
        .map_or(15, |value| value.parse().expect("samples"));
    let mut out = serde_json::Map::new();
    for format in [FileFormat::Xlsx, FileFormat::Xlsb, FileFormat::Xls] {
        // The BIFF8 editor's import grows with the worksheet: XLS runs a
        // smaller table, as record 0266 does.
        let rows = if format == FileFormat::Xls {
            count.min(1_000)
        } else {
            count
        };
        scenarios(format, rows, samples, &mut out);
    }
    let report = json!({
        "schema": "litchi.sheet-dispatch-bench.v1",
        "record": "0353",
        "rows": count,
        "samples": samples,
        "scenarios": out,
    });
    println!(
        "{}",
        serde_json::to_string_pretty(&report).expect("serialize the bench report")
    );
}
