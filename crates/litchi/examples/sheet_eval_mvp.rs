//! MVP formula evaluator example using the unified `sheet` API.
//!
//! This example builds a small XLSX workbook using the XLSX writer,
//! then reopens it via the unified `sheet::Workbook` and evaluates
//! a few simple formulas with `FormulaEvaluator`.
//!
//! It demonstrates Phase 1 capabilities:
//! - Prefer cached results when present (not exercised here, formulas have no cache).
//! - Evaluate literal formulas without cached values:
//!   - Numeric: `42`
//!   - Boolean: `TRUE`, `FALSE`
//!   - String: `"Hello"` (with escaped quotes as needed)
//! - Evaluate simple single-cell references: `A1`, `Sheet2!A1`.
//!
//! Run with:
//!
//! ```bash
//! cargo run --example sheet_eval_mvp --features ooxml -- sheet_eval_mvp.xlsx
//! ```

use litchi::ooxml::xlsx::Workbook as XlsxWorkbook;
use litchi::sheet::FormulaEvaluator;
use std::env;
use std::error::Error;

type ExampleResult<T> = Result<T, Box<dyn Error + Send + Sync>>;

#[tokio::main]
async fn main() -> ExampleResult<()> {
    let args: Vec<String> = env::args().collect();
    let path = if args.len() > 1 {
        args[1].as_str()
    } else {
        "sheet_eval_mvp.xlsx"
    };

    // 1) Build a small XLSX workbook with a few formulas and no cached values.
    build_sample_xlsx(path)?;

    // 2) Open via the XLSX workbook type that implements WorkbookTrait.
    let xlsx_wb = XlsxWorkbook::open(path)?;
    let evaluator = FormulaEvaluator::new(&xlsx_wb);

    // 3) Evaluate a few cells on Sheet1.
    println!("Evaluating formulas on Sheet1 in {}", path);

    for col in 3u32..=5 {
        let coord = match col {
            3 => "C1",
            4 => "D1",
            5 => "E1",
            _ => unreachable!(),
        };
        let value = evaluator.evaluate_cell("Sheet1", 1, col).await?;
        println!("  {} => {:?}", coord, value);
    }

    // Also demonstrate a cross-sheet reference: Sheet2!A1 referenced from Sheet1!F1.
    let value = evaluator.evaluate_cell("Sheet1", 1, 6).await?;
    println!("  F1 (Sheet2!A1) => {:?}", value);

    Ok(())
}

fn build_sample_xlsx(path: &str) -> ExampleResult<()> {
    let mut wb = XlsxWorkbook::create()?;

    // Sheet1: source values + formulas.
    {
        let ws = wb.worksheet_mut(0)?;
        ws.set_name("Sheet1".to_string());

        // A1: plain number
        ws.set_cell_value(1, 1, 10);

        // B1: reference to A1 (no cached value)
        ws.set_cell_formula(1, 2, "A1");

        // C1: string literal formula => "Hello"
        ws.set_cell_formula(1, 3, "\"Hello\"");

        // D1: boolean literal TRUE
        ws.set_cell_formula(1, 4, "TRUE");

        // E1: numeric literal 42
        ws.set_cell_formula(1, 5, "42");

        // F1: reference to Sheet2!A1 (created below)
        ws.set_cell_formula(1, 6, "Sheet2!A1");
    }

    // Sheet2: a single value that Sheet1!F1 refers to.
    wb.add_worksheet("Sheet2");
    {
        let ws2 = wb.worksheet_mut(1)?;
        ws2.set_cell_value(1, 1, 123);
    }

    wb.save(path)?;
    println!("Created test workbook at {}", path);

    Ok(())
}
