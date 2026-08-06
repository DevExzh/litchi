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

use litchi::sheet::{FormulaEvaluator, functions::open_workbook};
use litchi::xlsx::{Formula, Workbook};
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
    let xlsx_wb = open_workbook(path)?;
    let evaluator = FormulaEvaluator::new(xlsx_wb.as_ref());

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
    let wb = Workbook::create()?;
    let mut edit = wb.edit()?;
    edit.tab(0)?
        .ok_or("missing worksheet tab")?
        .rename("Sheet1")?;

    // Sheet1: source values + formulas.
    {
        let mut ws = edit.sheet(0)?.ok_or("missing worksheet")?;

        // A1: plain number
        ws.set("A1", 10_i32)?;

        // B1: reference to A1 (no cached value)
        ws.set("B1", Formula::new("A1")?)?;

        // C1: string literal formula => "Hello"
        ws.set("C1", Formula::new("\"Hello\"")?)?;

        // D1: boolean literal TRUE
        ws.set("D1", Formula::new("TRUE")?)?;

        // E1: numeric literal 42
        ws.set("E1", Formula::new("42")?)?;

        // F1: reference to Sheet2!A1 (created below)
        ws.set("F1", Formula::new("Sheet2!A1")?)?;
    }

    // Sheet2: a single value that Sheet1!F1 refers to.
    let mut ws2 = edit.add("Sheet2")?;
    {
        ws2.set("A1", 123_i32)?;
    }

    edit.commit()?.into_workbook().save(path)?;
    println!("Created test workbook at {}", path);

    Ok(())
}
