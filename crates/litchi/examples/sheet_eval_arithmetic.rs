//! Arithmetic formula evaluator example.
//!
//! This example focuses on the Phase 2 capabilities of the shared
//! formula evaluator: basic arithmetic over literals and simple
//! cell references.
//!
//! Run with:
//!
//! ```bash
//! cargo run --example sheet_eval_arithmetic --features ooxml -- sheet_eval_arithmetic.xlsx
//! ```

use litchi::ooxml::xlsx::{Formula, Workbook};
use litchi::sheet::{FormulaEvaluator, functions::open_workbook};
use std::env;
use std::error::Error;

type ExampleResult<T> = Result<T, Box<dyn Error + Send + Sync>>;

#[tokio::main]
async fn main() -> ExampleResult<()> {
    let args: Vec<String> = env::args().collect();
    let path = if args.len() > 1 {
        args[1].as_str()
    } else {
        "sheet_eval_arithmetic.xlsx"
    };

    build_sample_xlsx(path)?;

    let xlsx_wb = open_workbook(path)?;
    let evaluator = FormulaEvaluator::new(xlsx_wb.as_ref());

    println!("Evaluating arithmetic formulas on Sheet1 in {}", path);

    for (coord, col) in ["C1", "D1", "E1", "F1", "G1", "H1"].iter().zip(3u32..) {
        let value = evaluator.evaluate_cell("Sheet1", 1, col).await?;
        println!("  {} => {:?}", coord, value);
    }

    Ok(())
}

fn build_sample_xlsx(path: &str) -> ExampleResult<()> {
    let wb = Workbook::create()?;
    let mut edit = wb.edit()?;
    edit.tab(0)?
        .ok_or("missing worksheet tab")?
        .rename("Sheet1")?;

    // Sheet1: base values + arithmetic formulas.
    {
        let mut ws = edit.sheet(0)?.ok_or("missing worksheet")?;

        // Base values
        ws.set("A1", 2_i32)?; // A1
        ws.set("B1", 3_i32)?; // B1

        // C1: A1 + B1 = 5
        ws.set("C1", Formula::new("A1+B1")?)?;

        // D1: A1*B1 + 10 = 16
        ws.set("D1", Formula::new("A1*B1+10")?)?;

        // E1: -(A1) = -2
        ws.set("E1", Formula::new("-A1")?)?;

        // F1: (A1+B1)*10 = 50
        ws.set("F1", Formula::new("(A1+B1)*10")?)?;

        // G1: Sheet2!A1 + 1 (Sheet2 created below)
        ws.set("G1", Formula::new("Sheet2!A1+1")?)?;

        // H1: SUM over a range: SUM(A1:B1) = 5
        ws.set("H1", Formula::new("SUM(A1:B1)")?)?;
    }

    // Sheet2: value used in cross-sheet arithmetic.
    let mut ws2 = edit.add("Sheet2")?;
    {
        ws2.set("A1", 100_i32)?; // A1
    }

    edit.commit()?.into_workbook().save(path)?;
    println!("Created arithmetic test workbook at {}", path);

    Ok(())
}
