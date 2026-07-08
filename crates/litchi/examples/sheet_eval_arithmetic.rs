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
        "sheet_eval_arithmetic.xlsx"
    };

    build_sample_xlsx(path)?;

    let xlsx_wb = XlsxWorkbook::open(path)?;
    let evaluator = FormulaEvaluator::new(&xlsx_wb);

    println!("Evaluating arithmetic formulas on Sheet1 in {}", path);

    for (coord, col) in ["C1", "D1", "E1", "F1", "G1", "H1"].iter().zip(3u32..) {
        let value = evaluator.evaluate_cell("Sheet1", 1, col).await?;
        println!("  {} => {:?}", coord, value);
    }

    Ok(())
}

fn build_sample_xlsx(path: &str) -> ExampleResult<()> {
    let mut wb = XlsxWorkbook::create()?;

    // Sheet1: base values + arithmetic formulas.
    {
        let ws = wb.worksheet_mut(0)?;
        ws.set_name("Sheet1".to_string());

        // Base values
        ws.set_cell_value(1, 1, 2); // A1
        ws.set_cell_value(1, 2, 3); // B1

        // C1: A1 + B1 = 5
        ws.set_cell_formula(1, 3, "A1+B1");

        // D1: A1*B1 + 10 = 16
        ws.set_cell_formula(1, 4, "A1*B1+10");

        // E1: -(A1) = -2
        ws.set_cell_formula(1, 5, "-A1");

        // F1: (A1+B1)*10 = 50
        ws.set_cell_formula(1, 6, "(A1+B1)*10");

        // G1: Sheet2!A1 + 1 (Sheet2 created below)
        ws.set_cell_formula(1, 7, "Sheet2!A1+1");

        // H1: SUM over a range: SUM(A1:B1) = 5
        ws.set_cell_formula(1, 8, "SUM(A1:B1)");
    }

    // Sheet2: value used in cross-sheet arithmetic.
    wb.add_worksheet("Sheet2");
    {
        let ws2 = wb.worksheet_mut(1)?;
        ws2.set_cell_value(1, 1, 100); // A1
    }

    wb.save(path)?;
    println!("Created arithmetic test workbook at {}", path);

    Ok(())
}
