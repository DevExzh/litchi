//! Basic functions formula evaluator example.
//!
//! This example exercises Phase 1 core functions implemented in the
//! shared formula evaluator: aggregates, scalar math, logical, and
//! simple text functions.
//!
//! Run with:
//!
//! ```bash
//! cargo run --example sheet_eval_functions_basic --features ooxml -- sheet_eval_functions_basic.xlsx
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
        "sheet_eval_functions_basic.xlsx"
    };

    build_sample_xlsx(path)?;

    let xlsx_wb = XlsxWorkbook::open(path)?;
    let evaluator = FormulaEvaluator::new(&xlsx_wb);

    println!("Evaluating basic functions on Funcs in {}", path);

    // Aggregates
    for (coord, row, col) in [
        ("C1", 1, 3),
        ("C2", 2, 3),
        ("C3", 3, 3),
        ("C4", 4, 3),
        ("C5", 5, 3),
    ] {
        let value = evaluator.evaluate_cell("Funcs", row, col).await?;
        println!("  {} => {:?}", coord, value);
    }

    // Math and trig
    for (coord, row, col) in [("D1", 1, 4), ("D2", 2, 4), ("D3", 3, 4), ("D4", 4, 4)] {
        let value = evaluator.evaluate_cell("Funcs", row, col).await?;
        println!("  {} => {:?}", coord, value);
    }

    // Logical
    for (coord, row, col) in [("E1", 1, 5), ("E2", 2, 5), ("E3", 3, 5), ("E4", 4, 5)] {
        let value = evaluator.evaluate_cell("Funcs", row, col).await?;
        println!("  {} => {:?}", coord, value);
    }

    // Text
    for (coord, row, col) in [
        ("F1", 1, 6),
        ("F2", 2, 6),
        ("F3", 3, 6),
        ("F4", 4, 6),
        ("F5", 5, 6),
        ("F6", 6, 6),
    ] {
        let value = evaluator.evaluate_cell("Funcs", row, col).await?;
        println!("  {} => {:?}", coord, value);
    }

    Ok(())
}

fn build_sample_xlsx(path: &str) -> ExampleResult<()> {
    let mut wb = XlsxWorkbook::create()?;

    wb.add_worksheet("Funcs");
    {
        let ws = wb.worksheet_mut(0)?;
        ws.set_name("Funcs".to_string());

        // Base values for aggregates and TEXTJOIN
        ws.set_cell_value(1, 1, 1); // A1
        ws.set_cell_value(2, 1, 2); // A2
        ws.set_cell_value(3, 1, 3); // A3

        ws.set_cell_value(1, 2, " Hello "); // B1
        ws.set_cell_value(2, 2, "World"); // B2

        // C1..C5: aggregates
        ws.set_cell_formula(1, 3, "MIN(A1:A3)"); // C1 = 1
        ws.set_cell_formula(2, 3, "MAX(A1:A3)"); // C2 = 3
        ws.set_cell_formula(3, 3, "AVERAGE(A1:A3)"); // C3 = 2
        ws.set_cell_formula(4, 3, "COUNT(A1:A3)"); // C4 = 3
        ws.set_cell_formula(5, 3, "COUNTA(A1:B3)"); // C5 = 5

        // D1..D4: math/trig
        ws.set_cell_formula(1, 4, "ABS(-5)"); // D1 = 5
        ws.set_cell_formula(2, 4, "POWER(2,3)"); // D2 = 8
        ws.set_cell_formula(3, 4, "SQRT(16)"); // D3 = 4
        ws.set_cell_formula(4, 4, "SIN(PI()/2)"); // D4 ~= 1

        // E1..E4: logical
        ws.set_cell_value(4, 1, true); // A4: used as IF condition
        ws.set_cell_formula(1, 5, "IF(A4,1,0)"); // E1 = 1
        ws.set_cell_formula(2, 5, "AND(TRUE, FALSE)"); // E2 = FALSE
        ws.set_cell_formula(3, 5, "OR(FALSE, FALSE, TRUE)"); // E3 = TRUE
        ws.set_cell_formula(4, 5, "NOT(TRUE)"); // E4 = FALSE

        // F1..F6: text
        ws.set_cell_formula(1, 6, "LEN(\"abc\")"); // F1 = 3
        ws.set_cell_formula(2, 6, "LOWER(\"AbC\")"); // F2 = "abc"
        ws.set_cell_formula(3, 6, "UPPER(\"AbC\")"); // F3 = "ABC"
        ws.set_cell_formula(4, 6, "TRIM(B1)"); // F4 = "Hello"
        ws.set_cell_formula(5, 6, "CONCAT(\"Hello\",\" \",\"World\")"); // F5 = "Hello World"
        ws.set_cell_formula(6, 6, "TEXTJOIN(\"-\",TRUE,A1:A3)"); // F6 = "1-2-3"
    }

    wb.save(path)?;
    println!("Created basic functions test workbook at {}", path);

    Ok(())
}
