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
        "sheet_eval_functions_basic.xlsx"
    };

    build_sample_xlsx(path)?;

    let xlsx_wb = open_workbook(path)?;
    let evaluator = FormulaEvaluator::new(xlsx_wb.as_ref());

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
    let wb = Workbook::create()?;
    let mut edit = wb.edit()?;
    edit.tab(0)?
        .ok_or("missing worksheet tab")?
        .rename("Funcs")?;

    {
        let mut ws = edit.sheet(0)?.ok_or("missing worksheet")?;

        // Base values for aggregates and TEXTJOIN
        ws.set("A1", 1_i32)?; // A1
        ws.set("A2", 2_i32)?; // A2
        ws.set("A3", 3_i32)?; // A3

        ws.set("B1", " Hello ")?; // B1
        ws.set("B2", "World")?; // B2

        // C1..C5: aggregates
        ws.set("C1", Formula::new("MIN(A1:A3)")?)?; // C1 = 1
        ws.set("C2", Formula::new("MAX(A1:A3)")?)?; // C2 = 3
        ws.set("C3", Formula::new("AVERAGE(A1:A3)")?)?; // C3 = 2
        ws.set("C4", Formula::new("COUNT(A1:A3)")?)?; // C4 = 3
        ws.set("C5", Formula::new("COUNTA(A1:B3)")?)?; // C5 = 5

        // D1..D4: math/trig
        ws.set("D1", Formula::new("ABS(-5)")?)?; // D1 = 5
        ws.set("D2", Formula::new("POWER(2,3)")?)?; // D2 = 8
        ws.set("D3", Formula::new("SQRT(16)")?)?; // D3 = 4
        ws.set("D4", Formula::new("SIN(PI()/2)")?)?; // D4 ~= 1

        // E1..E4: logical
        ws.set("A4", true)?; // A4: used as IF condition
        ws.set("E1", Formula::new("IF(A4,1,0)")?)?; // E1 = 1
        ws.set("E2", Formula::new("AND(TRUE, FALSE)")?)?; // E2 = FALSE
        ws.set("E3", Formula::new("OR(FALSE, FALSE, TRUE)")?)?; // E3 = TRUE
        ws.set("E4", Formula::new("NOT(TRUE)")?)?; // E4 = FALSE

        // F1..F6: text
        ws.set("F1", Formula::new("LEN(\"abc\")")?)?; // F1 = 3
        ws.set("F2", Formula::new("LOWER(\"AbC\")")?)?; // F2 = "abc"
        ws.set("F3", Formula::new("UPPER(\"AbC\")")?)?; // F3 = "ABC"
        ws.set("F4", Formula::new("TRIM(B1)")?)?; // F4 = "Hello"
        ws.set("F5", Formula::new("CONCAT(\"Hello\",\" \",\"World\")")?)?; // F5 = "Hello World"
        ws.set("F6", Formula::new("TEXTJOIN(\"-\",TRUE,A1:A3)")?)?; // F6 = "1-2-3"
    }

    edit.commit()?.into_workbook().save(path)?;
    println!("Created basic functions test workbook at {}", path);

    Ok(())
}
