//! Lookup and reference functions formula evaluator example.
//!
//! Exercises INDEX, MATCH, VLOOKUP, HLOOKUP, XLOOKUP, and XMATCH
//! implemented in the shared formula evaluator.
//!
//! Run with:
//!
//! ```bash
//! cargo run --example sheet_eval_lookups --features ooxml -- sheet_eval_lookups.xlsx
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
        "sheet_eval_lookups.xlsx"
    };

    build_sample_xlsx(path)?;

    let xlsx_wb = open_workbook(path)?;
    let evaluator = FormulaEvaluator::new(xlsx_wb.as_ref());

    println!(
        "Evaluating lookup/reference functions on Lookup in {}",
        path
    );

    for (coord, row, col) in [
        ("D1", 1, 4),
        ("D2", 2, 4),
        ("D3", 3, 4),
        ("D4", 4, 4),
        ("D5", 5, 4),
        ("D6", 6, 4),
    ] {
        let value = evaluator.evaluate_cell("Lookup", row, col).await?;
        println!("  {} => {:?}", coord, value);
    }

    Ok(())
}

fn build_sample_xlsx(path: &str) -> ExampleResult<()> {
    let wb = Workbook::create()?;
    let mut edit = wb.edit()?;
    edit.tab(0)?
        .ok_or("missing worksheet tab")?
        .rename("Lookup")?;

    {
        let mut ws = edit.sheet(0)?.ok_or("missing worksheet")?;

        // Vertical table A1:C4
        ws.set("A1", 10_i32)?; // A1
        ws.set("A2", 20_i32)?; // A2
        ws.set("A3", 30_i32)?; // A3
        ws.set("A4", 40_i32)?; // A4

        ws.set("B1", "ten")?; // B1
        ws.set("B2", "twenty")?; // B2
        ws.set("B3", "thirty")?; // B3
        ws.set("B4", "forty")?; // B4

        ws.set("C1", "X")?; // C1
        ws.set("C2", "Y")?; // C2
        ws.set("C3", "Z")?; // C3
        ws.set("C4", "W")?; // C4

        // Horizontal table A10:D11
        ws.set("A10", 100_i32)?; // A10
        ws.set("B10", 200_i32)?; // B10
        ws.set("C10", 300_i32)?; // C10
        ws.set("D10", 400_i32)?; // D10

        ws.set("A11", "a")?; // A11
        ws.set("B11", "b")?; // B11
        ws.set("C11", "c")?; // C11
        ws.set("D11", "d")?; // D11

        // D1..D6: formulas using lookup functions
        // INDEX over B1:B4
        ws.set("D1", Formula::new("INDEX(B1:B4,3)")?)?; // "thirty"
        // MATCH over A1:A4
        ws.set("D2", Formula::new("MATCH(30,A1:A4,0)")?)?; // 3
        // XMATCH over B1:B4
        ws.set("D3", Formula::new("XMATCH(\"twenty\",B1:B4,0)")?)?; // 2
        // VLOOKUP into A1:C4
        ws.set("D4", Formula::new("VLOOKUP(30,A1:C4,3,FALSE)")?)?; // "Z"
        // HLOOKUP into A10:D11
        ws.set("D5", Formula::new("HLOOKUP(200,A10:D11,2,FALSE)")?)?; // "b"
        // XLOOKUP from A1:A4 to C1:C4 with if_not_found
        ws.set(
            "D6",
            Formula::new("XLOOKUP(25,A1:A4,C1:C4,\"not found\",0)")?,
        )?; // "not found"
    }

    edit.commit()?.into_workbook().save(path)?;
    println!("Created lookup functions test workbook at {}", path);

    Ok(())
}
