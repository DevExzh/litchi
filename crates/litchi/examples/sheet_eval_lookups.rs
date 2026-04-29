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
        "sheet_eval_lookups.xlsx"
    };

    build_sample_xlsx(path)?;

    let xlsx_wb = XlsxWorkbook::open(path)?;
    let evaluator = FormulaEvaluator::new(&xlsx_wb);

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
    let mut wb = XlsxWorkbook::create()?;

    wb.add_worksheet("Lookup");
    {
        let ws = wb.worksheet_mut(0)?;
        ws.set_name("Lookup".to_string());

        // Vertical table A1:C4
        ws.set_cell_value(1, 1, 10); // A1
        ws.set_cell_value(2, 1, 20); // A2
        ws.set_cell_value(3, 1, 30); // A3
        ws.set_cell_value(4, 1, 40); // A4

        ws.set_cell_value(1, 2, "ten"); // B1
        ws.set_cell_value(2, 2, "twenty"); // B2
        ws.set_cell_value(3, 2, "thirty"); // B3
        ws.set_cell_value(4, 2, "forty"); // B4

        ws.set_cell_value(1, 3, "X"); // C1
        ws.set_cell_value(2, 3, "Y"); // C2
        ws.set_cell_value(3, 3, "Z"); // C3
        ws.set_cell_value(4, 3, "W"); // C4

        // Horizontal table A10:D11
        ws.set_cell_value(10, 1, 100); // A10
        ws.set_cell_value(10, 2, 200); // B10
        ws.set_cell_value(10, 3, 300); // C10
        ws.set_cell_value(10, 4, 400); // D10

        ws.set_cell_value(11, 1, "a"); // A11
        ws.set_cell_value(11, 2, "b"); // B11
        ws.set_cell_value(11, 3, "c"); // C11
        ws.set_cell_value(11, 4, "d"); // D11

        // D1..D6: formulas using lookup functions
        // INDEX over B1:B4
        ws.set_cell_formula(1, 4, "INDEX(B1:B4,3)"); // "thirty"
        // MATCH over A1:A4
        ws.set_cell_formula(2, 4, "MATCH(30,A1:A4,0)"); // 3
        // XMATCH over B1:B4
        ws.set_cell_formula(3, 4, "XMATCH(\"twenty\",B1:B4,0)"); // 2
        // VLOOKUP into A1:C4
        ws.set_cell_formula(4, 4, "VLOOKUP(30,A1:C4,3,FALSE)"); // "Z"
        // HLOOKUP into A10:D11
        ws.set_cell_formula(5, 4, "HLOOKUP(200,A10:D11,2,FALSE)"); // "b"
        // XLOOKUP from A1:A4 to C1:C4 with if_not_found
        ws.set_cell_formula(6, 4, "XLOOKUP(25,A1:A4,C1:C4,\"not found\",0)"); // "not found"
    }

    wb.save(path)?;
    println!("Created lookup functions test workbook at {}", path);

    Ok(())
}
