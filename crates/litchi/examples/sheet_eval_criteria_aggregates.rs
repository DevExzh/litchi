//! Criteria-based aggregates formula evaluator example.
//!
//! Exercises SUMIF, SUMIFS, COUNTIF, COUNTIFS, AVERAGEIF, and AVERAGEIFS
//! implemented in the shared formula evaluator.
//!
//! Run with:
//!
//! ```bash
//! cargo run --example sheet_eval_criteria_aggregates --features ooxml -- sheet_eval_criteria_aggregates.xlsx
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
        "sheet_eval_criteria_aggregates.xlsx"
    };

    build_sample_xlsx(path)?;

    let xlsx_wb = XlsxWorkbook::open(path)?;
    let evaluator = FormulaEvaluator::new(&xlsx_wb);

    println!("Evaluating criteria-based aggregates on Crit in {}", path);

    // SUMIF / COUNTIF / AVERAGEIF
    for (coord, row, col) in [
        ("E1", 1, 5),
        ("E2", 2, 5),
        ("E3", 3, 5),
        ("E4", 4, 5),
        ("E5", 5, 5),
        ("E6", 6, 5),
    ] {
        let value = evaluator.evaluate_cell("Crit", row, col).await?;
        println!("  {} => {:?}", coord, value);
    }

    // SUMIFS / COUNTIFS / AVERAGEIFS
    for (coord, row, col) in [("F1", 1, 6), ("F2", 2, 6), ("F3", 3, 6)] {
        let value = evaluator.evaluate_cell("Crit", row, col).await?;
        println!("  {} => {:?}", coord, value);
    }

    Ok(())
}

fn build_sample_xlsx(path: &str) -> ExampleResult<()> {
    let mut wb = XlsxWorkbook::create()?;

    wb.add_worksheet("Crit");
    {
        let ws = wb.worksheet_mut(0)?;
        ws.set_name("Crit".to_string());

        // Data table in A1:C6
        // A: values; B: categories; C: flags
        ws.set_cell_value(1, 1, 10); // A1
        ws.set_cell_value(2, 1, 20); // A2
        ws.set_cell_value(3, 1, 5); // A3
        ws.set_cell_value(4, 1, 15); // A4
        ws.set_cell_value(5, 1, 25); // A5
        ws.set_cell_value(6, 1, 30); // A6

        ws.set_cell_value(1, 2, "A"); // B1
        ws.set_cell_value(2, 2, "B"); // B2
        ws.set_cell_value(3, 2, "A"); // B3
        ws.set_cell_value(4, 2, "B"); // B4
        ws.set_cell_value(5, 2, "A"); // B5
        ws.set_cell_value(6, 2, "C"); // B6

        ws.set_cell_value(1, 3, true); // C1
        ws.set_cell_value(2, 3, false); // C2
        ws.set_cell_value(3, 3, true); // C3
        ws.set_cell_value(4, 3, true); // C4
        ws.set_cell_value(5, 3, false); // C5
        ws.set_cell_value(6, 3, true); // C6

        // E1..E6: SUMIF / COUNTIF / AVERAGEIF with various criteria
        ws.set_cell_formula(1, 5, "SUMIF(A1:A6,\">10\")"); // E1: 20+15+25+30 = 90
        ws.set_cell_formula(2, 5, "SUMIF(B1:B6,\"A\",A1:A6)"); // E2: A-rows in A: 10+5+25 = 40
        ws.set_cell_formula(3, 5, "COUNTIF(B1:B6,\"A\")"); // E3: count of "A" = 3
        ws.set_cell_formula(4, 5, "COUNTIF(A1:A6,\">=20\")"); // E4: values >= 20: 20,25,30 = 3
        ws.set_cell_formula(5, 5, "AVERAGEIF(B1:B6,\"A\",A1:A6)"); // E5: avg of A rows: 40/3
        ws.set_cell_formula(6, 5, "AVERAGEIF(A1:A6,\">20\")"); // E6: avg of >20: (25+30)/2 = 27.5

        // F1..F3: SUMIFS / COUNTIFS / AVERAGEIFS
        // SUMIFS over A where B="A" and C=TRUE
        ws.set_cell_formula(1, 6, "SUMIFS(A1:A6,B1:B6,\"A\",C1:C6,TRUE)");
        // COUNTIFS over B and C
        ws.set_cell_formula(2, 6, "COUNTIFS(B1:B6,\"A\",C1:C6,TRUE)");
        // AVERAGEIFS over A with same criteria
        ws.set_cell_formula(3, 6, "AVERAGEIFS(A1:A6,B1:B6,\"A\",C1:C6,TRUE)");
    }

    wb.save(path)?;
    println!("Created criteria aggregates test workbook at {}", path);

    Ok(())
}
