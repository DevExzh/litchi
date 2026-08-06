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
        "sheet_eval_criteria_aggregates.xlsx"
    };

    build_sample_xlsx(path)?;

    let xlsx_wb = open_workbook(path)?;
    let evaluator = FormulaEvaluator::new(xlsx_wb.as_ref());

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
    let wb = Workbook::create()?;
    let mut edit = wb.edit()?;
    edit.tab(0)?
        .ok_or("missing worksheet tab")?
        .rename("Crit")?;

    {
        let mut ws = edit.sheet(0)?.ok_or("missing worksheet")?;

        // Data table in A1:C6
        // A: values; B: categories; C: flags
        ws.set("A1", 10_i32)?; // A1
        ws.set("A2", 20_i32)?; // A2
        ws.set("A3", 5_i32)?; // A3
        ws.set("A4", 15_i32)?; // A4
        ws.set("A5", 25_i32)?; // A5
        ws.set("A6", 30_i32)?; // A6

        ws.set("B1", "A")?; // B1
        ws.set("B2", "B")?; // B2
        ws.set("B3", "A")?; // B3
        ws.set("B4", "B")?; // B4
        ws.set("B5", "A")?; // B5
        ws.set("B6", "C")?; // B6

        ws.set("C1", true)?; // C1
        ws.set("C2", false)?; // C2
        ws.set("C3", true)?; // C3
        ws.set("C4", true)?; // C4
        ws.set("C5", false)?; // C5
        ws.set("C6", true)?; // C6

        // E1..E6: SUMIF / COUNTIF / AVERAGEIF with various criteria
        ws.set("E1", Formula::new("SUMIF(A1:A6,\">10\")")?)?; // E1: 20+15+25+30 = 90
        ws.set("E2", Formula::new("SUMIF(B1:B6,\"A\",A1:A6)")?)?; // E2: A-rows in A: 10+5+25 = 40
        ws.set("E3", Formula::new("COUNTIF(B1:B6,\"A\")")?)?; // E3: count of "A" = 3
        ws.set("E4", Formula::new("COUNTIF(A1:A6,\">=20\")")?)?; // E4: values >= 20: 20,25,30 = 3
        ws.set("E5", Formula::new("AVERAGEIF(B1:B6,\"A\",A1:A6)")?)?; // E5: avg of A rows: 40/3
        ws.set("E6", Formula::new("AVERAGEIF(A1:A6,\">20\")")?)?; // E6: avg of >20: (25+30)/2 = 27.5

        // F1..F3: SUMIFS / COUNTIFS / AVERAGEIFS
        // SUMIFS over A where B="A" and C=TRUE
        ws.set("F1", Formula::new("SUMIFS(A1:A6,B1:B6,\"A\",C1:C6,TRUE)")?)?;
        // COUNTIFS over B and C
        ws.set("F2", Formula::new("COUNTIFS(B1:B6,\"A\",C1:C6,TRUE)")?)?;
        // AVERAGEIFS over A with same criteria
        ws.set(
            "F3",
            Formula::new("AVERAGEIFS(A1:A6,B1:B6,\"A\",C1:C6,TRUE)")?,
        )?;
    }

    edit.commit()?.into_workbook().save(path)?;
    println!("Created criteria aggregates test workbook at {}", path);

    Ok(())
}
