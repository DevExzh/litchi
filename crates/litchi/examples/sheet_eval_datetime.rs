//! Date/time functions formula evaluator example.
//!
//! This example exercises Phase 4 date and time functions implemented in the
//! shared formula evaluator.
//!
//! Run with:
//!
//! ```bash
//! cargo run --example sheet_eval_datetime --features ooxml -- sheet_eval_datetime.xlsx
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
        "sheet_eval_datetime.xlsx"
    };

    build_sample_xlsx(path)?;

    let xlsx_wb = open_workbook(path)?;
    let evaluator = FormulaEvaluator::new(xlsx_wb.as_ref());

    println!("Evaluating date/time functions on DateTime in {}", path);

    for (coord, row, col) in [
        ("B1", 1, 2),
        ("B2", 2, 2),
        ("B3", 3, 2),
        ("B4", 4, 2),
        ("B5", 5, 2),
        ("B6", 6, 2),
        ("B7", 7, 2),
        ("B8", 8, 2),
        ("B9", 9, 2),
        ("B10", 10, 2),
        ("B11", 11, 2),
        ("B12", 12, 2),
    ] {
        let value = evaluator.evaluate_cell("DateTime", row, col).await?;
        println!("  {} => {:?}", coord, value);
    }

    Ok(())
}

fn build_sample_xlsx(path: &str) -> ExampleResult<()> {
    let wb = Workbook::create()?;
    let mut edit = wb.edit()?;
    edit.tab(0)?
        .ok_or("missing worksheet tab")?
        .rename("DateTime")?;

    {
        let mut ws = edit.sheet(0)?.ok_or("missing worksheet")?;

        // Column A: labels
        ws.set("A1", "TODAY()")?;
        ws.set("A2", "NOW()")?;
        ws.set("A3", "DATE(2024,1,1)")?;
        ws.set("A4", "TIME(12,0,0)")?;
        ws.set("A5", "DATEVALUE(\"2024-01-15\")")?;
        ws.set("A6", "TIMEVALUE(\"13:45\")")?;
        ws.set("A7", "EDATE(DATE(2024,1,1),1)")?;
        ws.set("A8", "EOMONTH(DATE(2024,1,1),1)")?;
        ws.set("A9", "WORKDAY(DATE(2024,1,1),5)")?;
        ws.set("A10", "WORKDAY.INTL(DATE(2024,1,1),5)")?;
        ws.set("A11", "NETWORKDAYS(DATE(2024,1,1),EDATE(DATE(2024,1,1),1))")?;
        ws.set(
            "A12",
            "NETWORKDAYS(DATE(2024,1,1),EDATE(DATE(2024,1,1),1),WORKDAY(DATE(2024,1,1),5))",
        )?;

        // Column B: formulas
        ws.set("B1", Formula::new("TODAY()")?)?;
        ws.set("B2", Formula::new("NOW()")?)?;
        ws.set("B3", Formula::new("DATE(2024,1,1)")?)?;
        ws.set("B4", Formula::new("TIME(12,0,0)")?)?;
        ws.set("B5", Formula::new("DATEVALUE(\"2024-01-15\")")?)?;
        ws.set("B6", Formula::new("TIMEVALUE(\"13:45\")")?)?;
        ws.set("B7", Formula::new("EDATE(DATE(2024,1,1),1)")?)?;
        ws.set("B8", Formula::new("EOMONTH(DATE(2024,1,1),1)")?)?;
        ws.set("B9", Formula::new("WORKDAY(DATE(2024,1,1),5)")?)?;
        ws.set("B10", Formula::new("WORKDAY.INTL(DATE(2024,1,1),5)")?)?;
        ws.set(
            "B11",
            Formula::new("NETWORKDAYS(DATE(2024,1,1),EDATE(DATE(2024,1,1),1))")?,
        )?;
        ws.set(
            "B12",
            Formula::new(
                "NETWORKDAYS(DATE(2024,1,1),EDATE(DATE(2024,1,1),1),WORKDAY(DATE(2024,1,1),5))",
            )?,
        )?;
    }

    edit.commit()?.into_workbook().save(path)?;
    println!("Created date/time test workbook at {}", path);

    Ok(())
}
