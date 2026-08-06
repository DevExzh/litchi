//! Statistical and financial functions formula evaluator example.
//!
//! This example exercises statistical distribution functions (NORM.*, CHISQ.*, T.*, F.*)
//! and core financial functions (PV, FV, RATE, NPV, IRR, XNPV, XIRR) implemented in the
//! shared formula evaluator.
//!
//! Run with:
//!
//! ```bash
//! cargo run --example sheet_eval_stat_financial --features ooxml -- sheet_eval_stat_financial.xlsx
//! ```

use litchi::sheet::{FormulaEvaluator, functions::open_workbook};
use litchi::xlsx::{Formula, Number, Workbook};
use std::env;
use std::error::Error;

type ExampleResult<T> = Result<T, Box<dyn Error + Send + Sync>>;

#[tokio::main]
async fn main() -> ExampleResult<()> {
    let args: Vec<String> = env::args().collect();
    let path = if args.len() > 1 {
        args[1].as_str()
    } else {
        "sheet_eval_stat_financial.xlsx"
    };

    build_sample_xlsx(path)?;

    let xlsx_wb = open_workbook(path)?;
    let evaluator = FormulaEvaluator::new(xlsx_wb.as_ref());

    println!("Evaluating statistical and financial functions in {}", path);

    for (sheet, coord, row, col) in [
        ("Stat", "B1", 1, 2),
        ("Stat", "B2", 2, 2),
        ("Stat", "B3", 3, 2),
        ("Stat", "B4", 4, 2),
        ("Stat", "B5", 5, 2),
        ("Stat", "B6", 6, 2),
        ("Stat", "B7", 7, 2),
        ("Fin", "B1", 1, 2),
        ("Fin", "B2", 2, 2),
        ("Fin", "B3", 3, 2),
        ("Fin", "B4", 4, 2),
        ("Fin", "B5", 5, 2),
        ("Fin", "B6", 6, 2),
        ("Fin", "B7", 7, 2),
        ("Fin", "B8", 8, 2),
        ("Fin", "B9", 9, 2),
        ("Fin", "B10", 10, 2),
    ] {
        let value = evaluator.evaluate_cell(sheet, row, col).await?;
        println!("  {}!{} => {:?}", sheet, coord, value);
    }

    Ok(())
}

fn build_sample_xlsx(path: &str) -> ExampleResult<()> {
    let wb = Workbook::create()?;
    let mut edit = wb.edit()?;

    // Statistical functions sheet
    {
        edit.tab(0)?
            .ok_or("missing worksheet tab")?
            .rename("Stat")?;
        let mut ws = edit.sheet(0)?.ok_or("missing worksheet")?;

        // Column A: labels
        ws.set("A1", "NORM.DIST(0,0,1,TRUE)")?;
        ws.set("A2", "NORM.S.INV(0.5)")?;
        ws.set("A3", "CHISQ.DIST(2,4,TRUE)")?;
        ws.set("A4", "CHISQ.INV(0.95,4)")?;
        ws.set("A5", "T.DIST.2T(1,10)")?;
        ws.set("A6", "F.DIST.RT(1.5,4,6)")?;
        ws.set("A7", "PROB(C1:C3,D1:D3,2,3)")?;

        // Column B: formulas
        ws.set("B1", Formula::new("NORM.DIST(0,0,1,TRUE)")?)?;
        ws.set("B2", Formula::new("NORM.S.INV(0.5)")?)?;
        ws.set("B3", Formula::new("CHISQ.DIST(2,4,TRUE)")?)?;
        ws.set("B4", Formula::new("CHISQ.INV(0.95,4)")?)?;
        ws.set("B5", Formula::new("T.DIST.2T(1,10)")?)?;
        ws.set("B6", Formula::new("F.DIST.RT(1.5,4,6)")?)?;
        ws.set("B7", Formula::new("PROB(C1:C3,D1:D3,2,3)")?)?;

        // Supporting data for PROB: x values in C1:C3, probabilities in D1:D3
        ws.set("C1", Number::new("1.0")?)?; // C1
        ws.set("C2", Number::new("2.0")?)?; // C2
        ws.set("C3", Number::new("3.0")?)?; // C3

        ws.set("D1", Number::new("0.2")?)?; // D1
        ws.set("D2", Number::new("0.5")?)?; // D2
        ws.set("D3", Number::new("0.3")?)?; // D3
    }

    // Financial functions sheet
    let fin = edit.add("Fin")?;
    {
        let mut ws = fin;

        // Column A: labels
        ws.set("A1", "PV(0,3,100,0,0)")?;
        ws.set("A2", "FV(0,3,100,0,0)")?;
        ws.set("A3", "RATE(3,-100,250,0,0,0.1)")?;
        ws.set("A4", "NPV(0,10,20,30)")?;
        ws.set("A5", "IRR(C1:C2,0.1)")?;
        ws.set("A6", "XNPV(0.1,D1:D2,E1:E2)")?;
        ws.set("A7", "XIRR(D1:D2,E1:E2,0.1)")?;
        ws.set("A8", "PRODUCT(F1:F3)")?;
        ws.set("A9", "YIELD(0,365,0.05,95,100,2,0)")?;
        ws.set("A10", "DURATION(0,365,0.05,0.10,2,0)")?;

        // Supporting cash flows and dates
        // IRR cash flows
        ws.set("C1", Number::new("-100.0")?)?; // C1
        ws.set("C2", Number::new("110.0")?)?; // C2

        // XNPV/XIRR cash flows and dates (same pattern)
        ws.set("D1", Number::new("-100.0")?)?; // D1
        ws.set("D2", Number::new("110.0")?)?; // D2

        // Dates as numeric serials (relative year apart)
        ws.set("E1", Number::new("0.0")?)?; // E1: base date
        ws.set("E2", Number::new("365.0")?)?; // E2: one year later

        // Values for PRODUCT in F1:F3
        ws.set("F1", Number::new("2.0")?)?; // F1
        ws.set("F2", Number::new("3.0")?)?; // F2
        ws.set("F3", Number::new("4.0")?)?; // F3

        // Column B: formulas
        ws.set("B1", Formula::new("PV(0,3,100,0,0)")?)?;
        ws.set("B2", Formula::new("FV(0,3,100,0,0)")?)?;
        ws.set("B3", Formula::new("RATE(3,-100,250,0,0,0.1)")?)?;
        ws.set("B4", Formula::new("NPV(0,10,20,30)")?)?;
        ws.set("B5", Formula::new("IRR(C1:C2,0.1)")?)?;
        ws.set("B6", Formula::new("XNPV(0.1,D1:D2,E1:E2)")?)?;
        ws.set("B7", Formula::new("XIRR(D1:D2,E1:E2,0.1)")?)?;
        ws.set("B8", Formula::new("PRODUCT(F1:F3)")?)?;
        ws.set("B9", Formula::new("YIELD(0,365,0.05,95,100,2,0)")?)?;
        ws.set("B10", Formula::new("DURATION(0,365,0.05,0.10,2,0)")?)?;
    }

    edit.commit()?.into_workbook().save(path)?;
    println!("Created statistical/financial test workbook at {}", path);

    Ok(())
}
