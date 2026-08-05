//! Complex chained formula evaluator example.
//!
//! This example builds a multi-sheet workbook that exercises most of the
//! implemented functions *in chains* across sheets:
//!
//! - Arithmetic and scalar math (operators, ROUND)
//! - Aggregates (SUM, AVERAGE)
//! - Criteria aggregates (SUMIFS, COUNTIFS, AVERAGEIFS, COUNTIF)
//! - Logical functions (IF)
//! - Text functions (CONCAT, TEXTJOIN, LOWER)
//! - Lookup functions (INDEX/XLOOKUP-style via XLOOKUP)
//! - Date/time functions (DATE, EDATE, WORKDAY, NETWORKDAYS)
//!
//! The goal is to stress recursive evaluation, caching, and cross-sheet
//! dependencies in the EngineCtx-based evaluator.
//!
//! Run with:
//!
//! ```bash
//! cargo run --example sheet_eval_complex_chains --features ooxml -- sheet_eval_complex_chains.xlsx
//! ```

use litchi::ooxml::xlsx::{Formula, Value, Workbook};
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
        "sheet_eval_complex_chains.xlsx"
    };

    build_sample_xlsx(path)?;

    let xlsx_wb = open_workbook(path)?;
    let evaluator = FormulaEvaluator::new(xlsx_wb.as_ref());

    println!(
        "Evaluating complex chained formulas across multiple sheets in {}",
        path
    );

    // A selection of cells that together exercise deep, cross-sheet chains.
    // Format: (sheet_name, coord, row, col)
    let targets: &[(&str, &str, u32, u32)] = &[
        // Chains sheet: aggregates that depend on criteria-aggregates on Data.
        ("Chains", "B2", 2, 2), // = Data!F2 (SUMIFS)
        ("Chains", "B3", 3, 2), // = Data!F6 (AVERAGEIFS)
        ("Chains", "G1", 1, 7), // = B2 + B3
        // Chains sheet: lookup + math + logical + arithmetic over other formulas.
        ("Chains", "B2_id1_net", 2, 2), // same as B2 but kept for clarity in output label
        ("Chains", "D2", 2, 4),         // IF over ROUND(XLOOKUP(...))
        ("Chains", "E2", 2, 5),         // arithmetic on IF result
        ("Chains", "G2", 2, 7),         // SUM over D2:D5
        ("Chains", "G3", 3, 7),         // AVERAGE over E2:E5
        // Chains sheet: text chains and date/time chains.
        ("Chains", "H2", 2, 8), // CONCAT + TEXTJOIN over A2:A5
        ("Chains", "I2", 2, 9), // = DatesChain!D2
        ("Chains", "I3", 3, 9), // I2 + G2
        // Data sheet: criteria aggregates and TEXTJOIN.
        ("Data", "F2", 2, 6), // SUMIFS
        ("Data", "F4", 4, 6), // COUNTIFS
        ("Data", "F6", 6, 6), // AVERAGEIFS
        ("Data", "H2", 2, 8), // TEXTJOIN of categories
        // LookupChain sheet: final prices after discounts (driven by formulas).
        ("LookupChain", "G2", 2, 7),
        ("LookupChain", "G4", 4, 7),
        // DatesChain sheet: composed date/time logic.
        ("DatesChain", "A2", 2, 1), // DATE(2024,1,1)
        ("DatesChain", "B2", 2, 2), // EDATE(A2,1)
        ("DatesChain", "C2", 2, 3), // WORKDAY(A2,5)
        ("DatesChain", "D2", 2, 4), // NETWORKDAYS(A2,B2,C2)
    ];

    for (sheet, coord, row, col) in targets {
        let value = evaluator.evaluate_cell(sheet, *row, *col).await?;
        println!("  {}!{} => {:?}", sheet, coord, value);
    }

    Ok(())
}

fn build_sample_xlsx(path: &str) -> ExampleResult<()> {
    let wb = Workbook::create()?;
    let mut edit = wb.edit()?;

    // Sheet 0: Chains - summary and chain-heavy formulas.
    {
        edit.tab(0)?
            .ok_or("missing initial worksheet")?
            .rename("Chains")?;
        let mut ws = edit.sheet(0)?.ok_or("missing Chains worksheet")?;

        // A2..A5: IDs used to drive lookups on LookupChain.
        ws.set("A2", Value::from(1))?;
        ws.set("A3", Value::from(2))?;
        ws.set("A4", Value::from(3))?;
        ws.set("A5", Value::from(4))?;

        // B2..B5: XLOOKUP over final prices on LookupChain!G2:G5.
        ws.set(
            "B2",
            Formula::new("XLOOKUP(A2,LookupChain!A2:A5,LookupChain!G2:G5,\"NA\",0)")?,
        )?;
        ws.set(
            "B3",
            Formula::new("XLOOKUP(A3,LookupChain!A2:A5,LookupChain!G2:G5,\"NA\",0)")?,
        )?;
        ws.set(
            "B4",
            Formula::new("XLOOKUP(A4,LookupChain!A2:A5,LookupChain!G2:G5,\"NA\",0)")?,
        )?;
        ws.set(
            "B5",
            Formula::new("XLOOKUP(A5,LookupChain!A2:A5,LookupChain!G2:G5,\"NA\",0)")?,
        )?;

        // C2..C5: math over lookups.
        ws.set("C2", Formula::new("ROUND(B2*1.2,2)")?)?;
        ws.set("C3", Formula::new("ROUND(B3*1.2,2)")?)?;
        ws.set("C4", Formula::new("ROUND(B4*1.2,2)")?)?;
        ws.set("C5", Formula::new("ROUND(B5*1.2,2)")?)?;

        // D2..D5: logical IF over C, using comparison against a threshold.
        ws.set("D2", Formula::new("IF(C2>100,C2,0)")?)?;
        ws.set("D3", Formula::new("IF(C3>100,C3,0)")?)?;
        ws.set("D4", Formula::new("IF(C4>100,C4,0)")?)?;
        ws.set("D5", Formula::new("IF(C5>100,C5,0)")?)?;

        // E2..E5: arithmetic over IF results.
        ws.set("E2", Formula::new("D2/2")?)?;
        ws.set("E3", Formula::new("D3/2")?)?;
        ws.set("E4", Formula::new("D4/2")?)?;
        ws.set("E5", Formula::new("D5/2")?)?;

        // B2..B3 as summary cells of criteria aggregates on Data.
        ws.set("B1", Value::from("Data SUMIFS"))?;
        ws.set("B2", Formula::new("Data!F2")?)?;
        ws.set("C1", Value::from("Data AVERAGEIFS"))?;
        ws.set("B3", Formula::new("Data!F6")?)?;

        // G1: sum of those two aggregates.
        ws.set("G1", Formula::new("B2+B3")?)?;

        // G2/G3: aggregates over chains on this sheet.
        ws.set("G2", Formula::new("SUM(D2:D5)")?)?;
        ws.set("G3", Formula::new("AVERAGE(E2:E5)")?)?;

        // H2: text chain combining IDs using TEXTJOIN.
        ws.set(
            "H2",
            Formula::new("CONCAT(\"IDs: \",TEXTJOIN(\",\",TRUE,A2:A5))")?,
        )?;

        // I2: pull in NETWORKDAYS result from DatesChain.
        ws.set("I2", Formula::new("DatesChain!D2")?)?;
        // I3: combine date-based result with a local aggregate.
        ws.set("I3", Formula::new("I2+G2")?)?;
    }

    // Sheet 1: Data - criteria aggregates and text chains.
    {
        let mut ws = edit.add("Data")?;

        // A2..A6: numeric values.
        ws.set("A2", Value::from(10))?;
        ws.set("A3", Value::from(20))?;
        ws.set("A4", Value::from(30))?;
        ws.set("A5", Value::from(40))?;
        ws.set("A6", Value::from(50))?;

        // B2..B6: categories.
        ws.set("B2", Value::from("A"))?;
        ws.set("B3", Value::from("B"))?;
        ws.set("B4", Value::from("A"))?;
        ws.set("B5", Value::from("C"))?;
        ws.set("B6", Value::from("A"))?;

        // C2..C6: boolean flags.
        ws.set("C2", Value::from(true))?;
        ws.set("C3", Value::from(false))?;
        ws.set("C4", Value::from(true))?;
        ws.set("C5", Value::from(true))?;
        ws.set("C6", Value::from(false))?;

        // Criteria-based aggregates over the table.
        ws.set("F2", Formula::new("SUMIFS(A2:A6,B2:B6,\"A\",C2:C6,TRUE)")?)?;
        ws.set("F4", Formula::new("COUNTIFS(B2:B6,\"A\",C2:C6,TRUE)")?)?;
        ws.set(
            "F6",
            Formula::new("AVERAGEIFS(A2:A6,B2:B6,\"A\",C2:C6,TRUE)")?,
        )?;

        // H2: joined categories as a text chain.
        ws.set("H2", Formula::new("TEXTJOIN(\"-\",TRUE,B2:B6)")?)?;
    }

    // Sheet 2: LookupChain - line items and discounts used by XLOOKUP.
    {
        let mut ws = edit.add("LookupChain")?;

        // IDs.
        ws.set("A2", Value::from(1))?;
        ws.set("A3", Value::from(2))?;
        ws.set("A4", Value::from(3))?;
        ws.set("A5", Value::from(4))?;

        // Categories.
        ws.set("B2", Value::from("A"))?;
        ws.set("B3", Value::from("B"))?;
        ws.set("B4", Value::from("C"))?;
        ws.set("B5", Value::from("A"))?;

        // Base prices.
        ws.set("C2", Value::from(50))?;
        ws.set("C3", Value::from(75))?;
        ws.set("C4", Value::from(100))?;
        ws.set("C5", Value::from(150))?;

        // Quantities.
        ws.set("D2", Value::from(1))?;
        ws.set("D3", Value::from(2))?;
        ws.set("D4", Value::from(3))?;
        ws.set("D5", Value::from(4))?;

        // Line totals E2..E5 = price * qty.
        ws.set("E2", Formula::new("C2*D2")?)?;
        ws.set("E3", Formula::new("C3*D3")?)?;
        ws.set("E4", Formula::new("C4*D4")?)?;
        ws.set("E5", Formula::new("C5*D5")?)?;

        // Discount rate F2..F5: 10% if line total is above a threshold.
        ws.set("F2", Formula::new("IF(E2>100,0.1,0)")?)?;
        ws.set("F3", Formula::new("IF(E3>100,0.1,0)")?)?;
        ws.set("F4", Formula::new("IF(E4>100,0.1,0)")?)?;
        ws.set("F5", Formula::new("IF(E5>100,0.1,0)")?)?;

        // Final price G2..G5 = E * (1 - F).
        ws.set("G2", Formula::new("E2*(1-F2)")?)?;
        ws.set("G3", Formula::new("E3*(1-F3)")?)?;
        ws.set("G4", Formula::new("E4*(1-F4)")?)?;
        ws.set("G5", Formula::new("E5*(1-F5)")?)?;
    }

    // Sheet 3: DatesChain - composed date/time logic.
    {
        let mut ws = edit.add("DatesChain")?;

        // A2: base date; B2: one month later; C2: 5th workday; D2: network days.
        ws.set("A2", Formula::new("DATE(2024,1,1)")?)?;
        ws.set("B2", Formula::new("EDATE(A2,1)")?)?;
        ws.set("C2", Formula::new("WORKDAY(A2,5)")?)?;
        ws.set("D2", Formula::new("NETWORKDAYS(A2,B2,C2)")?)?;
    }

    edit.commit()?.into_workbook().save(path)?;
    println!("Created complex chains test workbook at {}", path);

    Ok(())
}
