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
        "sheet_eval_complex_chains.xlsx"
    };

    build_sample_xlsx(path)?;

    let xlsx_wb = XlsxWorkbook::open(path)?;
    let evaluator = FormulaEvaluator::new(&xlsx_wb);

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
    let mut wb = XlsxWorkbook::create()?;

    // Sheet 0: Chains - summary and chain-heavy formulas.
    {
        let ws = wb.worksheet_mut(0)?;
        ws.set_name("Chains".to_string());

        // A2..A5: IDs used to drive lookups on LookupChain.
        ws.set_cell_value(2, 1, 1); // A2
        ws.set_cell_value(3, 1, 2); // A3
        ws.set_cell_value(4, 1, 3); // A4
        ws.set_cell_value(5, 1, 4); // A5

        // B2..B5: XLOOKUP over final prices on LookupChain!G2:G5.
        ws.set_cell_formula(
            2,
            2,
            "XLOOKUP(A2,LookupChain!A2:A5,LookupChain!G2:G5,\"NA\",0)",
        );
        ws.set_cell_formula(
            3,
            2,
            "XLOOKUP(A3,LookupChain!A2:A5,LookupChain!G2:G5,\"NA\",0)",
        );
        ws.set_cell_formula(
            4,
            2,
            "XLOOKUP(A4,LookupChain!A2:A5,LookupChain!G2:G5,\"NA\",0)",
        );
        ws.set_cell_formula(
            5,
            2,
            "XLOOKUP(A5,LookupChain!A2:A5,LookupChain!G2:G5,\"NA\",0)",
        );

        // C2..C5: math over lookups.
        ws.set_cell_formula(2, 3, "ROUND(B2*1.2,2)");
        ws.set_cell_formula(3, 3, "ROUND(B3*1.2,2)");
        ws.set_cell_formula(4, 3, "ROUND(B4*1.2,2)");
        ws.set_cell_formula(5, 3, "ROUND(B5*1.2,2)");

        // D2..D5: logical IF over C, using comparison against a threshold.
        ws.set_cell_formula(2, 4, "IF(C2>100,C2,0)");
        ws.set_cell_formula(3, 4, "IF(C3>100,C3,0)");
        ws.set_cell_formula(4, 4, "IF(C4>100,C4,0)");
        ws.set_cell_formula(5, 4, "IF(C5>100,C5,0)");

        // E2..E5: arithmetic over IF results.
        ws.set_cell_formula(2, 5, "D2/2");
        ws.set_cell_formula(3, 5, "D3/2");
        ws.set_cell_formula(4, 5, "D4/2");
        ws.set_cell_formula(5, 5, "D5/2");

        // B2..B3 as summary cells of criteria aggregates on Data.
        ws.set_cell_value(1, 2, "Data SUMIFS");
        ws.set_cell_formula(2, 2, "Data!F2");
        ws.set_cell_value(1, 3, "Data AVERAGEIFS");
        ws.set_cell_formula(3, 2, "Data!F6");

        // G1: sum of those two aggregates.
        ws.set_cell_formula(1, 7, "B2+B3");

        // G2/G3: aggregates over chains on this sheet.
        ws.set_cell_formula(2, 7, "SUM(D2:D5)");
        ws.set_cell_formula(3, 7, "AVERAGE(E2:E5)");

        // H2: text chain combining IDs using TEXTJOIN.
        ws.set_cell_formula(2, 8, "CONCAT(\"IDs: \",TEXTJOIN(\",\",TRUE,A2:A5))");

        // I2: pull in NETWORKDAYS result from DatesChain.
        ws.set_cell_formula(2, 9, "DatesChain!D2");
        // I3: combine date-based result with a local aggregate.
        ws.set_cell_formula(3, 9, "I2+G2");
    }

    // Sheet 1: Data - criteria aggregates and text chains.
    wb.add_worksheet("Data");
    {
        let ws = wb.worksheet_mut(1)?;
        ws.set_name("Data".to_string());

        // A2..A6: numeric values.
        ws.set_cell_value(2, 1, 10); // A2
        ws.set_cell_value(3, 1, 20); // A3
        ws.set_cell_value(4, 1, 30); // A4
        ws.set_cell_value(5, 1, 40); // A5
        ws.set_cell_value(6, 1, 50); // A6

        // B2..B6: categories.
        ws.set_cell_value(2, 2, "A");
        ws.set_cell_value(3, 2, "B");
        ws.set_cell_value(4, 2, "A");
        ws.set_cell_value(5, 2, "C");
        ws.set_cell_value(6, 2, "A");

        // C2..C6: boolean flags.
        ws.set_cell_value(2, 3, true);
        ws.set_cell_value(3, 3, false);
        ws.set_cell_value(4, 3, true);
        ws.set_cell_value(5, 3, true);
        ws.set_cell_value(6, 3, false);

        // Criteria-based aggregates over the table.
        ws.set_cell_value(2, 6, "SUMIFS A & TRUE");
        ws.set_cell_formula(3, 6, "SUMIFS(A2:A6,B2:B6,\"A\",C2:C6,TRUE)");

        ws.set_cell_value(4, 6, "COUNTIFS A & TRUE");
        ws.set_cell_formula(5, 6, "COUNTIFS(B2:B6,\"A\",C2:C6,TRUE)");

        ws.set_cell_value(6, 6, "AVERAGEIFS A & TRUE");
        ws.set_cell_formula(7, 6, "AVERAGEIFS(A2:A6,B2:B6,\"A\",C2:C6,TRUE)");

        // For convenience, mirror the useful numeric results into F2/F4/F6.
        ws.set_cell_formula(2, 6, "SUMIFS(A2:A6,B2:B6,\"A\",C2:C6,TRUE)"); // F2
        ws.set_cell_formula(4, 6, "COUNTIFS(B2:B6,\"A\",C2:C6,TRUE)"); // F4
        ws.set_cell_formula(6, 6, "AVERAGEIFS(A2:A6,B2:B6,\"A\",C2:C6,TRUE)"); // F6

        // H2: joined categories as a text chain.
        ws.set_cell_formula(2, 8, "TEXTJOIN(\"-\",TRUE,B2:B6)");
    }

    // Sheet 2: LookupChain - line items and discounts used by XLOOKUP.
    wb.add_worksheet("LookupChain");
    {
        let ws = wb.worksheet_mut(2)?;
        ws.set_name("LookupChain".to_string());

        // IDs.
        ws.set_cell_value(2, 1, 1); // A2
        ws.set_cell_value(3, 1, 2); // A3
        ws.set_cell_value(4, 1, 3); // A4
        ws.set_cell_value(5, 1, 4); // A5

        // Categories.
        ws.set_cell_value(2, 2, "A");
        ws.set_cell_value(3, 2, "B");
        ws.set_cell_value(4, 2, "C");
        ws.set_cell_value(5, 2, "A");

        // Base prices.
        ws.set_cell_value(2, 3, 50); // C2
        ws.set_cell_value(3, 3, 75); // C3
        ws.set_cell_value(4, 3, 100); // C4
        ws.set_cell_value(5, 3, 150); // C5

        // Quantities.
        ws.set_cell_value(2, 4, 1); // D2
        ws.set_cell_value(3, 4, 2); // D3
        ws.set_cell_value(4, 4, 3); // D4
        ws.set_cell_value(5, 4, 4); // D5

        // Line totals E2..E5 = price * qty.
        ws.set_cell_formula(2, 5, "C2*D2");
        ws.set_cell_formula(3, 5, "C3*D3");
        ws.set_cell_formula(4, 5, "C4*D4");
        ws.set_cell_formula(5, 5, "C5*D5");

        // Discount rate F2..F5: 10% if line total is above a threshold.
        ws.set_cell_formula(2, 6, "IF(E2>100,0.1,0)");
        ws.set_cell_formula(3, 6, "IF(E3>100,0.1,0)");
        ws.set_cell_formula(4, 6, "IF(E4>100,0.1,0)");
        ws.set_cell_formula(5, 6, "IF(E5>100,0.1,0)");

        // Final price G2..G5 = E * (1 - F).
        ws.set_cell_formula(2, 7, "E2*(1-F2)");
        ws.set_cell_formula(3, 7, "E3*(1-F3)");
        ws.set_cell_formula(4, 7, "E4*(1-F4)");
        ws.set_cell_formula(5, 7, "E5*(1-F5)");
    }

    // Sheet 3: DatesChain - composed date/time logic.
    wb.add_worksheet("DatesChain");
    {
        let ws = wb.worksheet_mut(3)?;
        ws.set_name("DatesChain".to_string());

        // A2: base date; B2: one month later; C2: 5th workday; D2: network days.
        ws.set_cell_formula(2, 1, "DATE(2024,1,1)");
        ws.set_cell_formula(2, 2, "EDATE(A2,1)");
        ws.set_cell_formula(2, 3, "WORKDAY(A2,5)");
        ws.set_cell_formula(2, 4, "NETWORKDAYS(A2,B2,C2)");
    }

    wb.save(path)?;
    println!("Created complex chains test workbook at {}", path);

    Ok(())
}
