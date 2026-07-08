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
        "sheet_eval_stat_financial.xlsx"
    };

    build_sample_xlsx(path)?;

    let xlsx_wb = XlsxWorkbook::open(path)?;
    let evaluator = FormulaEvaluator::new(&xlsx_wb);

    println!("Evaluating statistical and financial functions in {}", path);

    for (sheet, coord, row, col) in [
        ("Stats", "B1", 1, 2),
        ("Stats", "B2", 2, 2),
        ("Stats", "B3", 3, 2),
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
    let mut wb = XlsxWorkbook::create()?;

    // Statistical functions sheet
    wb.add_worksheet("Stat");
    {
        let ws = wb.worksheet_mut(0)?;
        ws.set_name("Stat".to_string());

        // Column A: labels
        ws.set_cell_value(1, 1, "NORM.DIST(0,0,1,TRUE)");
        ws.set_cell_value(2, 1, "NORM.S.INV(0.5)");
        ws.set_cell_value(3, 1, "CHISQ.DIST(2,4,TRUE)");
        ws.set_cell_value(4, 1, "CHISQ.INV(0.95,4)");
        ws.set_cell_value(5, 1, "T.DIST.2T(1,10)");
        ws.set_cell_value(6, 1, "F.DIST.RT(1.5,4,6)");
        ws.set_cell_value(7, 1, "PROB(C1:C3,D1:D3,2,3)");

        // Column B: formulas
        ws.set_cell_formula(1, 2, "NORM.DIST(0,0,1,TRUE)");
        ws.set_cell_formula(2, 2, "NORM.S.INV(0.5)");
        ws.set_cell_formula(3, 2, "CHISQ.DIST(2,4,TRUE)");
        ws.set_cell_formula(4, 2, "CHISQ.INV(0.95,4)");
        ws.set_cell_formula(5, 2, "T.DIST.2T(1,10)");
        ws.set_cell_formula(6, 2, "F.DIST.RT(1.5,4,6)");
        ws.set_cell_formula(7, 2, "PROB(C1:C3,D1:D3,2,3)");

        // Supporting data for PROB: x values in C1:C3, probabilities in D1:D3
        ws.set_cell_value(1, 3, 1.0_f64); // C1
        ws.set_cell_value(2, 3, 2.0_f64); // C2
        ws.set_cell_value(3, 3, 3.0_f64); // C3

        ws.set_cell_value(1, 4, 0.2_f64); // D1
        ws.set_cell_value(2, 4, 0.5_f64); // D2
        ws.set_cell_value(3, 4, 0.3_f64); // D3
    }

    // Financial functions sheet
    wb.add_worksheet("Fin");
    {
        let ws = wb.worksheet_mut(1)?;
        ws.set_name("Fin".to_string());

        // Column A: labels
        ws.set_cell_value(1, 1, "PV(0,3,100,0,0)");
        ws.set_cell_value(2, 1, "FV(0,3,100,0,0)");
        ws.set_cell_value(3, 1, "RATE(3,-100,250,0,0,0.1)");
        ws.set_cell_value(4, 1, "NPV(0,10,20,30)");
        ws.set_cell_value(5, 1, "IRR(C1:C2,0.1)");
        ws.set_cell_value(6, 1, "XNPV(0.1,D1:D2,E1:E2)");
        ws.set_cell_value(7, 1, "XIRR(D1:D2,E1:E2,0.1)");
        ws.set_cell_value(8, 1, "PRODUCT(F1:F3)");
        ws.set_cell_value(9, 1, "YIELD(0,365,0.05,95,100,2,0)");
        ws.set_cell_value(10, 1, "DURATION(0,365,0.05,0.10,2,0)");

        // Supporting cash flows and dates
        // IRR cash flows
        ws.set_cell_value(1, 3, -100.0_f64); // C1
        ws.set_cell_value(2, 3, 110.0_f64); // C2

        // XNPV/XIRR cash flows and dates (same pattern)
        ws.set_cell_value(1, 4, -100.0_f64); // D1
        ws.set_cell_value(2, 4, 110.0_f64); // D2

        // Dates as numeric serials (relative year apart)
        ws.set_cell_value(1, 5, 0.0_f64); // E1: base date
        ws.set_cell_value(2, 5, 365.0_f64); // E2: one year later

        // Values for PRODUCT in F1:F3
        ws.set_cell_value(1, 6, 2.0_f64); // F1
        ws.set_cell_value(2, 6, 3.0_f64); // F2
        ws.set_cell_value(3, 6, 4.0_f64); // F3

        // Column B: formulas
        ws.set_cell_formula(1, 2, "PV(0,3,100,0,0)");
        ws.set_cell_formula(2, 2, "FV(0,3,100,0,0)");
        ws.set_cell_formula(3, 2, "RATE(3,-100,250,0,0,0.1)");
        ws.set_cell_formula(4, 2, "NPV(0,10,20,30)");
        ws.set_cell_formula(5, 2, "IRR(C1:C2,0.1)");
        ws.set_cell_formula(6, 2, "XNPV(0.1,D1:D2,E1:E2)");
        ws.set_cell_formula(7, 2, "XIRR(D1:D2,E1:E2,0.1)");
        ws.set_cell_formula(8, 2, "PRODUCT(F1:F3)");
        ws.set_cell_formula(9, 2, "YIELD(0,365,0.05,95,100,2,0)");
        ws.set_cell_formula(10, 2, "DURATION(0,365,0.05,0.10,2,0)");
    }

    wb.save(path)?;
    println!("Created statistical/financial test workbook at {}", path);

    Ok(())
}
