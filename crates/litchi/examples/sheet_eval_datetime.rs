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
        "sheet_eval_datetime.xlsx"
    };

    build_sample_xlsx(path)?;

    let xlsx_wb = XlsxWorkbook::open(path)?;
    let evaluator = FormulaEvaluator::new(&xlsx_wb);

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
    let mut wb = XlsxWorkbook::create()?;

    wb.add_worksheet("DateTime");
    {
        let ws = wb.worksheet_mut(0)?;
        ws.set_name("DateTime".to_string());

        // Column A: labels
        ws.set_cell_value(1, 1, "TODAY()");
        ws.set_cell_value(2, 1, "NOW()");
        ws.set_cell_value(3, 1, "DATE(2024,1,1)");
        ws.set_cell_value(4, 1, "TIME(12,0,0)");
        ws.set_cell_value(5, 1, "DATEVALUE(\"2024-01-15\")");
        ws.set_cell_value(6, 1, "TIMEVALUE(\"13:45\")");
        ws.set_cell_value(7, 1, "EDATE(DATE(2024,1,1),1)");
        ws.set_cell_value(8, 1, "EOMONTH(DATE(2024,1,1),1)");
        ws.set_cell_value(9, 1, "WORKDAY(DATE(2024,1,1),5)");
        ws.set_cell_value(10, 1, "WORKDAY.INTL(DATE(2024,1,1),5)");
        ws.set_cell_value(11, 1, "NETWORKDAYS(DATE(2024,1,1),EDATE(DATE(2024,1,1),1))");
        ws.set_cell_value(
            12,
            1,
            "NETWORKDAYS(DATE(2024,1,1),EDATE(DATE(2024,1,1),1),WORKDAY(DATE(2024,1,1),5))",
        );

        // Column B: formulas
        ws.set_cell_formula(1, 2, "TODAY()");
        ws.set_cell_formula(2, 2, "NOW()");
        ws.set_cell_formula(3, 2, "DATE(2024,1,1)");
        ws.set_cell_formula(4, 2, "TIME(12,0,0)");
        ws.set_cell_formula(5, 2, "DATEVALUE(\"2024-01-15\")");
        ws.set_cell_formula(6, 2, "TIMEVALUE(\"13:45\")");
        ws.set_cell_formula(7, 2, "EDATE(DATE(2024,1,1),1)");
        ws.set_cell_formula(8, 2, "EOMONTH(DATE(2024,1,1),1)");
        ws.set_cell_formula(9, 2, "WORKDAY(DATE(2024,1,1),5)");
        ws.set_cell_formula(10, 2, "WORKDAY.INTL(DATE(2024,1,1),5)");
        ws.set_cell_formula(11, 2, "NETWORKDAYS(DATE(2024,1,1),EDATE(DATE(2024,1,1),1))");
        ws.set_cell_formula(
            12,
            2,
            "NETWORKDAYS(DATE(2024,1,1),EDATE(DATE(2024,1,1),1),WORKDAY(DATE(2024,1,1),5))",
        );
    }

    wb.save(path)?;
    println!("Created date/time test workbook at {}", path);

    Ok(())
}
