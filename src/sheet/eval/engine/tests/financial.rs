#![cfg(all(test, feature = "eval_engine", feature = "ooxml"))]

use crate::ooxml::xlsx::Workbook as XlsxWorkbook;
use crate::sheet::{CellValue, FormulaEvaluator};
use tempfile::tempdir;

const TOL: f64 = 1e-9;

#[tokio::test]
async fn eval_pduration_and_rri() {
    let dir = tempdir().expect("create temp dir");
    let path = dir.path().join("financial.xlsx");
    let path_str = path.to_str().expect("utf-8 path");

    build_financial_workbook(path_str);

    let wb = XlsxWorkbook::open(path_str).expect("open workbook");
    let evaluator = FormulaEvaluator::new(&wb);

    // PDURATION(rate, pv, fv)
    match evaluator
        .evaluate_cell("Financial", 1, 2)
        .await
        .expect("eval B1 (PDURATION)")
    {
        CellValue::Float(v) => assert!((v - 7.272_540_897_341_713).abs() < TOL),
        other => panic!("Unexpected value for B1: {:?}", other),
    }

    // RRI(nper, pv, fv)
    match evaluator
        .evaluate_cell("Financial", 2, 2)
        .await
        .expect("eval B2 (RRI)")
    {
        CellValue::Float(v) => assert!((v - 0.071_773_462_536_293_14).abs() < TOL),
        other => panic!("Unexpected value for B2: {:?}", other),
    }
}

fn build_financial_workbook(path: &str) {
    let mut wb = XlsxWorkbook::create().expect("create workbook");
    wb.add_worksheet("Financial");

    {
        let ws = wb.worksheet_mut(0).expect("worksheet 0");
        ws.set_name("Financial".to_string());

        // Base inputs for PDURATION and RRI
        ws.set_cell_value(1, 1, 0.1); // A1 rate
        ws.set_cell_value(2, 1, 10_i32); // A2 nper
        ws.set_cell_value(1, 3, 1000_i32); // C1 pv
        ws.set_cell_value(1, 4, 2000_i32); // D1 fv

        // Formulas under test
        ws.set_cell_formula(1, 2, "PDURATION(A1, C1, D1)");
        ws.set_cell_formula(2, 2, "RRI(A2, C1, D1)");
    }

    wb.save(path).expect("save workbook");
}
