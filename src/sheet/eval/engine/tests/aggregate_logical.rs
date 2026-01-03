#![cfg(all(test, feature = "eval_engine", feature = "ooxml"))]

use crate::ooxml::xlsx::Workbook as XlsxWorkbook;
use crate::sheet::{CellValue, FormulaEvaluator};
use tempfile::tempdir;

const TOL: f64 = 1e-9;

#[tokio::test]
async fn eval_aggregates_and_logical() {
    let dir = tempdir().expect("create temp dir");
    let path = dir.path().join("aggregate_logical.xlsx");
    let path_str = path.to_str().expect("utf-8 path");

    build_workbook(path_str);

    let wb = XlsxWorkbook::open(path_str).expect("open workbook");
    let evaluator = FormulaEvaluator::new(&wb);

    // Aggregate tests
    assert_float(
        evaluator.evaluate_cell("Sheet1", 1, 4).await.expect("SUM"),
        10.0,
    );
    assert_float(
        evaluator
            .evaluate_cell("Sheet1", 2, 4)
            .await
            .expect("AVERAGE"),
        2.5,
    );
    assert_float(
        evaluator
            .evaluate_cell("Sheet1", 3, 4)
            .await
            .expect("COUNT"),
        4.0,
    );
    assert_float(
        evaluator.evaluate_cell("Sheet1", 4, 4).await.expect("MAX"),
        4.0,
    );
    assert_float(
        evaluator.evaluate_cell("Sheet1", 5, 4).await.expect("MIN"),
        1.0,
    );

    // Logical tests
    assert_bool(
        evaluator.evaluate_cell("Sheet1", 1, 6).await.expect("AND"),
        true,
    );
    assert_bool(
        evaluator.evaluate_cell("Sheet1", 2, 6).await.expect("OR"),
        true,
    );
    assert_bool(
        evaluator.evaluate_cell("Sheet1", 3, 6).await.expect("NOT"),
        true,
    );
    match evaluator.evaluate_cell("Sheet1", 4, 6).await.expect("IF") {
        CellValue::String(s) => assert_eq!(s, "yes"),
        other => panic!("Unexpected IF result: {:?}", other),
    }
}

fn build_workbook(path: &str) {
    let mut wb = XlsxWorkbook::create().expect("create workbook");
    wb.add_worksheet("Sheet1");

    {
        let ws = wb.worksheet_mut(0).expect("worksheet 0");
        ws.set_name("Sheet1".to_string());

        // Data for aggregates
        ws.set_cell_value(1, 1, 1);
        ws.set_cell_value(2, 1, 2);
        ws.set_cell_value(3, 1, 3);
        ws.set_cell_value(4, 1, 4);

        // Aggregate formulas
        ws.set_cell_formula(1, 4, "SUM(A1:A4)");
        ws.set_cell_formula(2, 4, "AVERAGE(A1:A4)");
        ws.set_cell_formula(3, 4, "COUNT(A1:A4)");
        ws.set_cell_formula(4, 4, "MAX(A1:A4)");
        ws.set_cell_formula(5, 4, "MIN(A1:A4)");

        // Logical formulas
        ws.set_cell_formula(1, 6, "AND(TRUE, 1=1, A1=1)");
        ws.set_cell_formula(2, 6, "OR(FALSE, 2=2, A4=5)");
        ws.set_cell_formula(3, 6, "NOT(OR(FALSE, FALSE))");
        ws.set_cell_formula(4, 6, "IF(AND(A1=1, A4=4), \"yes\", \"no\")");
    }

    wb.save(path).expect("save workbook");
}

fn assert_float(value: CellValue, expected: f64) {
    match value {
        CellValue::Float(v) => assert!((v - expected).abs() < TOL, "expected {expected}, got {v}"),
        CellValue::Int(v) => assert!(
            ((v as f64) - expected).abs() < TOL,
            "expected {expected}, got {v}"
        ),
        other => panic!("Unexpected numeric value: {:?}", other),
    }
}

fn assert_bool(value: CellValue, expected: bool) {
    match value {
        CellValue::Bool(v) => assert_eq!(v, expected),
        other => panic!("Unexpected bool value: {:?}", other),
    }
}
