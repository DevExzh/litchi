#![cfg(all(test, feature = "eval_engine", feature = "ooxml"))]

use crate::ooxml::xlsx::Workbook as XlsxWorkbook;
use crate::sheet::{CellValue, FormulaEvaluator};
use tempfile::tempdir;

#[tokio::test]
async fn eval_lookup_and_text_functions() {
    let dir = tempdir().expect("create temp dir");
    let path = dir.path().join("lookup_text.xlsx");
    let path_str = path.to_str().expect("utf-8 path");

    build_workbook(path_str);

    let wb = XlsxWorkbook::open(path_str).expect("open workbook");
    let evaluator = FormulaEvaluator::new(&wb);

    assert_int(
        evaluator
            .evaluate_cell("Sheet1", 1, 4)
            .await
            .expect("VLOOKUP"),
        10,
    );
    assert_int(
        evaluator
            .evaluate_cell("Sheet1", 2, 4)
            .await
            .expect("INDEX/MATCH"),
        20,
    );
    assert_int(
        evaluator
            .evaluate_cell("Sheet1", 3, 4)
            .await
            .expect("XLOOKUP"),
        30,
    );

    assert_text(
        evaluator
            .evaluate_cell("Sheet1", 1, 6)
            .await
            .expect("CONCAT"),
        "Hello World",
    );
    assert_text(
        evaluator.evaluate_cell("Sheet1", 2, 6).await.expect("LEFT"),
        "Sam",
    );
    assert_int(
        evaluator.evaluate_cell("Sheet1", 3, 6).await.expect("FIND"),
        2,
    );
}

fn build_workbook(path: &str) {
    let mut wb = XlsxWorkbook::create().expect("create workbook");
    wb.add_worksheet("Sheet1");

    {
        let ws = wb.worksheet_mut(0).expect("worksheet 0");
        ws.set_name("Sheet1".to_string());

        // Lookup table
        ws.set_cell_value(1, 1, "Apples");
        ws.set_cell_value(2, 1, "Pears");
        ws.set_cell_value(3, 1, "Bananas");

        ws.set_cell_value(1, 2, 10);
        ws.set_cell_value(2, 2, 20);
        ws.set_cell_value(3, 2, 30);

        // Lookup formulas
        ws.set_cell_formula(1, 4, "VLOOKUP(\"Apples\", A1:B3, 2, FALSE)");
        ws.set_cell_formula(2, 4, "INDEX(B1:B3, MATCH(\"Pears\", A1:A3, 0))");
        ws.set_cell_formula(3, 4, "XLOOKUP(\"Bananas\", A1:A3, B1:B3)");

        // Text formulas
        ws.set_cell_formula(1, 6, "CONCAT(\"Hello\", \" \", \"World\")");
        ws.set_cell_formula(2, 6, "LEFT(\"Sample\", 3)");
        ws.set_cell_formula(3, 6, "FIND(\"am\", \"Sample\")");
    }

    wb.save(path).expect("save workbook");
}

fn assert_int(value: CellValue, expected: i64) {
    match value {
        CellValue::Int(v) => assert_eq!(v, expected),
        CellValue::Float(v) => assert!(
            (v - expected as f64).abs() < 1e-9,
            "expected {expected}, got {v}"
        ),
        other => panic!("Unexpected int value: {:?}", other),
    }
}

fn assert_text(value: CellValue, expected: &str) {
    match value {
        CellValue::String(s) => assert_eq!(s, expected),
        other => panic!("Unexpected text value: {:?}", other),
    }
}
