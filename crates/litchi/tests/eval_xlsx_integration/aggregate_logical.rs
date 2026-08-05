use litchi::ooxml::xlsx::{Formula, Workbook};
use litchi::sheet::{CellValue, FormulaEvaluator, functions::open_workbook};
use tempfile::tempdir;

const TOL: f64 = 1e-9;

#[tokio::test]
async fn eval_aggregates_and_logical() {
    let dir = tempdir().expect("create temp dir");
    let path = dir.path().join("aggregate_logical.xlsx");
    let path_str = path.to_str().expect("utf-8 path");

    build_workbook(path_str);

    let wb = open_workbook(path_str).expect("open workbook");
    let evaluator = FormulaEvaluator::new(wb.as_ref());

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
    let wb = Workbook::create().expect("create workbook");
    let mut edit = wb.edit().expect("start workbook edit");
    edit.tab(0)
        .expect("worksheet tab lookup")
        .expect("worksheet tab")
        .rename("Sheet1")
        .expect("rename worksheet");

    {
        let mut ws = edit
            .sheet(0)
            .expect("worksheet lookup")
            .expect("worksheet 0");
        // Data for aggregates
        ws.set("A1", 1).expect("A1");
        ws.set("A2", 2).expect("A2");
        ws.set("A3", 3).expect("A3");
        ws.set("A4", 4).expect("A4");

        // Aggregate formulas
        ws.set("D1", Formula::new("SUM(A1:A4)").expect("SUM formula"))
            .expect("D1");
        ws.set(
            "D2",
            Formula::new("AVERAGE(A1:A4)").expect("AVERAGE formula"),
        )
        .expect("D2");
        ws.set("D3", Formula::new("COUNT(A1:A4)").expect("COUNT formula"))
            .expect("D3");
        ws.set("D4", Formula::new("MAX(A1:A4)").expect("MAX formula"))
            .expect("D4");
        ws.set("D5", Formula::new("MIN(A1:A4)").expect("MIN formula"))
            .expect("D5");

        // Logical formulas
        ws.set(
            "F1",
            Formula::new("AND(TRUE, 1=1, A1=1)").expect("AND formula"),
        )
        .expect("F1");
        ws.set(
            "F2",
            Formula::new("OR(FALSE, 2=2, A4=5)").expect("OR formula"),
        )
        .expect("F2");
        ws.set(
            "F3",
            Formula::new("NOT(OR(FALSE, FALSE))").expect("NOT formula"),
        )
        .expect("F3");
        ws.set(
            "F4",
            Formula::new("IF(AND(A1=1, A4=4), \"yes\", \"no\")").expect("IF formula"),
        )
        .expect("F4");
    }

    edit.commit()
        .expect("commit workbook edit")
        .workbook()
        .save(path)
        .expect("save workbook");
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
