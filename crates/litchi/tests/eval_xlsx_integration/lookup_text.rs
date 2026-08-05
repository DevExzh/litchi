use litchi::ooxml::xlsx::{Formula, Workbook};
use litchi::sheet::{CellValue, FormulaEvaluator, functions::open_workbook};
use tempfile::tempdir;

#[tokio::test]
async fn eval_lookup_and_text_functions() {
    let dir = tempdir().expect("create temp dir");
    let path = dir.path().join("lookup_text.xlsx");
    let path_str = path.to_str().expect("utf-8 path");

    build_workbook(path_str);

    let wb = open_workbook(path_str).expect("open workbook");
    let evaluator = FormulaEvaluator::new(wb.as_ref());

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
        // Lookup table
        ws.set("A1", "Apples").expect("A1");
        ws.set("A2", "Pears").expect("A2");
        ws.set("A3", "Bananas").expect("A3");

        ws.set("B1", 10).expect("B1");
        ws.set("B2", 20).expect("B2");
        ws.set("B3", 30).expect("B3");

        // Lookup formulas
        ws.set(
            "D1",
            Formula::new("VLOOKUP(\"Apples\", A1:B3, 2, FALSE)").expect("VLOOKUP formula"),
        )
        .expect("D1");
        ws.set(
            "D2",
            Formula::new("INDEX(B1:B3, MATCH(\"Pears\", A1:A3, 0))").expect("INDEX/MATCH formula"),
        )
        .expect("D2");
        ws.set(
            "D3",
            Formula::new("XLOOKUP(\"Bananas\", A1:A3, B1:B3)").expect("XLOOKUP formula"),
        )
        .expect("D3");

        // Text formulas
        ws.set(
            "F1",
            Formula::new("CONCAT(\"Hello\", \" \", \"World\")").expect("CONCAT formula"),
        )
        .expect("F1");
        ws.set(
            "F2",
            Formula::new("LEFT(\"Sample\", 3)").expect("LEFT formula"),
        )
        .expect("F2");
        ws.set(
            "F3",
            Formula::new("FIND(\"am\", \"Sample\")").expect("FIND formula"),
        )
        .expect("F3");
    }

    edit.commit()
        .expect("commit workbook edit")
        .workbook()
        .save(path)
        .expect("save workbook");
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
