use litchi::sheet::{CellValue, FormulaEvaluator, functions::open_workbook};
use litchi::xlsx::{Formula, Number, Workbook};
use tempfile::tempdir;

const TOL: f64 = 1e-9;

#[tokio::test]
async fn eval_pduration_and_rri() {
    let dir = tempdir().expect("create temp dir");
    let path = dir.path().join("financial.xlsx");
    let path_str = path.to_str().expect("utf-8 path");

    build_financial_workbook(path_str);

    let wb = open_workbook(path_str).expect("open workbook");
    let evaluator = FormulaEvaluator::new(wb.as_ref());

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
    let wb = Workbook::create().expect("create workbook");
    let mut edit = wb.edit().expect("start workbook edit");
    edit.tab(0)
        .expect("worksheet tab lookup")
        .expect("worksheet tab")
        .rename("Financial")
        .expect("rename worksheet");

    {
        let mut ws = edit
            .sheet(0)
            .expect("worksheet lookup")
            .expect("worksheet 0");
        // Base inputs for PDURATION and RRI
        ws.set("A1", Number::new("0.1").expect("A1 rate"))
            .expect("A1 rate"); // A1 rate
        ws.set("A2", 10_i32).expect("A2 nper"); // A2 nper
        ws.set("C1", 1000_i32).expect("C1 pv"); // C1 pv
        ws.set("D1", 2000_i32).expect("D1 fv"); // D1 fv

        // Formulas under test
        ws.set(
            "B1",
            Formula::new("PDURATION(A1, C1, D1)").expect("PDURATION formula"),
        )
        .expect("B1");
        ws.set("B2", Formula::new("RRI(A2, C1, D1)").expect("RRI formula"))
            .expect("B2");
    }

    edit.commit()
        .expect("commit workbook edit")
        .workbook()
        .save(path)
        .expect("save workbook");
}
