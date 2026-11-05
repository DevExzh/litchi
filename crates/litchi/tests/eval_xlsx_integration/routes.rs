//! The evaluator reads the workbook view of every route `litchi::open*`
//! takes for an XLSX: the source-backed file and bytes, and the eager owner.

use litchi::sheet::{CellValue, FormulaEvaluator, WorkbookTrait};
use litchi::xlsx::{Formula, Package};
use litchi::{File, OpenOptions, SourceMode};
use tempfile::tempdir;

#[tokio::test]
async fn eval_reads_the_workbook_view_of_every_route() {
    let dir = tempdir().expect("create temp dir");
    let path = dir.path().join("routes.xlsx");

    let workbook = Package::new().expect("create workbook");
    let mut edit = workbook.edit().expect("start workbook edit");
    edit.tab(0)
        .expect("worksheet tab lookup")
        .expect("worksheet tab")
        .rename("Data")
        .expect("rename worksheet");
    {
        let mut sheet = edit
            .sheet(0)
            .expect("worksheet lookup")
            .expect("worksheet 0");
        sheet.set("A1", 2).expect("A1");
        sheet.set("A2", 3).expect("A2");
        sheet
            .set("B1", Formula::new("SUM(A1:A2)*A2").expect("formula"))
            .expect("B1");
    }
    edit.commit()
        .expect("commit workbook edit")
        .package()
        .save(&path)
        .expect("save workbook");

    let bytes = std::fs::read(&path).expect("read workbook");
    let eager = OpenOptions::new().source_mode(SourceMode::Eager);
    let files = [
        ("path", litchi::open(&path).expect("open path")),
        (
            "bytes",
            litchi::open_bytes(bytes.clone()).expect("open bytes"),
        ),
        (
            "eager",
            litchi::open_bytes_with(bytes, &eager).expect("open eager"),
        ),
    ];
    for (route, file) in &files {
        assert!(matches!(file, File::Xlsx(_)), "{route}: {file:?}");
        let workbook = file.workbook().expect("workbook view");
        assert_eq!(
            WorkbookTrait::worksheet_names(&workbook),
            ["Data"],
            "{route}"
        );
        let evaluator = FormulaEvaluator::new(&workbook);
        match evaluator
            .evaluate_cell("Data", 1, 2)
            .await
            .unwrap_or_else(|error| panic!("{route}: {error}"))
        {
            CellValue::Int(15) => {},
            CellValue::Float(value) if (value - 15.0).abs() < 1e-9 => {},
            other => panic!("{route}: unexpected B1 value {other:?}"),
        }
    }
}
