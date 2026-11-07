//! Direct trait consumers and the public facade must agree on native Excel addresses.
#![cfg(all(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "eval"))]

use litchi_core::sheet::{CellValue, WorkbookTrait};
use litchi_eval::FormulaEvaluator;
use std::{io::Cursor, path::PathBuf};

fn path(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/excel")
        .join(name)
}

async fn form_grid<W: WorkbookTrait>(book: &W) -> Vec<Vec<CellValue>>
where
    litchi_core::sheet::Error: From<W::Error>,
{
    let evaluator = FormulaEvaluator::new(book);
    assert_eq!(
        evaluator.evaluate_cell("Sheet1", 2, 2).await.unwrap(),
        CellValue::Bool(true)
    );
    let grid = evaluator.evaluate_sheet("Sheet1").await.unwrap();
    assert_eq!(grid.len(), 12);
    assert!(grid.iter().all(|row| row.len() == 4));
    assert_eq!(grid[1][1], CellValue::Bool(true));
    assert_eq!(grid[11][1], CellValue::Error("#N/A".into()));
    grid
}

#[tokio::test]
async fn native_direct_evaluator_and_facade_share_one_based_cell_and_sheet_reads() {
    for extension in ["xls", "xlsx", "xlsb"] {
        let input = path(&format!("excel-form-controls.{extension}"));
        let bytes = std::fs::read(&input).unwrap();
        let direct = match extension {
            "xls" => {
                Some(form_grid(&litchi_xls::raw::Reader::new(Cursor::new(&bytes)).unwrap()).await)
            },
            "xlsb" => {
                Some(form_grid(&litchi_xlsb::Workbook::new(Cursor::new(&bytes)).unwrap()).await)
            },
            _ => None,
        };
        let mut prior = direct;
        for mode in [litchi::SourceMode::Auto, litchi::SourceMode::Eager] {
            let file =
                litchi::open_with(&input, &litchi::OpenOptions::new().source_mode(mode)).unwrap();
            let book = file.workbook().unwrap();
            let actual = form_grid(&book).await;
            if let Some(expected) = &prior {
                assert_eq!(&actual, expected);
            }
            prior = Some(actual);
            assert!(file.to_bytes().unwrap() == bytes);
        }
        assert!(std::fs::read(input).unwrap() == bytes);
    }
}

#[tokio::test]
async fn native_metadata_numeric_b2_reads_match_coordinate_reads_in_direct_evaluators() {
    use litchi_core::sheet::{Cell, Worksheet};
    async fn check<W: WorkbookTrait>(book: &W, expected: CellValue)
    where
        litchi_core::sheet::Error: From<W::Error>,
    {
        let sheet = book.worksheet_by_name("Print_015").unwrap();
        assert_eq!(sheet.cell_value(2, 2).unwrap().as_ref(), &expected);
        let cell = sheet.cell_by_coordinate("B2").unwrap();
        assert_eq!(cell.value(), &expected);
        assert_eq!((cell.row(), cell.column()), (2, 2));
        let evaluator = FormulaEvaluator::new(book);
        assert_eq!(
            evaluator.evaluate_cell("Print_015", 2, 2).await.unwrap(),
            expected
        );
    }
    let xls = std::fs::read(path("excel-metadata.xls")).unwrap();
    check(
        &litchi_xls::raw::Reader::new(Cursor::new(&xls)).unwrap(),
        CellValue::Float(42.0),
    )
    .await;
    let xlsb = std::fs::read(path("excel-metadata.xlsb")).unwrap();
    check(
        &litchi_xlsb::Workbook::new(Cursor::new(&xlsb)).unwrap(),
        CellValue::Int(42),
    )
    .await;
    assert!(std::fs::read(path("excel-metadata.xls")).unwrap() == xls);
    assert!(std::fs::read(path("excel-metadata.xlsb")).unwrap() == xlsb);
}
