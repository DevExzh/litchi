use litchi_ooxml::xlsx::{ChartAnchor, Workbook, WorksheetChart};
use litchi_opc::{OpcPackage, PackURI};

fn bar_chart() -> WorksheetChart {
    WorksheetChart::bar_chart(
        "Sales",
        "Sheet1!$A$2:$A$3",
        "Sheet1!$B$2:$B$3",
        ChartAnchor::new(0, 0, 10, 15),
    )
    .unwrap()
}

#[test]
fn second_save_after_worksheet_removal_drops_the_stale_part() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("double-save.xlsx");

    let mut workbook = Workbook::create().unwrap();
    workbook.add_worksheet("Sheet2");
    workbook.add_worksheet("Sheet3");
    workbook.save(&path).unwrap();

    workbook.remove_worksheet(1).unwrap();
    workbook.save(&path).unwrap();

    // The second save must not carry the removed sheet's part forward.
    let package = OpcPackage::open(&path).unwrap();
    assert!(
        package
            .get_part(&PackURI::new("/xl/worksheets/sheet2.xml").unwrap())
            .is_err(),
        "stale worksheet part survived the second save"
    );
    assert!(
        package
            .get_part(&PackURI::new("/xl/worksheets/sheet3.xml").unwrap())
            .is_ok()
    );
}

#[test]
fn second_save_after_chartsheet_removal_drops_stale_parts() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("double-save-chartsheet.xlsx");

    let mut workbook = Workbook::create().unwrap();
    workbook.add_chart_sheet("Chart A", bar_chart()).unwrap();
    workbook.add_chart_sheet("Chart B", bar_chart()).unwrap();
    workbook.save(&path).unwrap();

    workbook.remove_chart_sheet(0).unwrap();
    workbook.save(&path).unwrap();

    let package = OpcPackage::open(&path).unwrap();
    for uri in [
        // Index-named chartsheet infrastructure from the first save.
        "/xl/chartsheets/sheet2.xml",
        "/xl/drawings/drawingChartsheet2.xml",
        // The removed chartsheet's hosted chart part.
        "/xl/charts/chart2_1.xml",
    ] {
        assert!(
            package.get_part(&PackURI::new(uri).unwrap()).is_err(),
            "stale part {uri} survived the second save"
        );
    }
    // The surviving chartsheet was re-enumerated to slot 1.
    assert!(
        package
            .get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
            .is_ok()
    );
    assert!(
        package
            .get_part(&PackURI::new("/xl/charts/chart3_1.xml").unwrap())
            .is_ok()
    );
}
