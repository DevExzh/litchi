use litchi_core::sheet::WorkbookTrait;
use litchi_ooxml::xlsx::{Chart, ChartAnchor, Workbook};
use litchi_opc::{OpcPackage, PackURI};

fn bar_chart() -> Chart {
    Chart::bar_chart(
        "Sales",
        "Sheet1!$A$2:$A$3",
        "Sheet1!$B$2:$B$3",
        ChartAnchor::new(0, 0, 10, 15),
    )
    .unwrap()
}

#[test]
fn removed_chartsheet_drops_parts_and_remaps_defined_names() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("remove-chartsheet.xlsx");

    let mut workbook = Workbook::create().unwrap();
    workbook
        .add_chart_sheet("Sales Chart", bar_chart())
        .unwrap();
    workbook.add_worksheet("Sheet2");
    workbook.define_name_local("ScopedChart", "Sheet1!$A$1", 1);
    workbook.define_name_local("ScopedSheet2", "Sheet2!$A$1", 2);

    let removed = workbook.remove_chart_sheet(0).unwrap();
    assert_eq!(removed.name(), "Sales Chart");
    // Nothing left to remove; the second call fails cleanly.
    assert!(workbook.remove_chart_sheet(0).is_err());
    workbook.save(&path).unwrap();

    let package = OpcPackage::open(&path).unwrap();
    for uri in [
        "/xl/chartsheets/sheet1.xml",
        "/xl/drawings/drawingChartsheet1.xml",
        "/xl/charts/chart2_1.xml",
    ] {
        assert!(
            package.get_part(&PackURI::new(uri).unwrap()).is_err(),
            "stale part {uri}"
        );
    }

    let workbook_part = package
        .get_part(&PackURI::new("/xl/workbook.xml").unwrap())
        .unwrap();
    let xml = std::str::from_utf8(workbook_part.blob()).unwrap();
    assert!(!xml.contains("Sales Chart"));
    assert!(!xml.contains("ScopedChart"));
    assert!(xml.contains(r#"<definedName name="ScopedSheet2" localSheetId="1">"#));

    let reopened = Workbook::open(&path).unwrap();
    assert_eq!(
        WorkbookTrait::worksheet_names(&reopened),
        ["Sheet1", "Sheet2"]
    );
}

#[test]
fn validated_worksheet_creation_round_trips() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("validated-sheets.xlsx");

    let mut workbook = Workbook::create().unwrap();
    workbook.try_add_worksheet("Summary").unwrap();
    assert!(workbook.try_add_worksheet("summary").is_err());
    assert!(workbook.try_add_worksheet("a/b").is_err());
    workbook.save(&path).unwrap();

    let reopened = Workbook::open(&path).unwrap();
    assert_eq!(
        WorkbookTrait::worksheet_names(&reopened),
        ["Sheet1", "Summary"]
    );
}
