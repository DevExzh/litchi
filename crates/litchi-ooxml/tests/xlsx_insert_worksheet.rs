use litchi_core::sheet::traits::WorkbookTrait;
use litchi_ooxml::xlsx::Workbook;
use litchi_opc::{OpcPackage, PackURI};

#[test]
fn inserted_worksheet_keeps_workbook_order_after_round_trip() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("insert-worksheet.xlsx");

    let mut workbook = Workbook::create().unwrap();
    workbook.add_worksheet("Sheet2");
    workbook.add_worksheet("Sheet3");
    workbook
        .insert_worksheet(1, "Inserted")
        .unwrap()
        .set_cell_value(1, 1, "marker");
    workbook.insert_worksheet(0, "First").unwrap();
    assert!(workbook.insert_worksheet(7, "Nope").is_err());
    workbook.save(&path).unwrap();

    // workbook.xml lists the sheets in the requested order.
    let package = OpcPackage::open(&path).unwrap();
    let workbook_part = package
        .get_part(&PackURI::new("/xl/workbook.xml").unwrap())
        .unwrap();
    let xml = std::str::from_utf8(workbook_part.blob()).unwrap();
    let positions = [
        xml.find("First").unwrap(),
        xml.find("Sheet1").unwrap(),
        xml.find("Inserted").unwrap(),
        xml.find("Sheet2").unwrap(),
        xml.find("Sheet3").unwrap(),
    ];
    assert!(
        positions.windows(2).all(|pair| pair[0] < pair[1]),
        "sheet order in workbook.xml: {xml}"
    );

    let reopened = Workbook::open(&path).unwrap();
    assert_eq!(reopened.worksheet_count(), 5);
    let mut names = Vec::new();
    let mut worksheets = reopened.worksheets();
    while let Some(worksheet) = worksheets.next() {
        names.push(worksheet.unwrap().name().to_string());
    }
    assert_eq!(names, ["First", "Sheet1", "Inserted", "Sheet2", "Sheet3"]);

    // The inserted sheet keeps its own content at its new position.
    let inserted = reopened.worksheet_by_index(2).unwrap();
    let marker = inserted.cell_value(1, 1).unwrap();
    assert_eq!(
        marker.as_ref(),
        &litchi_core::sheet::CellValue::String("marker".into())
    );
}
