use litchi_core::sheet::traits::WorkbookTrait;
use litchi_ooxml::xlsx::Workbook;
use litchi_opc::{OpcPackage, PackURI};

#[test]
fn removed_worksheet_drops_part_and_remaps_defined_names() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("remove-worksheet.xlsx");

    let mut workbook = Workbook::create().unwrap();
    workbook.add_worksheet("Sheet2");
    workbook.add_worksheet("Sheet3");
    workbook.define_name("Global", "Sheet1!$A$1");
    workbook.define_name_local("Scoped2", "Sheet2!$A$1", 1);
    workbook.define_name_local("Scoped3", "Sheet3!$A$1", 2);
    workbook
        .worksheet_mut(2)
        .unwrap()
        .set_print_area("A1:B2");

    let removed = workbook.remove_worksheet(1).unwrap();
    assert_eq!(removed.name(), "Sheet2");
    workbook.save(&path).unwrap();

    let package = OpcPackage::open(&path).unwrap();
    // The removed sheet's part is gone; the later sheet keeps its own part.
    assert!(
        package
            .get_part(&PackURI::new("/xl/worksheets/sheet2.xml").unwrap())
            .is_err()
    );
    assert!(
        package
            .get_part(&PackURI::new("/xl/worksheets/sheet3.xml").unwrap())
            .is_ok()
    );

    let workbook_part = package
        .get_part(&PackURI::new("/xl/workbook.xml").unwrap())
        .unwrap();
    let xml = std::str::from_utf8(workbook_part.blob()).unwrap();
    assert!(xml.contains("Sheet1"));
    assert!(xml.contains("Sheet3"));
    assert!(!xml.contains("Sheet2"));
    // Names scoped to the removed sheet are dropped; later scopes shift.
    assert!(!xml.contains("Scoped2"));
    assert!(xml.contains(r#"<definedName name="Scoped3" localSheetId="1">"#));
    assert!(xml.contains(r#"<definedName name="Global">"#));
    // The regenerated print area tracks Sheet3's new workbook position.
    assert!(xml.contains(r#"_xlnm.Print_Area" localSheetId="1"#));

    let reopened = Workbook::open(&path).unwrap();
    assert_eq!(reopened.worksheet_count(), 2);
    let mut names = Vec::new();
    let mut worksheets = reopened.worksheets();
    while let Some(worksheet) = worksheets.next() {
        names.push(worksheet.unwrap().name().to_string());
    }
    assert_eq!(names, ["Sheet1", "Sheet3"]);
}

#[test]
fn remove_worksheet_requires_the_writer_data_model() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("opened.xlsx");

    let mut workbook = Workbook::create().unwrap();
    workbook.save(&path).unwrap();

    // A workbook that was only opened for reading tracks no mutable
    // worksheets, so removal fails instead of silently doing nothing.
    let mut reopened = Workbook::open(&path).unwrap();
    assert!(reopened.remove_worksheet(0).is_err());
}
