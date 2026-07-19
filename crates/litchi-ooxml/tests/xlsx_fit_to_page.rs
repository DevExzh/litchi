use litchi_core::sheet::traits::WorkbookTrait;
use litchi_ooxml::xlsx::{Workbook, parse_worksheet_sheet_properties};
use litchi_opc::{OpcPackage, PackURI};

#[test]
fn fit_to_one_by_n_pages_survives_package_save_and_reopen() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("fit-to-page.xlsx");

    let mut workbook = Workbook::create().unwrap();
    let worksheet = workbook.worksheet_mut(0).unwrap();
    worksheet
        .set_page_setup_with_options("landscape", 9, None, Some(1), Some(0))
        .unwrap();
    worksheet.set_automatic_page_breaks(false);
    workbook.save(&path).unwrap();

    let package = OpcPackage::open(&path).unwrap();
    let worksheet_part = package
        .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
        .unwrap();
    let xml = std::str::from_utf8(worksheet_part.blob()).unwrap();
    assert!(xml.contains(r#"<pageSetUpPr autoPageBreaks="0" fitToPage="1"/>"#));
    assert!(xml.contains(r#"fitToWidth="1" fitToHeight="0""#));

    let properties = parse_worksheet_sheet_properties(worksheet_part.blob())
        .unwrap()
        .unwrap();
    let page_setup = properties.page_setup_properties().unwrap();
    assert!(!page_setup.automatic_page_breaks());
    assert!(page_setup.fit_to_page());

    let reopened = Workbook::open(&path).unwrap();
    assert_eq!(reopened.worksheet_count(), 1);
}
