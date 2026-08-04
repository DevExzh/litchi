use litchi_core::sheet::traits::WorkbookTrait;
use litchi_ooxml::xlsx::{
    Workbook,
    page_setup::{
        Comments, Copies, Dpi, ErrorMode, FirstPage, Fit, Measure, Order, Orientation, Paper,
        Scale, Setup, Unit,
    },
    parse_sheet_properties, parse_worksheet_page_setup,
};
use litchi_opc::{OpcPackage, PackURI};

fn saved_page_setup(path: &std::path::Path) -> Setup {
    let package = OpcPackage::open(path).unwrap();
    let worksheet_part = package
        .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
        .unwrap();
    parse_worksheet_page_setup(worksheet_part.blob())
        .unwrap()
        .unwrap()
}

#[test]
fn fit_to_one_by_n_pages_survives_package_save_and_reopen() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("fit-to-page.xlsx");

    let mut workbook = Workbook::create().unwrap();
    let worksheet = workbook.worksheet_mut(0).unwrap();
    worksheet.set_page(Setup {
        orientation: Some(Orientation::Landscape),
        paper: Some(Paper::A4),
        ..Setup::default()
    });
    worksheet.set_fit(Fit::ONE, Fit::NONE);
    worksheet.set_automatic_page_breaks(false);
    workbook.save(&path).unwrap();

    let package = OpcPackage::open(&path).unwrap();
    let worksheet_part = package
        .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
        .unwrap();
    let xml = std::str::from_utf8(worksheet_part.blob()).unwrap();
    assert!(xml.contains(r#"<pageSetUpPr autoPageBreaks="0" fitToPage="1"/>"#));
    assert!(xml.contains(r#"fitToWidth="1" fitToHeight="0""#));

    let properties = parse_sheet_properties(worksheet_part.blob())
        .unwrap()
        .unwrap();
    let page_setup = properties.page_setup_properties().unwrap();
    assert!(!page_setup.automatic_page_breaks());
    assert!(page_setup.fit_to_page());

    let reopened = Workbook::open(&path).unwrap();
    assert_eq!(reopened.worksheet_count(), 1);
}

#[test]
fn complete_typed_page_setup_survives_two_package_publications() {
    let directory = tempfile::tempdir().unwrap();
    let first_path = directory.path().join("page-setup-first.xlsx");
    let second_path = directory.path().join("page-setup-second.xlsx");
    let original = Setup {
        paper: Some(Paper::new(u32::MAX).unwrap()),
        paper_width: Some(Measure::new("00.50", Unit::Centimeter).unwrap()),
        paper_height: Some(Measure::new("297", Unit::Millimeter).unwrap()),
        scale: Some(Scale::AUTO),
        first_page: Some(FirstPage::new(-32_767).unwrap()),
        fit_to_width: Some(Fit::NONE),
        fit_to_height: Some(Fit::new(Fit::MAX).unwrap()),
        order: Some(Order::OverThenDown),
        orientation: Some(Orientation::Default),
        use_printer_defaults: Some(false),
        black_and_white: Some(true),
        draft: Some(false),
        comments: Some(Comments::AtEnd),
        use_first_page: Some(true),
        errors: Some(ErrorMode::NotAvailable),
        horizontal_dpi: Some(Dpi::new(1).unwrap()),
        vertical_dpi: Some(Dpi::new(u32::MAX).unwrap()),
        copies: Some(Copies::new(Copies::MAX).unwrap()),
    };

    let mut first = Workbook::create().unwrap();
    first.worksheet_mut(0).unwrap().set_page(original.clone());
    first.save(&first_path).unwrap();
    let parsed = saved_page_setup(&first_path);
    assert_eq!(parsed, original);

    let mut second = Workbook::create().unwrap();
    second.worksheet_mut(0).unwrap().set_page(parsed.clone());
    second.save(&second_path).unwrap();
    assert_eq!(saved_page_setup(&second_path), parsed);
}
