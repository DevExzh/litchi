use std::io::Cursor;

use litchi_ole::xls::{
    XlsPrintComments, XlsPrintErrors, XlsPrintOrder, XlsPrintOrientation,
};
use litchi_ole::xls::writer::{XlsPageSetupOptions, XlsWriter};
use litchi_ole::xls::XlsWorkbook;

#[test]
fn page_settings_round_trip_with_breaks_and_continued_pls() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Print").unwrap();
    let options = XlsPageSetupOptions {
        print_headers: true,
        print_gridlines: true,
        header: "&L报告&Cpage &P".to_string(),
        footer: "&Rconfidential".to_string(),
        horizontally_centered: true,
        vertically_centered: false,
        left_margin_inches: 0.6,
        right_margin_inches: 0.7,
        top_margin_inches: 0.8,
        bottom_margin_inches: 0.9,
        paper_size: 9,
        scale_percent: 85,
        starting_page_number: Some(3),
        fit_width_pages: 2,
        fit_height_pages: 3,
        print_order: XlsPrintOrder::OverThenDown,
        orientation: Some(XlsPrintOrientation::Landscape),
        black_and_white: true,
        draft_quality: true,
        comments: XlsPrintComments::AtEnd,
        errors: XlsPrintErrors::Dashes,
        horizontal_resolution_dpi: 300,
        vertical_resolution_dpi: 600,
        header_margin_inches: 0.3,
        footer_margin_inches: 0.4,
        copies: 2,
        printer_driver_data: Some(vec![0x5a; 9000]),
    };
    writer.set_page_setup(sheet, options).unwrap();
    writer.add_horizontal_page_break(sheet, 20, 0, 10).unwrap();
    writer.add_horizontal_page_break(sheet, 10, 0, 10).unwrap();
    writer.add_vertical_page_break(sheet, 5, 0, 100).unwrap();

    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let page = workbook.xls_worksheet(0).unwrap().page_setup().unwrap();

    assert!(page.print_headers());
    assert!(page.print_gridlines());
    assert_eq!(page.header(), "&L报告&Cpage &P");
    assert!(page.is_horizontally_centered());
    assert_eq!(page.horizontal_page_breaks().len(), 2);
    assert_eq!(page.horizontal_page_breaks()[0].position(), 10);
    assert_eq!(page.vertical_page_breaks()[0].position(), 5);
    assert_eq!(page.printer_driver_data()[0], vec![0x5a; 9000]);
    assert_eq!(page.print_setup().paper_size(), Some(9));
    assert_eq!(page.print_setup().scale_percent(), Some(85));
    assert_eq!(page.print_setup().starting_page_number(), Some(3));
    assert_eq!(page.print_setup().print_order(), XlsPrintOrder::OverThenDown);
    assert_eq!(page.print_setup().orientation(), Some(XlsPrintOrientation::Landscape));
    assert_eq!(page.print_setup().comments(), XlsPrintComments::AtEnd);
    assert_eq!(page.print_setup().errors(), XlsPrintErrors::Dashes);
}

#[test]
fn writer_rejects_invalid_dimensions_and_breaks() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Print").unwrap();
    let invalid = XlsPageSetupOptions {
        left_margin_inches: f64::NAN,
        ..XlsPageSetupOptions::default()
    };
    assert!(writer.set_page_setup(sheet, invalid).is_err());
    assert!(writer.add_horizontal_page_break(sheet, 1, 4, 4).is_err());
    assert!(writer.add_vertical_page_break(sheet, 256, 0, 10).is_err());
}
