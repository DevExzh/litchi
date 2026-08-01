//! Round-trip tests for the BIFF8 HeaderFooter record (even/first pages).

use litchi_ole::xls::writer::{XlsPageSetupOptions, XlsWriter};
use litchi_ole::xls::{XlsHeaderFooter, XlsWorkbook};
use std::io::Cursor;

fn written_workbook(options: XlsPageSetupOptions) -> Vec<u8> {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("HF").unwrap();
    writer.write_string(sheet, 0, 0, "content").unwrap();
    writer.set_page_setup(sheet, options).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn even_and_first_pages_round_trip() {
    let mut header_footer = XlsHeaderFooter::default();
    header_footer
        .set_even("&LEven header".to_string(), "&REven footer".to_string())
        .unwrap();
    header_footer
        .set_first("&C第一页".to_string(), String::new())
        .unwrap();
    header_footer.set_scale_with_doc(true);
    header_footer.set_align_margins(true);

    let options = XlsPageSetupOptions {
        header: "&COdd header".to_string(),
        footer: "&COdd footer".to_string(),
        header_footer: Some(header_footer.clone()),
        ..XlsPageSetupOptions::default()
    };
    let workbook = XlsWorkbook::new(Cursor::new(written_workbook(options))).unwrap();
    let page = workbook.xls_worksheet(0).unwrap().page_setup().unwrap();

    assert_eq!(page.header(), "&COdd header");
    assert_eq!(page.footer(), "&COdd footer");
    let parsed = page.header_footer().expect("HeaderFooter record present");
    assert_eq!(parsed, &header_footer);
}

#[test]
fn no_header_footer_record_by_default() {
    let workbook = XlsWorkbook::new(Cursor::new(
        written_workbook(XlsPageSetupOptions::default()),
    ))
    .unwrap();
    let page = workbook.xls_worksheet(0).unwrap().page_setup().unwrap();
    assert!(page.header_footer().is_none());
}
