//! Round-trip tests for the BIFF8 `WebPub` record (published Web pages).

use litchi_xls::writer::Writer;
use litchi_xls::{WebPageType, WebPub, WebPubRange, WebSourceType, Workbook};
use std::io::Cursor;

fn workbook_publication() -> WebPub {
    WebPub {
        source: WebSourceType::Workbook,
        page_type: WebPageType::WorkbookFunctionality,
        range: None,
        auto_republish: true,
        single_file: true,
        style_id: 0x1122_3344,
        source_name: None,
        file_destination: "https://example.com/report.mht".to_string(),
        div_id: "top".to_string(),
        title: "Quarterly report".to_string(),
        chart_shape_id: None,
        reserved: Vec::new(),
    }
}

fn range_publication() -> WebPub {
    WebPub {
        source: WebSourceType::Range,
        page_type: WebPageType::ViewOnly,
        range: Some(WebPubRange::new(1, 9, 2, 5).unwrap()),
        auto_republish: false,
        single_file: false,
        style_id: 7,
        source_name: None,
        file_destination: "C:\\pub\\range.htm".to_string(),
        div_id: String::new(),
        title: "Range".to_string(),
        chart_shape_id: None,
        reserved: vec![0xAA, 0xBB],
    }
}

#[test]
fn workbook_web_pub_round_trips_through_writer_and_reader() {
    let publication = workbook_publication();

    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Report").unwrap();
    writer.write_string(sheet, 0, 0, "content").unwrap();
    writer.add_web_publication(publication.clone()).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();

    let workbook = Workbook::new(Cursor::new(output.into_inner())).unwrap();
    assert_eq!(workbook.web_publications(), &[publication]);
}

#[test]
fn worksheet_web_pub_round_trips_through_writer_and_reader() {
    let publication = range_publication();

    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Data").unwrap();
    writer.write_number(sheet, 1, 2, 3.5).unwrap();
    writer
        .add_sheet_web_publication(sheet, publication.clone())
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();

    let workbook = Workbook::new(Cursor::new(output.into_inner())).unwrap();
    assert!(workbook.web_publications().is_empty());
    let worksheet = workbook.xls_worksheet(0).unwrap();
    assert_eq!(worksheet.web_publications(), &[publication]);
}

#[test]
fn worksheet_web_pub_rejects_unknown_sheet() {
    let mut writer = Writer::new();
    assert!(
        writer
            .add_sheet_web_publication(0, range_publication())
            .is_err()
    );
}

#[test]
fn workbook_without_web_pub_has_none() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Plain").unwrap();
    writer.write_string(sheet, 0, 0, "content").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();

    let workbook = Workbook::new(Cursor::new(output.into_inner())).unwrap();
    assert!(workbook.web_publications().is_empty());
    assert!(
        workbook
            .xls_worksheet(0)
            .unwrap()
            .web_publications()
            .is_empty()
    );
}
